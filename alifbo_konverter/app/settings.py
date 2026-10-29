"""Sozlamalar va istisnolar: %APPDATA%\\AlifboKonverter\\ ichida JSON sifatida saqlanadi."""
from __future__ import annotations

import json
import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

from .converter import ConvertOptions, DEFAULT_APOSTROPHE, STANDARD_APOSTROPHES

DEFAULTS = {
    "apostrophe": DEFAULT_APOSTROPHE,
    "only_uzbek_lang": True,
    "cyrillic": False,
    "suffix": "_yangi_alifbo",
    "replace_original": False,
    "theme": "light",          # "light" | "dark"
    "output_dir": "",
}


def resource_path(rel: str) -> Path:
    """Dastur papkasi yoki PyInstaller (_MEIPASS) ichidagi resurs yo'li."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / rel


def data_dir() -> Path:
    if os.name == "nt":
        root = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
    else:
        root = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
    d = root / "AlifboKonverter"
    try:
        d.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass
    return d


def setup_logging() -> logging.Logger:
    """Log faylga hujjat matni yozilmaydi: faqat fayl nomi va statistika."""
    log = logging.getLogger("alifbo")
    if log.handlers:
        return log
    log.setLevel(logging.INFO)
    try:
        h = RotatingFileHandler(data_dir() / "alifbo.log", maxBytes=500_000, backupCount=2, encoding="utf-8")
        h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        log.addHandler(h)
    except OSError:
        log.addHandler(logging.NullHandler())
    return log


class Settings:
    def __init__(self, directory: Path | None = None):
        self.dir = Path(directory) if directory else data_dir()
        self.path = self.dir / "settings.json"
        self.exceptions_path = self.dir / "exceptions.json"
        self.values = dict(DEFAULTS)
        self.exceptions: dict[str, str] = {}
        self.stop_list: list[str] = []
        self.load()

    # -- yuklash/saqlash --
    def load(self) -> None:
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                for k, default in DEFAULTS.items():
                    v = raw.get(k, default)
                    if isinstance(v, type(default)):
                        self.values[k] = v
        except (OSError, ValueError):
            pass
        if self.values["apostrophe"] not in STANDARD_APOSTROPHES:
            self.values["apostrophe"] = DEFAULT_APOSTROPHE
        if self.values["theme"] not in ("light", "dark"):
            self.values["theme"] = "light"
        self._load_exceptions()

    def _load_exceptions(self) -> None:
        candidates = [self.exceptions_path, resource_path("app/exceptions.json")]
        for p in candidates:
            try:
                raw = json.loads(p.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                continue
            if isinstance(raw, dict):
                exc = raw.get("exceptions", {})
                stop = raw.get("stop_list", [])
                if isinstance(exc, dict):
                    self.exceptions = {str(k): str(v) for k, v in exc.items()}
                if isinstance(stop, list):
                    self.stop_list = [str(x) for x in stop if str(x).strip()]
                return

    def save(self) -> None:
        try:
            self.dir.mkdir(parents=True, exist_ok=True)
            self.path.write_text(json.dumps(self.values, ensure_ascii=False, indent=2), encoding="utf-8")
            self.exceptions_path.write_text(
                json.dumps({"exceptions": self.exceptions, "stop_list": self.stop_list},
                           ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass

    # -- qulay kirish --
    def __getitem__(self, key):
        return self.values[key]

    def __setitem__(self, key, value):
        self.values[key] = value

    def add_exceptions(self, pairs: dict[str, str]) -> None:
        self.exceptions.update(pairs)
        self.save()

    def options(self) -> ConvertOptions:
        return ConvertOptions(
            apostrophe=self.values["apostrophe"],
            exceptions=dict(self.exceptions),
            stop_words=list(self.stop_list),
            cyrillic=bool(self.values["cyrillic"]),
            only_uzbek_lang=bool(self.values["only_uzbek_lang"]),
        )


def parse_exceptions_text(text: str) -> dict[str, str]:
    out: dict[str, str] = {}
    for line in text.splitlines():
        for sep in ("=>", "->", "→", "="):
            if sep in line:
                a, b = line.split(sep, 1)
                a, b = a.strip(), b.strip()
                if a and b:
                    out[a] = b
                break
    return out


def parse_stop_text(text: str) -> list[str]:
    seen, out = set(), []
    for line in text.splitlines():
        w = line.strip()
        if w and w.lower() not in seen:
            seen.add(w.lower())
            out.append(w)
    return out
