"""Interfeys matnlari: app/strings_uz.json dan olinadi (keyin osongina almashtirsa bo'ladi).

Foydalanuvchi o'z tarjimasini %APPDATA%\\AlifboKonverter\\strings.json faylida berishi mumkin:
u yerdagi kalitlar standart matnlarni almashtiradi.
"""
from __future__ import annotations

import json
from pathlib import Path

from .settings import data_dir, resource_path

_cache: dict[str, str] | None = None


def _load_file(p: Path) -> dict[str, str]:
    try:
        raw = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(raw, dict):
            return {str(k): str(v) for k, v in raw.items()}
    except (OSError, ValueError):
        pass
    return {}


def load(force: bool = False) -> dict[str, str]:
    global _cache
    if _cache is None or force:
        data = _load_file(resource_path("app/strings_uz.json"))
        data.update(_load_file(data_dir() / "strings.json"))
        _cache = data
    return _cache


def T(key: str, **kw) -> str:
    text = load().get(key, key)
    if kw:
        try:
            return text.format(**kw)
        except (KeyError, IndexError, ValueError):
            return text
    return text
