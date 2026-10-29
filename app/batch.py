"""Ko'p fayllarni ketma-ket qayta ishlash (GUI va CLI uchun umumiy qatlam)."""
from __future__ import annotations

import logging
import os
import re
import threading
import traceback
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable

from .converter import ConvertOptions
from .docx_processor import (SUPPORTED_EXT, ConversionStats, DocxError, convert_file,
                             preview_docx)

log = logging.getLogger("alifbo")

MAX_FILE_BYTES = 500 * 1024 * 1024

ST_WAIT, ST_RUN, ST_DONE, ST_ERR = "wait", "run", "done", "error"


def _is_output_name(stem: str, suffix: str) -> bool:
    """Avvalgi natija fayli (…_yangi_alifbo yoki …_yangi_alifbo_2) ekanini aniqlaydi."""
    return stem.endswith(suffix) or re.fullmatch(r".*" + re.escape(suffix) + r"_\d+", stem) is not None


def collect_files(paths: Iterable, recursive: bool = True, skip_outputs_suffix: str = "_yangi_alifbo") -> list[Path]:
    """Fayl va papkalardan .docx ro'yxatini yig'adi (takrorsiz, tartibli).

    Papka ichida: Word'ning vaqtinchalik (~$...) fayllari va avvalgi natijalar (…_yangi_alifbo) o'tkazib yuboriladi.
    Foydalanuvchi aniq tanlagan fayl esa har doim qo'shiladi.
    """
    seen: set[str] = set()
    result: list[Path] = []

    def add(p: Path) -> None:
        key = os.path.normcase(str(p.resolve()))
        if key not in seen:
            seen.add(key)
            result.append(p)

    for raw in paths:
        p = Path(raw)
        if p.is_dir():
            it = p.rglob("*") if recursive else p.glob("*")
            for f in sorted(it):
                if not f.is_file() or f.suffix.lower() not in SUPPORTED_EXT:
                    continue
                if f.name.startswith("~$") or f.name.startswith("."):
                    continue
                if skip_outputs_suffix and _is_output_name(f.stem, skip_outputs_suffix):
                    continue
                add(f)
        elif p.exists():
            add(p)
        else:
            add(p)   # mavjud bo'lmasa ham ro'yxatga tushadi; keyin tushunarli xato beradi
    return result


@dataclass
class FileResult:
    path: Path
    status: str = ST_WAIT
    message: str = ""
    output: Path | None = None
    stats: ConversionStats = field(default_factory=ConversionStats)


@dataclass
class BatchResult:
    files: list[FileResult] = field(default_factory=list)
    cancelled: bool = False

    @property
    def ok(self) -> int:
        return sum(1 for f in self.files if f.status == ST_DONE)

    @property
    def failed(self) -> int:
        return sum(1 for f in self.files if f.status == ST_ERR)

    @property
    def counts(self) -> Counter:
        c: Counter = Counter()
        for f in self.files:
            if f.status == ST_DONE:
                c.update(f.stats.counts)
        return c


def format_report(res: BatchResult) -> str:
    c = res.counts
    lines = [
        f"Qayta ishlangan fayllar: {res.ok} ta" + (f", xato: {res.failed} ta" if res.failed else ""),
        f"oʻ → ö: {c['o']} ta",
        f"gʻ → ğ: {c['g']} ta",
        f"sh → ş: {c['sh']} ta",
        f"ch → ç: {c['ch']} ta",
    ]
    if c["apos"]:
        lines.append(f"tutuq belgisi standartlashtirildi: {c['apos']} ta")
    if c["exc"]:
        lines.append(f"istisnolar bo'yicha: {c['exc']} ta")
    if c["cyr"]:
        lines.append(f"kirill harflari: {c['cyr']} ta")
    return "\n".join(lines)


ProgressCB = Callable[[int, int, FileResult], None]


def run_batch(files: list[Path], options: ConvertOptions, output_dir=None, suffix: str = "_yangi_alifbo",
              replace_original: bool = False, progress: ProgressCB | None = None,
              cancel: threading.Event | None = None) -> BatchResult:
    """Har bir fayl alohida: bittasidagi xato qolganlariga halaqit bermaydi."""
    res = BatchResult(files=[FileResult(Path(f)) for f in files])
    total = len(res.files)
    for i, fr in enumerate(res.files):
        if cancel is not None and cancel.is_set():
            res.cancelled = True
            break
        fr.status = ST_RUN
        if progress:
            progress(i, total, fr)
        try:
            try:
                size = fr.path.stat().st_size
            except OSError:
                size = 0
            if size > MAX_FILE_BYTES:
                raise DocxError("Fayl juda katta (500 MB dan oshadi).")
            out, stats = convert_file(fr.path, options, output_dir=output_dir, suffix=suffix,
                                      replace_original=replace_original)
            fr.output, fr.stats, fr.status = out, stats, ST_DONE
            fr.message = f"{stats.total} ta almashtirish"
            log.info("OK %s -> %s: %s", fr.path.name, out.name, dict(stats.counts))
        except DocxError as e:
            fr.status, fr.message = ST_ERR, str(e)
            log.warning("XATO %s: %s", fr.path.name, e)
        except MemoryError:
            fr.status, fr.message = ST_ERR, "Xotira yetishmadi (fayl juda katta)."
            log.warning("XATO %s: MemoryError", fr.path.name)
        except Exception as e:   # kutilmagan: hujjat matnini logga yozmaymiz
            fr.status, fr.message = ST_ERR, f"Kutilmagan xato ({type(e).__name__}). Fayl o'tkazib yuborildi."
            log.error("XATO %s: %s\n%s", fr.path.name, type(e).__name__,
                      "".join(traceback.format_tb(e.__traceback__)))
        if progress:
            progress(i + 1, total, fr)
    return res


def run_preview(files: list[Path], options: ConvertOptions,
                progress: Callable[[int, int, Path], None] | None = None,
                cancel: threading.Event | None = None) -> tuple[Counter, list[tuple[Path, str]]]:
    """Fayllarni o'zgartirmasdan, o'zgaradigan so'zlar: Counter[(eski, yangi)]; xatolar ro'yxati."""
    words: Counter = Counter()
    errors: list[tuple[Path, str]] = []
    for i, f in enumerate(files):
        if cancel is not None and cancel.is_set():
            break
        if progress:
            progress(i, len(files), Path(f))
        try:
            w, _ = preview_docx(f, options)
            words.update(w)
        except DocxError as e:
            errors.append((Path(f), str(e)))
        except Exception as e:
            errors.append((Path(f), f"Kutilmagan xato ({type(e).__name__})."))
    if progress:
        progress(len(files), len(files), Path(""))
    return words, errors
