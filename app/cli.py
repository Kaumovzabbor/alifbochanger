"""Buyruq satri: alifbo_konverter.exe fayl.docx -o papka"""
from __future__ import annotations

import argparse
import sys

from . import APP_NAME, __version__
from .batch import ST_DONE, collect_files, format_report, run_batch, run_preview
from .converter import ConvertOptions
from .settings import Settings, setup_logging

_APOS = {"curly": "’", "modifier": "ʼ", "straight": "'"}


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="alifbo_konverter",
        description=f"{APP_NAME}: .docx fayllarni yangi o'zbek lotin alifbosiga o'tkazadi.",
    )
    p.add_argument("paths", nargs="+", help=".docx fayllar yoki papkalar")
    p.add_argument("-o", "--output", help="natija papkasi (standart: asl fayl yonida)")
    p.add_argument("--preview", action="store_true", help="faylni o'zgartirmasdan, o'zgaradigan so'zlarni ko'rsatish")
    p.add_argument("--limit", type=int, default=200, help="--preview: ko'rsatiladigan so'zlar soni (standart 200)")
    p.add_argument("--replace", action="store_true", help="asl fayl o'rniga yozish (asl fayl .bak nusxa bilan saqlanadi)")
    p.add_argument("--suffix", help="natija fayl nomi qo'shimchasi (standart: _yangi_alifbo)")
    p.add_argument("--apostrophe", choices=sorted(_APOS), help="tutuq belgisi: curly ’, modifier ʼ, straight '")
    p.add_argument("--all-languages", action="store_true", help="til belgisidan qat'i nazar hamma matnni o'zgartirish")
    p.add_argument("--cyrillic", action="store_true", help="kirill matnini ham yangi lotinga o'tkazish")
    p.add_argument("--no-recursive", action="store_true", help="papkalarning ichki papkalariga kirmaslik")
    p.add_argument("--version", action="version", version=f"{APP_NAME} {__version__}")
    return p


def _setup_stdout() -> None:
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def main(argv=None) -> int:
    _setup_stdout()
    setup_logging()
    args = build_parser().parse_args(argv)
    st = Settings()
    opt = ConvertOptions(
        apostrophe=_APOS[args.apostrophe] if args.apostrophe else st["apostrophe"],
        exceptions=dict(st.exceptions),
        stop_words=list(st.stop_list),
        cyrillic=args.cyrillic or bool(st["cyrillic"]),
        only_uzbek_lang=False if args.all_languages else bool(st["only_uzbek_lang"]),
    )
    suffix = args.suffix or st["suffix"]
    files = collect_files(args.paths, recursive=not args.no_recursive, skip_outputs_suffix=suffix)
    if not files:
        print("Hech qanday .docx fayl topilmadi.")
        return 2

    if args.preview:
        words, errors = run_preview(files, opt)
        for (old, new), n in sorted(words.items(), key=lambda kv: (-kv[1], kv[0][0]))[: max(1, args.limit)]:
            print(f"{old} → {new}\t{n}")
        print(f"Takrorlanmaydigan so'zlar: {len(words)}")
        for path, msg in errors:
            print(f"[XATO] {path.name}: {msg}")
        return 1 if errors else 0

    def progress(i, total, fr):
        if fr.status == ST_DONE:
            print(f"[OK]   {fr.path.name} → {fr.output.name} ({fr.message})")
        elif fr.status == "error":
            print(f"[XATO] {fr.path.name}: {fr.message}")

    res = run_batch(files, opt, output_dir=args.output, suffix=suffix,
                    replace_original=args.replace or bool(st["replace_original"]), progress=progress)
    print()
    print(format_report(res))
    return 0 if res.failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
