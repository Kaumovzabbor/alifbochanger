"""Kirish nuqtasi: argumentlar bo'lsa buyruq satri, bo'lmasa oyna."""
import os
import sys


def _attach_console_windows() -> None:
    """--windowed .exe da ham buyruq satri natijasi ko'rinishi uchun mavjud konsolga ulanamiz."""
    if os.name != "nt":
        return
    try:
        import ctypes
        if ctypes.windll.kernel32.AttachConsole(-1):
            sys.stdout = open("CONOUT$", "w", encoding="utf-8", errors="replace")
            sys.stderr = open("CONOUT$", "w", encoding="utf-8", errors="replace")
    except Exception:
        pass


def main() -> int:
    if len(sys.argv) > 1:
        if getattr(sys, "frozen", False):
            _attach_console_windows()
        from app.cli import main as cli_main
        return cli_main()
    from app.gui import run_gui
    return run_gui()


if __name__ == "__main__":
    sys.exit(main())
