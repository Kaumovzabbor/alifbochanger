"""Ixtiyoriy rejim: kirill alifbosidagi o'zbek matnini yangi lotinga o'tkazish.

Asosiy lotin -> lotin rejimiga ta'sir qilmaydi (faqat kirill harflariga tegadi).
"""
from __future__ import annotations

from .converter import Span

_MAP = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d",
    "ж": "j", "з": "z", "и": "i", "й": "y", "к": "k",
    "л": "l", "м": "m", "н": "n", "о": "o", "п": "p",
    "р": "r", "с": "s", "т": "t", "у": "u", "ф": "f",
    "х": "x", "э": "e", "ы": "i",
    "ў": "ö",   # ў -> ö
    "қ": "q",        # қ
    "ғ": "ğ",   # ғ -> ğ
    "ҳ": "h",        # ҳ
    "ч": "ç",   # ч -> ç
    "ш": "ş",   # ш -> ş
    "ё": "yo", "ю": "yu", "я": "ya", "ц": "ts",
    "щ": "şç",  # щ -> şç
    "ь": "",         # ь tushib qoladi
}
_YE = "е"            # е
_HARD = "ъ"          # ъ -> tutuq belgisi
_VOWELS = set("аеёиоуўэюяы")


def _is_letter(ch: str) -> bool:
    return ch.isalpha()


def cyrillic_spans(text: str, apostrophe: str = "’") -> list[Span]:
    spans: list[Span] = []
    n = len(text)
    for i, ch in enumerate(text):
        low = ch.lower()
        if low == _HARD:
            spans.append(Span(i, i + 1, apostrophe, "cyr"))
            continue
        if low == _YE:
            prev = text[i - 1] if i > 0 else ""
            rep = "ye" if (not prev or not _is_letter(prev) or prev.lower() in _VOWELS) else "e"
        elif low in _MAP:
            rep = _MAP[low]
        else:
            continue
        if ch.isupper() and rep:
            if len(rep) == 1:
                rep = rep.upper()
            else:
                prev = text[i - 1] if i > 0 else ""
                nxt = text[i + 1] if i + 1 < n else ""
                full = (nxt.isalpha() and nxt.isupper()) or (prev.isalpha() and prev.isupper())
                rep = rep.upper() if full else rep[0].upper() + rep[1:]
        spans.append(Span(i, i + 1, rep, "cyr"))
    return spans


def cyrillic_to_new_latin(text: str, apostrophe: str = "’") -> str:
    """Faqat kirill harflarini almashtiradi."""
    out, pos = [], 0
    for sp in cyrillic_spans(text, apostrophe):
        out.append(text[pos:sp.start])
        out.append(sp.repl)
        pos = sp.end
    out.append(text[pos:])
    return "".join(out)
