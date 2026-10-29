"""Sof alifbo mantig'i: eski o'zbek lotin alifbosi -> yangi o'zbek lotin alifbosi.

Bu modul .docx haqida hech narsa bilmaydi.

Asosiy funksiyalar:
    convert_spans(text, options) -> list[Span]   (qaysi belgilar nima bilan almashadi)
    convert_text(text, options)  -> str          (tayyor natija)
    changed_words(text, spans)   -> [(eski, yangi)] (ko'rib chiqish oynasi uchun)

Natijada faqat quyidagi aniq (NFC) belgilar ishlatiladi:
    O-umlaut U+00D6/U+00F6, G-breve U+011E/U+011F,
    S-cedilla U+015E/U+015F, C-cedilla U+00C7/U+00E7.
"""
from __future__ import annotations

import bisect
import difflib
import re
from dataclasses import dataclass, field
from typing import Iterable, NamedTuple

# --- Belgilar -----------------------------------------------------------------

O_LOW, O_UP = "\u00f6", "\u00d6"
G_LOW, G_UP = "\u011f", "\u011e"
SH_LOW, SH_UP = "\u015f", "\u015e"
CH_LOW, CH_UP = "\u00e7", "\u00c7"

# ' U+0027, U+2018, U+2019, U+02BC, U+02BB, ` U+0060, U+00B4, U+2032
APOSTROPHES = "'\u2018\u2019\u02bc\u02bb`\u00b4\u2032"
_APOS_SET = frozenset(APOSTROPHES)
STANDARD_APOSTROPHES = ("\u2019", "\u02bc", "'")
DEFAULT_APOSTROPHE = "\u2019"
_APOS_TO_STD = str.maketrans({c: "\u2019" for c in APOSTROPHES})

_AP_CLASS = re.escape(APOSTROPHES)


class Span(NamedTuple):
    """text[start:end] o'rniga `repl` yoziladi. kind: o, g, sh, ch, apos, exc, cyr."""
    start: int
    end: int
    repl: str
    kind: str


def _norm_key(word: str) -> str:
    return word.strip().lower().translate(_APOS_TO_STD)


@dataclass
class ConvertOptions:
    apostrophe: str = DEFAULT_APOSTROPHE
    exceptions: dict = field(default_factory=dict)   # {"so'z": "natija"}
    stop_words: Iterable[str] = ()                   # o'zgartirilmaydigan so'zlar
    cyrillic: bool = False                           # kirill -> yangi lotin (qo'shimcha)
    only_uzbek_lang: bool = True                     # w:lang boshqa til bo'lsa o'tkazib yuborish
    _stop: frozenset = field(default=frozenset(), init=False, repr=False)
    _exc: dict = field(default_factory=dict, init=False, repr=False)

    def __post_init__(self):
        if self.apostrophe not in STANDARD_APOSTROPHES:
            raise ValueError("Noto'g'ri tutuq belgisi")
        self._stop = frozenset(_norm_key(w) for w in self.stop_words if w and w.strip())
        self._exc = {
            _norm_key(k): str(v)
            for k, v in dict(self.exceptions).items()
            if k and k.strip() and v is not None
        }


DEFAULT_OPTIONS = ConvertOptions()

# --- Regexlar -----------------------------------------------------------------

# 1) o'/g' juftligi, 2) sh/ch, 3) yolg'iz tutuq belgisi. Qo'shni harflar tekshiruvi kodda.
_MAIN_RE = re.compile(
    rf"(?P<ap>[oOgG])(?P<apc>[{_AP_CLASS}])"
    r"|(?P<dg>[sScC][hH])"
    rf"|(?P<tut>[{_AP_CLASS}])"
)

_WORD_RE = re.compile(rf"[^\W_]+(?:[{_AP_CLASS}][^\W_]+)*")

_TLDS = r"(?:uz|com|org|net|ru|edu|gov|info|io|kz)"
_PROTECT_RE = re.compile(
    r"(?i:(?:https?|ftp)://\S*|\bwww\.\S+|\bmailto:\S+)"            # URL
    r"|[\w.+\-]+@[\w\-]+(?:\.[\w\-]+)+"                              # e-pochta
    r"|(?<![\w])[A-Za-z]:\\\S*"                                      # C:\yo'l
    r"|\\\\[\w.$\-]+\\\S*"                                           # \\server\ulush
    r"|(?<!\S)~?/[\w.\-]+(?:/[\w.\-]+)+"                             # /yo'l/fayl
    r"|(?i:\b[\w\-]+(?:\.[\w\-]+)*\.(?:docx?|xlsx?|pptx?|pdf|txt|png|jpe?g|gif|zip|rar|exe|py|html?|csv|json|xml)\b)"
    rf"|(?i:\b[\w\-]+(?:\.[\w\-]+)*\.{_TLDS}\b)"                     # domen
)


def _is_letter(ch: str) -> bool:
    return ch.isalpha() and ch not in _APOS_SET


# --- Yordamchi: bosh harf mosligi (istisnolar uchun) ---------------------------

def _adapt_case(token: str, result: str) -> str:
    letters = [c for c in token if c.isalpha()]
    if len(letters) > 1 and token.isupper():
        return result.upper()
    if token[:1].isupper() and (len(token) == 1 or token[1:].islower()):
        return result[:1].upper() + result[1:]
    if token.islower():
        return result
    if len(token) == len(result):
        return "".join(r.upper() if t.isupper() else r for t, r in zip(token, result))
    return result


def _diff_spans(old: str, new: str, offset: int, kind: str) -> list[Span]:
    out: list[Span] = []
    sm = difflib.SequenceMatcher(None, old, new, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        out.append(Span(offset + i1, offset + i2, new[j1:j2], kind))
    return out


# --- Asosiy hisoblash ---------------------------------------------------------

def convert_spans(text: str, options: ConvertOptions | None = None) -> list[Span]:
    """text ichida qaysi oraliq nima bilan almashishini qaytaradi (kesishmaydi, tartiblangan)."""
    opt = options or DEFAULT_OPTIONS
    if not text:
        return []
    n = len(text)
    spans: list[Span] = []
    blocked = bytearray(n)

    def block(s: int, e: int) -> None:
        blocked[s:e] = b"\x01" * (e - s)

    for m in _PROTECT_RE.finditer(text):
        block(*m.span())

    if opt._stop or opt._exc:
        for m in _WORD_RE.finditer(text):
            key = _norm_key(m.group())
            if key in opt._stop:
                block(*m.span())
            elif key in opt._exc:
                block(*m.span())
                token = m.group()
                result = _adapt_case(token, opt._exc[key])
                if result != token:
                    spans.extend(_diff_spans(token, result, m.start(), "exc"))

    std = opt.apostrophe
    for m in _MAIN_RE.finditer(text):
        s, e = m.span()
        if blocked[s]:
            continue
        if m.group("ap"):
            nxt = text[e] if e < n else ""
            if not (nxt and _is_letter(nxt)):
                continue
            c = m.group("ap")
            repl = {"o": O_LOW, "O": O_UP, "g": G_LOW, "G": G_UP}[c]
            spans.append(Span(s, e, repl, "o" if c in "oO" else "g"))
        elif m.group("dg"):
            pair = m.group("dg")
            first, second = pair[0], pair[1]
            is_s = first in "sS"
            if first.isupper():
                upper = True
            elif second.isupper():
                prev_c = text[s - 1] if s > 0 else ""
                next_c = text[e] if e < n else ""
                upper = prev_c.isupper() or next_c.isupper()
            else:
                upper = False
            if is_s:
                repl = SH_UP if upper else SH_LOW
            else:
                repl = CH_UP if upper else CH_LOW
            spans.append(Span(s, e, repl, "sh" if is_s else "ch"))
        else:  # yolg'iz tutuq belgisi: ikki harf orasida bo'lsa standartlashtiramiz
            prev_c = text[s - 1] if s > 0 else ""
            next_c = text[e] if e < n else ""
            if prev_c and next_c and _is_letter(prev_c) and _is_letter(next_c):
                if m.group("tut") != std:
                    spans.append(Span(s, e, std, "apos"))

    if opt.cyrillic:
        from . import cyrillic  # kechiktirilgan import (aylanma import bo'lmasligi uchun)
        for sp in cyrillic.cyrillic_spans(text, std):
            if not blocked[sp.start]:
                spans.append(sp)

    spans.sort(key=lambda sp: (sp.start, sp.end))
    return spans


def apply_spans(text: str, spans: list[Span]) -> str:
    out: list[str] = []
    pos = 0
    for sp in spans:
        if sp.start < pos:  # kesishish bo'lmasligi kerak; xavfsizlik uchun
            continue
        out.append(text[pos:sp.start])
        out.append(sp.repl)
        pos = sp.end
    out.append(text[pos:])
    return "".join(out)


def convert_text(text: str, options: ConvertOptions | None = None) -> str:
    """Matnni yangi alifboga o'tkazadi (sof funksiya, idempotent)."""
    return apply_spans(text, convert_spans(text, options))


def changed_words(text: str, spans: list[Span]) -> list[tuple[str, str]]:
    """O'zgargan so'zlar ro'yxati: [(eski, yangi), ...] (ko'rib chiqish uchun)."""
    if not spans:
        return []
    starts = [sp.start for sp in spans]
    result: list[tuple[str, str]] = []
    for m in _WORD_RE.finditer(text):
        ts, te = m.span()
        lo = bisect.bisect_left(starts, ts)
        mine = []
        for sp in spans[lo:]:
            if sp.start > te:
                break
            if sp.start < te or (sp.start == te and sp.end == sp.start):
                mine.append(Span(sp.start - ts, sp.end - ts, sp.repl, sp.kind))
        if mine:
            old = m.group()
            new = apply_spans(old, mine)
            if new != old:
                result.append((old, new))
    return result
