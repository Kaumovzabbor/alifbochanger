"""Word (.docx) faylini zipfile + lxml bilan to'g'ridan-to'g'ri XML ustida qayta ishlash.

Asosiy g'oya: har bir paragrafdagi barcha matn bo'laklarini (run'larni) bitta satrga
birlashtiramiz, alifbo mantig'ini shu satrga qo'llaymiz, so'ng har bir almashtirishni
(span) tegishli run'larga qaytaramiz. Har bir run o'z formatini (w:rPr) saqlaydi.
"""
from __future__ import annotations

import logging
import os
import re
import shutil
import tempfile
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

from lxml import etree

from .converter import (ConvertOptions, DEFAULT_OPTIONS, apply_spans, changed_words,
                        convert_spans)

log = logging.getLogger("alifbo")

MAX_TOTAL_UNCOMPRESSED = 2 * 1024 ** 3   # zip-bomba himoyasi
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ENC_MARK = "EncryptedPackage".encode("utf-16-le")
SUPPORTED_EXT = (".docx", ".docm", ".dotx", ".dotm")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
DC = "http://purl.org/dc/elements/1.1/"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
XML_NS = "http://www.w3.org/XML/1998/namespace"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def _w(n): return f"{{{W}}}{n}"
def _a(n): return f"{{{A}}}{n}"


W_P, W_R, W_T, W_DELTEXT, W_RPR, W_LANG = _w("p"), _w("r"), _w("t"), _w("delText"), _w("rPr"), _w("lang")
W_TAB, W_PTAB, W_BR, W_CR = _w("tab"), _w("ptab"), _w("br"), _w("cr")
A_P, A_R, A_T, A_BR, A_FLD = _a("p"), _a("r"), _a("t"), _a("br"), _a("fld")
C_V = f"{{{C}}}v"
M_OMATH, M_OMATHPARA = f"{{{M}}}oMath", f"{{{M}}}oMathPara"
XML_SPACE = f"{{{XML_NS}}}space"

_CONTAINERS = {_w(n) for n in (
    "hyperlink", "ins", "del", "moveFrom", "moveTo", "smartTag", "sdt", "sdtContent",
    "fldSimple", "customXml", "bdo", "dir")}
_REVISION = {_w(n) for n in ("ins", "del", "moveFrom", "moveTo")}
_RUN_IGNORE = {_w(n) for n in ("lastRenderedPageBreak", "annotationRef")}
_AUTHOR = _w("author")

BOUNDARY = "\x0b"   # so'z chegarasi (matn belgisi emas)

# Matn bo'lagi bo'ladigan qismlar
_TEXT_CT_KEYS = (
    "wordprocessingml.document.main", "ms-word.document.macroEnabled.main",
    "wordprocessingml.template.main", "ms-word.template.macroEnabledTemplate.main",
    "wordprocessingml.header+xml", "wordprocessingml.footer+xml",
    "wordprocessingml.footnotes+xml", "wordprocessingml.endnotes+xml",
    "wordprocessingml.comments+xml", "drawingml.chart+xml",
    "drawingml.diagramData+xml", "drawingml.diagramDrawing+xml",
)
_TEXT_NAME_RE = re.compile(
    r"^word/(?:document\d*|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$"
    r"|^word/charts/chart[^/]*\.xml$"
    r"|^word/diagrams/(?:data|drawing)[^/]*\.xml$")
CORE_PROPS = "docProps/core.xml"

_DOCPR_TAGS = ("{*}docPr", "{*}cNvPr")


class DocxError(Exception):
    """Foydalanuvchiga ko'rsatiladigan tushunarli o'zbekcha xabar."""


@dataclass
class ConversionStats:
    counts: Counter = field(default_factory=Counter)   # o, g, sh, ch, apos, exc, cyr
    skipped_lang_runs: int = 0
    parts_changed: int = 0

    def add(self, other: "ConversionStats") -> None:
        self.counts.update(other.counts)
        self.skipped_lang_runs += other.skipped_lang_runs
        self.parts_changed += other.parts_changed

    @property
    def total(self) -> int:
        return sum(self.counts.values())


# --- Fayl tekshiruvi -------------------------------------------------------------

def check_file(path) -> zipfile.ZipFile:
    """Faylni ochadi va tekshiradi; muammo bo'lsa DocxError."""
    p = Path(path)
    if not p.is_file():
        raise DocxError("Fayl topilmadi.")
    try:
        with open(p, "rb") as f:
            head = f.read(8)
            if head.startswith(OLE_MAGIC):
                if p.suffix.lower() == ".doc":
                    raise DocxError("Bu eski .doc formati. Word'da faylni ochib, «Boshqacha saqlash» "
                                    "orqali .docx formatida saqlang, keyin qayta urinib ko'ring.")
                f.seek(0)
                if ENC_MARK in f.read():
                    raise DocxError("Fayl parol bilan himoyalangan. Parolni olib tashlab "
                                    "(Word: Fayl → Ma'lumot → Hujjatni himoyalash), qayta urinib ko'ring.")
                raise DocxError("Bu fayl .docx emas (eski .doc formati bo'lishi mumkin). "
                                "Word'da .docx qilib saqlang.")
    except PermissionError:
        raise DocxError("Faylni o'qib bo'lmadi: u boshqa dastur tomonidan band qilingan "
                        "yoki ruxsat yo'q.") from None
    except OSError:
        raise DocxError("Faylni o'qib bo'lmadi.") from None
    try:
        zf = zipfile.ZipFile(p, "r")
    except zipfile.BadZipFile:
        raise DocxError("Fayl buzilgan yoki .docx formatida emas.") from None
    except OSError:
        raise DocxError("Faylni o'qib bo'lmadi.") from None
    try:
        infos = zf.infolist()
        names = {i.filename for i in infos}
        if any(i.flag_bits & 0x1 for i in infos):
            raise DocxError("Fayl parol bilan himoyalangan.")
        if "[Content_Types].xml" not in names or not any(n.startswith("word/") for n in names):
            raise DocxError("Bu Word (.docx) fayli emas yoki buzilgan.")
        if sum(i.file_size for i in infos) > MAX_TOTAL_UNCOMPRESSED:
            raise DocxError("Fayl juda katta (ochilgan hajmi 2 GB dan oshadi).")
    except DocxError:
        zf.close()
        raise
    return zf


def _text_parts(zf: zipfile.ZipFile) -> set[str]:
    names = {i.filename for i in zf.infolist()}
    parts = {n for n in names if _TEXT_NAME_RE.match(n)}
    try:
        root = etree.fromstring(zf.read("[Content_Types].xml"), _parser())
        for ov in root.iter(f"{{{CT_NS}}}Override"):
            ct = ov.get("ContentType", "")
            pn = (ov.get("PartName") or "").lstrip("/")
            if pn in names and any(k in ct for k in _TEXT_CT_KEYS):
                parts.add(pn)
    except Exception:  # [Content_Types].xml buzilgan bo'lsa nom bo'yicha davom etamiz
        pass
    return parts


def _parser() -> etree.XMLParser:
    return etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False,
                           huge_tree=True, strip_cdata=False, remove_blank_text=False)


# --- Paragrafni qayta ishlash ----------------------------------------------------

_NOCTX = object()


class _Processor:
    def __init__(self, options: ConvertOptions, apply: bool = True, words: Counter | None = None):
        self.opt = options
        self.apply = apply
        self.words = words            # ko'rib chiqish rejimi: (eski, yangi) -> soni
        self.stats = ConversionStats()
        self.changed = False

    # -- yig'ish --
    def _run_skipped(self, r) -> bool:
        if not self.opt.only_uzbek_lang:
            return False
        rpr = r.find(W_RPR)
        if rpr is None:
            return False
        lang = rpr.find(W_LANG)
        if lang is None:
            return False
        val = (lang.get(_w("val")) or "").strip().lower()
        if not val:
            return False
        base = re.split(r"[-_]", val)[0]
        return base not in ("uz", "x", "und", "zxx", "mul", "")

    def _collect_w(self, p):
        items = []

        def handle_run(r, ctx):
            skip = self._run_skipped(r)
            if skip:
                self.stats.skipped_lang_runs += 1
            for ch in r:
                tag = ch.tag
                if not isinstance(tag, str) or tag == W_RPR or tag in _RUN_IGNORE:
                    continue
                if tag == W_T or tag == W_DELTEXT:
                    items.append(("t", ch, skip, ctx))
                elif tag == W_TAB or tag == W_PTAB:
                    items.append(("b", "\t"))
                elif tag == W_BR or tag == W_CR:
                    items.append(("b", "\n"))
                else:   # noBreakHyphen, sym, fldChar, instrText, drawing, pict, ...
                    items.append(("b", BOUNDARY))

        def walk(node, ctx):
            for ch in node:
                tag = ch.tag
                if not isinstance(tag, str):
                    continue
                if tag == W_R:
                    handle_run(ch, ctx)
                elif tag in _CONTAINERS:
                    nctx = ctx
                    if tag in _REVISION:
                        nctx = (tag, ch.get(_AUTHOR))
                    walk(ch, nctx)
                elif tag == M_OMATH or tag == M_OMATHPARA:
                    items.append(("b", BOUNDARY))

        walk(p, None)
        return items

    def _collect_a(self, p):
        items = []
        for ch in p:
            tag = ch.tag
            if not isinstance(tag, str):
                continue
            if tag == A_R:
                for t in ch:
                    if t.tag == A_T:
                        items.append(("t", t, False, None))
            elif tag == A_BR:
                items.append(("b", "\n"))
            elif tag == A_FLD:
                items.append(("b", BOUNDARY))
        return items

    # -- bitta paragraf --
    def paragraph(self, p) -> None:
        items = self._collect_w(p) if p.tag == W_P else self._collect_a(p)
        if not items:
            return
        chars: list[str] = []
        cmap: list = []          # har bir belgi uchun (element_indeksi, ofset) yoki None
        elems: list = []
        prev_ctx = _NOCTX
        for it in items:
            if it[0] == "b":
                chars.append(it[1])
                cmap.append(None)
                prev_ctx = _NOCTX
                continue
            _, el, skip, ctx = it
            txt = el.text or ""
            if not txt:
                continue
            if prev_ctx is not _NOCTX and ctx != prev_ctx:
                chars.append(BOUNDARY)
                cmap.append(None)
            prev_ctx = ctx
            idx = len(elems)
            elems.append(el)
            if skip:
                chars.append(BOUNDARY * len(txt))
            else:
                chars.append(txt)
            cmap.extend((idx, k) for k in range(len(txt)))
        if not elems:
            return
        text = "".join(chars)
        spans = convert_spans(text, self.opt)
        if not spans:
            return
        for sp in spans:
            self.stats.counts[sp.kind] += 1
        if self.words is not None:
            for pair in changed_words(text, spans):
                self.words[pair] += 1
        if not self.apply:
            return

        lists: dict[int, list[str]] = {}

        def get(idx):
            lst = lists.get(idx)
            if lst is None:
                lst = list(elems[idx].text)
                lists[idx] = lst
            return lst

        n = len(cmap)
        for sp in spans:
            if sp.end > sp.start:
                placed = False
                for i in range(sp.start, sp.end):
                    c = cmap[i]
                    if c is None:
                        continue
                    lst = get(c[0])
                    if not placed:
                        lst[c[1]] = sp.repl
                        placed = True
                    else:
                        lst[c[1]] = ""
            else:   # kiritish (nol uzunlikdagi span)
                i = sp.start
                if i > 0 and cmap[i - 1] is not None:
                    c = cmap[i - 1]
                    get(c[0])[c[1]] += sp.repl
                elif i < n and cmap[i] is not None:
                    c = cmap[i]
                    lst = get(c[0])
                    lst[c[1]] = sp.repl + lst[c[1]]
        for idx, lst in lists.items():
            new = "".join(lst)
            el = elems[idx]
            if new != el.text:
                el.text = new
                if (new[:1].isspace() or new[-1:].isspace()) and el.get(XML_SPACE) != "preserve":
                    el.set(XML_SPACE, "preserve")
                self.changed = True

    # -- oddiy matn (atribut, c:v, core.xml) --
    def plain(self, text: str) -> str:
        spans = convert_spans(text, self.opt)
        if not spans:
            return text
        for sp in spans:
            self.stats.counts[sp.kind] += 1
        if self.words is not None:
            for pair in changed_words(text, spans):
                self.words[pair] += 1
        return apply_spans(text, spans) if self.apply else text

    def element_text(self, el) -> None:
        if el.text:
            new = self.plain(el.text)
            if new != el.text:
                el.text = new
                self.changed = True

    def part(self, root) -> None:
        for p in root.iter(W_P, A_P):
            self.paragraph(p)
        for v in root.iter(C_V):
            self.element_text(v)
        for el in root.iter(*_DOCPR_TAGS):
            for attr in ("descr", "title"):
                val = el.get(attr)
                if val:
                    new = self.plain(val)
                    if new != val:
                        el.set(attr, new)
                        self.changed = True

    def core(self, root) -> None:
        for ns, name in ((DC, "title"), (DC, "subject"), (CP, "keywords")):
            for el in root.iter(f"{{{ns}}}{name}"):
                self.element_text(el)


_DECL_RE = re.compile(rb"^(\xef\xbb\xbf)?\s*<\?xml.*?\?>\s*", re.S)


def _process_part(name: str, data: bytes, proc: _Processor, core: bool) -> bytes | None:
    """O'zgargan bo'lsa yangi baytlarni, aks holda None qaytaradi."""
    try:
        tree = etree.parse(BytesIO(data), _parser())
    except etree.XMLSyntaxError:
        raise DocxError(f"Fayl ichidagi «{name}» qismi buzilgan.") from None
    proc.changed = False
    root = tree.getroot()
    if core:
        proc.core(root)
    else:
        proc.part(root)
    if not proc.changed:
        return None
    m = _DECL_RE.match(data)
    prefix = m.group(0) if m else b""
    if m and b"utf-8" not in prefix.lower():
        return etree.tostring(tree, encoding="UTF-8", xml_declaration=True, standalone=True)
    body = etree.tostring(tree, encoding="UTF-8", xml_declaration=False)
    return prefix + body


def _iter_work(zf: zipfile.ZipFile):
    text_parts = _text_parts(zf)
    names = {i.filename for i in zf.infolist()}
    for info in zf.infolist():
        if info.filename in text_parts:
            yield info, False
        elif info.filename == CORE_PROPS and CORE_PROPS in names:
            yield info, True
        else:
            yield info, None


# --- Ommaviy API -----------------------------------------------------------------

def preview_docx(src, options: ConvertOptions | None = None) -> tuple[Counter, ConversionStats]:
    """Faylni o'zgartirmasdan, o'zgaradigan so'zlar: Counter[(eski, yangi)] qaytaradi."""
    opt = options or DEFAULT_OPTIONS
    words: Counter = Counter()
    proc = _Processor(opt, apply=False, words=words)
    zf = check_file(src)
    try:
        for info, kind in _iter_work(zf):
            if kind is None:
                continue
            _process_part(info.filename, zf.read(info), proc, core=kind)
    except DocxError:
        raise
    except zipfile.BadZipFile:
        raise DocxError("Fayl buzilgan (zip xatosi).") from None
    finally:
        zf.close()
    return words, proc.stats


def convert_docx(src, dst, options: ConvertOptions | None = None) -> ConversionStats:
    """src ni o'qib, natijani dst ga yozadi (src ga tegilmaydi). dst mavjud bo'lsa ustiga yoziladi."""
    opt = options or DEFAULT_OPTIONS
    src, dst = Path(src), Path(dst)
    if src.resolve() == dst.resolve():
        raise DocxError("Natija fayli asl fayl bilan bir xil bo'lishi mumkin emas.")
    proc = _Processor(opt)
    total = ConversionStats()
    zf = check_file(src)
    modified: set[str] = set()
    try:
        infos = zf.infolist()
        order = sorted(range(len(infos)), key=lambda i: (infos[i].filename != "[Content_Types].xml", i))
        work = list(_iter_work(zf))
        kinds = {info.filename: kind for info, kind in work}
        try:
            with zipfile.ZipFile(dst, "w", allowZip64=True) as out:
                for i in order:
                    info = infos[i]
                    data = zf.read(info)
                    kind = kinds[info.filename]
                    if kind is not None:
                        new = _process_part(info.filename, data, proc, core=kind)
                        if new is not None:
                            data = new
                            modified.add(info.filename)
                    zi = zipfile.ZipInfo(info.filename, info.date_time)
                    zi.compress_type = info.compress_type
                    zi.external_attr = info.external_attr
                    zi.create_system = info.create_system
                    zi.comment = info.comment
                    out.writestr(zi, data, compress_type=info.compress_type)
        except zipfile.BadZipFile:
            raise DocxError("Fayl buzilgan (zip xatosi).") from None
        except PermissionError:
            raise DocxError("Natija faylini yozib bo'lmadi (ruxsat yo'q yoki fayl band).") from None
        except OSError as e:
            if getattr(e, "errno", None) == 28:
                raise DocxError("Diskda bo'sh joy yetarli emas.") from None
            raise DocxError("Natija faylini yozib bo'lmadi.") from None
        _verify_output(zf, dst, modified)
    finally:
        zf.close()
    total.add(proc.stats)
    total.parts_changed = len(modified)
    return total


def _verify_output(src_zf: zipfile.ZipFile, dst: Path, modified: set[str]) -> None:
    """Natijani tekshiradi: nomlar bir xil, o'zgarmagan qismlar bir xil, o'zgarganlari to'g'ri XML."""
    try:
        with zipfile.ZipFile(dst, "r") as chk:
            if chk.testzip() is not None:
                raise DocxError("Natija fayli tekshiruvdan o'tmadi (zip xatosi).")
            a = [i.filename for i in src_zf.infolist()]
            b = [i.filename for i in chk.infolist()]
            if sorted(a) != sorted(b):
                raise DocxError("Natija fayli tekshiruvdan o'tmadi (qismlar mos emas).")
            if b[0] != "[Content_Types].xml" and "[Content_Types].xml" in b:
                raise DocxError("Natija fayli tekshiruvdan o'tmadi (tuzilma).")
            src_crc = {i.filename: i.CRC for i in src_zf.infolist()}
            for i in chk.infolist():
                if i.filename in modified:
                    etree.fromstring(chk.read(i.filename), _parser())
                elif src_crc[i.filename] != i.CRC:
                    raise DocxError("Natija fayli tekshiruvdan o'tmadi (o'zgarmasligi kerak bo'lgan qism o'zgargan).")
    except DocxError:
        raise
    except (zipfile.BadZipFile, etree.XMLSyntaxError):
        raise DocxError("Natija fayli tekshiruvdan o'tmadi.") from None


# --- Fayl nomlari va yozish ------------------------------------------------------

def unique_path(directory: Path, stem: str, ext: str) -> Path:
    cand = directory / f"{stem}{ext}"
    n = 2
    while cand.exists():
        cand = directory / f"{stem}_{n}{ext}"
        n += 1
    return cand


def convert_file(src, options: ConvertOptions | None = None, output_dir=None,
                 suffix: str = "_yangi_alifbo", replace_original: bool = False) -> tuple[Path, ConversionStats]:
    """Faylni o'tkazib, tayyor natija yo'lini qaytaradi.

    replace_original=True bo'lsa asl fayl .bak nusxa sifatida saqlanadi va natija uning o'rniga yoziladi.
    """
    src = Path(src)
    suffix = suffix if suffix else "_yangi_alifbo"
    out_dir = Path(output_dir) if output_dir else src.parent
    if not replace_original:
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
        except OSError:
            raise DocxError("Natija papkasini yaratib bo'lmadi.") from None
    tmp_dir = src.parent if replace_original else out_dir
    fd, tmp_name = tempfile.mkstemp(prefix=".alifbo_", suffix=".tmp", dir=str(tmp_dir))
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        stats = convert_docx(src, tmp, options)
        if replace_original:
            bak = Path(str(src) + ".bak")
            n = 2
            while bak.exists():
                bak = Path(f"{src}.bak{n}")
                n += 1
            try:
                shutil.copy2(src, bak)
                os.replace(tmp, src)
            except PermissionError:
                raise DocxError("Asl faylni almashtirib bo'lmadi: u Word'da ochiq bo'lishi mumkin. "
                                "Faylni yopib, qayta urinib ko'ring.") from None
            return src, stats
        final = unique_path(out_dir, src.stem + suffix, src.suffix)
        os.replace(tmp, final)
        return final, stats
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
