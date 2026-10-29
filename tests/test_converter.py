import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pytest

from app.converter import ConvertOptions, changed_words, convert_spans, convert_text
from app.cyrillic import cyrillic_to_new_latin

T = "\u2019"  # standart tutuq belgisi


@pytest.mark.parametrize("src,expected", [
    ("bo\u2019limi", "bölimi"),
    ("bo'lim", "bölim"),
    ("bo\u2018lim", "bölim"),
    ("YIG\u2019MA JILDI", "YIĞMA JILDI"),
    ("AFCHO'SMUO'TM", "AFÇÖSMUÖTM"),
    ("maktabgacha va maktab ta\u2019limi boshqarmasi", "maktabgaça va maktab ta\u2019limi boşqarmasi"),
    ("O'zbekiston Respublikasi", "Özbekiston Respublikasi"),
    ("O\u02bbZBEKISTON", "ÖZBEKISTON"),
    ("G\u2018oya", "Ğoya"),
    ("g'oya", "ğoya"),
    ("to'g'ri", "töğri"),
    ("shahar, Shahar, SHAHAR", "şahar, Şahar, ŞAHAR"),
    ("Chilonzor, CHIROYLI", "Çilonzor, ÇIROYLI"),
    ("ishchi", "işçi"),
    ("mashhur", "maşhur"),
    ("tong, ming, uzoq", "tong, ming, uzoq"),
    ("ta'lim, a'zo, ma'no, san'at", f"ta{T}lim, a{T}zo, ma{T}no, san{T}at"),
    ("mas'hul", f"mas{T}hul"),
    ("as'hob", f"as{T}hob"),
    ("https://example.com/shop?ch=1", "https://example.com/shop?ch=1"),
    ("info@shahar.uz", "info@shahar.uz"),
    ("12345", "12345"),
    (r"C:\Shaxsiy\Chiqish\fayl.txt", r"C:\Shaxsiy\Chiqish\fayl.txt"),
    ("www.shahar.uz saytida", "www.shahar.uz saytida"),
    ("sho`ro, ko\u00b4z, o\u2032rik", "şöro, köz, örik"),
    ("Sho'ro", "Şöro"),
    ("g'o'za", "ğöza"),
    ("YOSH-chi", "YOŞ-çi"),
])
def test_convert(src, expected):
    assert convert_text(src) == expected


def test_idempotent():
    samples = [
        "O'zbekiston Respublikasi, shahar, ta'lim, to'g'ri, ISHCHI, mas'hul, Chilonzor",
        "bo\u2019limi YIG\u2019MA AFCHO'SMUO'TM mashhur",
    ]
    for s in samples:
        once = convert_text(s)
        assert convert_text(once) == once


def test_new_letters_unchanged():
    s = "Özbekiston ğoya şahar çiroyli ŞAHAR ÇIROYLI"
    assert convert_text(s) == s


def test_result_is_nfc_and_exact_codepoints():
    import unicodedata
    out = convert_text("O'zbek g'oya shahar chiroyli")
    assert unicodedata.normalize("NFC", out) == out
    assert "\u0219" not in out and "\u0218" not in out
    assert {"\u00d6", "\u011f", "\u015f", "\u00e7"} <= set(convert_text("O'z g'o sho cho"))


def test_apostrophe_choice():
    opt = ConvertOptions(apostrophe="\u02bc")
    assert convert_text("ta'lim bo'lim", opt) == "ta\u02bclim bölim"
    opt = ConvertOptions(apostrophe="'")
    assert convert_text("ta\u2019lim", opt) == "ta'lim"
    assert convert_text(convert_text("ta\u2019lim", opt), opt) == "ta'lim"


def test_quotes_not_touched():
    # tutuq belgisi emas: qo'shtirnoq/probel yonida
    assert convert_text("\u2018salom\u2019 dedi") == "\u2018salom\u2019 dedi"
    assert convert_text("90'lar") == "90'lar"
    # so'z oxirida o' + probel
    assert convert_text("go' ") == "go' "


def test_stop_words():
    opt = ConvertOptions(stop_words=["Chelsea", "Shell"])
    assert convert_text("Chelsea va Shell, shahar CHELSEA", opt) == "Chelsea va Shell, şahar CHELSEA"


def test_exceptions():
    opt = ConvertOptions(exceptions={"mo'tabar": "mo\u2019tabar", "g'ayrat": "g'ayrat"})
    assert convert_text("mo'tabar Mo'tabar MO'TABAR", opt) == "mo\u2019tabar Mo\u2019tabar MO\u2019TABAR"
    assert convert_text("G'ayrat g'ayrat", opt) == "G'ayrat g'ayrat"
    once = convert_text("mo'tabar", opt)
    assert convert_text(once, opt) == once


def test_spans_consistency():
    text = "bo'lim SHAHAR to\u2018g\u2018ri ta'lim"
    spans = convert_spans(text)
    last = 0
    for sp in spans:
        assert sp.start >= last and sp.end >= sp.start
        last = sp.end
    kinds = {sp.kind for sp in spans}
    assert {"o", "g", "sh", "apos"} <= kinds


def test_changed_words():
    text = "bo'lim shahar tong"
    ws = changed_words(text, convert_spans(text))
    assert ws == [("bo'lim", "bölim"), ("shahar", "şahar")]


def test_mixed_case_digraph():
    assert convert_text("ShAHAR") == "ŞAHAR"
    assert convert_text("sHAHAR") == "ŞAHAR"
    assert convert_text("sHahar") == "şahar"


def test_cyrillic_option():
    opt = ConvertOptions(cyrillic=True)
    assert convert_text("Ўзбекистон", opt) == "Özbekiston"
    assert convert_text("шаҳар чиройли ғоя", opt) == "şahar çiroyli ğoya"
    assert convert_text("Ёш, ЮРТ, ЕЛКА, мавзуе", opt) == "Yoş, YURT, YELKA, mavzuye"
    assert convert_text("маъно", opt) == f"ma{T}no"
    assert convert_text("Цирк", opt) == "Tsirk"
    # kirill o'chiq bo'lsa tegilmaydi
    assert convert_text("Ўзбекистон") == "Ўзбекистон"
    assert cyrillic_to_new_latin("қўл") == "qöl"
