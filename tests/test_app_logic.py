import json
import os
import re
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import pytest

from app.batch import ST_DONE, ST_ERR, collect_files, format_report, run_batch, run_preview
from app.converter import ConvertOptions
from app.settings import Settings, parse_exceptions_text, parse_stop_text
from tests.test_docx import make_docx, para, run, texts


def test_collect_files_skips_lock_and_previous_outputs(tmp_path):
    a = make_docx(tmp_path / "a.docx", para(run("x")))
    (tmp_path / "sub").mkdir()
    b = make_docx(tmp_path / "sub" / "b.docx", para(run("x")))
    (tmp_path / "~$a.docx").write_bytes(b"lock")
    make_docx(tmp_path / "a_yangi_alifbo.docx", para(run("x")))
    make_docx(tmp_path / "a_yangi_alifbo_2.docx", para(run("x")))
    (tmp_path / "notes.txt").write_text("x")
    got = collect_files([tmp_path])
    assert {p.name for p in got} == {"a.docx", "b.docx"}
    assert [p.name for p in collect_files([tmp_path], recursive=False)] == ["a.docx"]
    # takrorlar yo'q; aniq tanlangan fayl har doim qo'shiladi
    assert len(collect_files([a, a, tmp_path])) == 2
    assert [p.name for p in collect_files([tmp_path / "a_yangi_alifbo.docx"])] == ["a_yangi_alifbo.docx"]


def test_run_batch_continues_after_errors(tmp_path):
    good1 = make_docx(tmp_path / "g1.docx", para(run("O'zbek shahar")))
    bad = tmp_path / "bad.docx"
    bad.write_bytes(b"zip emas")
    good2 = make_docx(tmp_path / "g2.docx", para(run("to'g'ri chiroyli")))
    events = []
    res = run_batch([good1, bad, good2], ConvertOptions(), progress=lambda i, n, fr: events.append((i, n, fr.status)))
    assert [f.status for f in res.files] == [ST_DONE, ST_ERR, ST_DONE]
    assert res.ok == 2 and res.failed == 1
    assert "buzilgan" in res.files[1].message
    assert texts(res.files[0].output)[0] == "Özbek şahar"
    assert texts(res.files[2].output)[0] == "töğri çiroyli"
    assert res.counts["o"] == 2 and res.counts["g"] == 1 and res.counts["sh"] == 1 and res.counts["ch"] == 1
    rep = format_report(res)
    assert "3 ta" not in rep and "2 ta" in rep and "ö" in rep and "ş" in rep
    assert events[0][2] == "run" and events[-1][0] == 3


def test_run_batch_cancel(tmp_path):
    import threading
    f = make_docx(tmp_path / "g.docx", para(run("shahar")))
    ev = threading.Event()
    ev.set()
    res = run_batch([f], ConvertOptions(), cancel=ev)
    assert res.cancelled and res.files[0].status == "wait"


def test_preview_batch(tmp_path):
    f = make_docx(tmp_path / "g.docx", para(run("bo'lim bo'lim shahar")))
    bad = tmp_path / "bad.docx"
    bad.write_bytes(b"x")
    words, errors = run_preview([f, bad], ConvertOptions())
    assert words[("bo'lim", "bölim")] == 2 and len(errors) == 1


def test_settings_roundtrip_and_options(tmp_path):
    s = Settings(tmp_path)
    s["apostrophe"] = "ʼ"
    s["cyrillic"] = True
    s["theme"] = "dark"
    s.stop_list = ["Chelsea"]
    s.exceptions = {"mo'tabar": "mo’tabar"}
    s.save()
    s2 = Settings(tmp_path)
    assert s2["apostrophe"] == "ʼ" and s2["cyrillic"] is True and s2["theme"] == "dark"
    o = s2.options()
    assert o.apostrophe == "ʼ" and o.cyrillic and "Chelsea" in o.stop_words


def test_settings_survive_corrupt_files(tmp_path):
    (tmp_path / "settings.json").write_text("{buzuq", encoding="utf-8")
    (tmp_path / "exceptions.json").write_text("[1,2", encoding="utf-8")
    s = Settings(tmp_path)
    assert s["apostrophe"] == "’" and s.options() is not None
    (tmp_path / "settings.json").write_text(json.dumps({"apostrophe": "X", "theme": 5}), encoding="utf-8")
    s = Settings(tmp_path)
    assert s["apostrophe"] == "’" and s["theme"] == "light"


def test_add_exceptions(tmp_path):
    s = Settings(tmp_path)
    s.add_exceptions({"mashhur": "mashhur"})
    assert Settings(tmp_path).exceptions == {"mashhur": "mashhur"}


def test_parsers():
    assert parse_exceptions_text("a => b\nbo'tabar -> bo’tabar\nx = y\nyolg'iz\n =>\n") == {
        "a": "b", "bo'tabar": "bo’tabar", "x": "y"}
    assert parse_stop_text("Chelsea\n\n shell \nCHELSEA\n") == ["Chelsea", "shell"]


def test_all_used_string_keys_exist():
    keys = set(json.loads((ROOT / "app" / "strings_uz.json").read_text(encoding="utf-8")))
    used = set()
    for f in ("gui.py", "cli.py", "batch.py"):
        used |= set(re.findall(r'\bT\(\s*"([a-z_0-9]+)"', (ROOT / "app" / f).read_text(encoding="utf-8")))
    used |= set(re.findall(r'"((?:btn|col|st|msg|set|pv|dlg|ft|about)_[a-z_0-9]+)"', (ROOT / "app" / "gui.py").read_text(encoding="utf-8")))
    assert used - keys == set(), used - keys


def test_cli_end_to_end(tmp_path):
    src = make_docx(tmp_path / "Hujjat.docx", para(run("O'zbekiston shahri")))
    out = tmp_path / "natija"
    env = dict(os.environ, XDG_CONFIG_HOME=str(tmp_path / "cfg"), APPDATA=str(tmp_path / "cfg"),
               PYTHONIOENCODING="utf-8")
    r = subprocess.run([sys.executable, str(ROOT / "main.py"), str(src), "-o", str(out)],
                       capture_output=True, text=True, encoding="utf-8", env=env, timeout=60)
    assert r.returncode == 0, r.stderr
    assert "[OK]" in r.stdout and "Hujjat_yangi_alifbo.docx" in r.stdout
    assert texts(out / "Hujjat_yangi_alifbo.docx")[0] == "Özbekiston şahri"
    r = subprocess.run([sys.executable, str(ROOT / "main.py"), str(src), "--preview"],
                       capture_output=True, text=True, encoding="utf-8", env=env, timeout=60)
    assert "Özbekiston" in r.stdout and r.returncode == 0
    bad = tmp_path / "bad.docx"
    bad.write_bytes(b"x")
    r = subprocess.run([sys.executable, str(ROOT / "main.py"), str(bad)],
                       capture_output=True, text=True, encoding="utf-8", env=env, timeout=60)
    assert r.returncode == 1 and "[XATO]" in r.stdout
