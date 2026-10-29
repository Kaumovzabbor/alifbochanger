"""Oyna (tkinter + tkinterdnd2). Og'ir ish alohida oqimda bajariladi, interfeys qotmaydi."""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import APP_NAME, __version__
from .batch import (ST_DONE, ST_ERR, ST_RUN, BatchResult, collect_files, format_report,
                    run_batch, run_preview)
from .settings import (Settings, parse_exceptions_text, parse_stop_text, resource_path,
                       setup_logging)
from .strings import T

try:  # sudrab tashlash ixtiyoriy: ishlamasa tugmalar bilan ishlayveradi
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _HAS_DND = True
except Exception:  # pragma: no cover
    DND_FILES, TkinterDnD, _HAS_DND = None, None, False

LIGHT = {"bg": "#f4f5f7", "fg": "#1d2733", "card": "#ffffff", "accent": "#1f6feb", "accent_fg": "#ffffff",
         "muted": "#5b6776", "border": "#b8c0cc", "sel": "#cfe2ff", "ok": "#1a7f37", "err": "#c62828"}
DARK = {"bg": "#1e2228", "fg": "#e6e9ee", "card": "#2a2f37", "accent": "#4c8dff", "accent_fg": "#ffffff",
        "muted": "#9aa5b4", "border": "#4a5260", "sel": "#35507d", "ok": "#56d364", "err": "#ff7b72"}


def _fmt_size(n: int) -> str:
    if n < 1024:
        return f"{n} B"
    if n < 1024 ** 2:
        return f"{n / 1024:.0f} KB"
    return f"{n / 1024 ** 2:.1f} MB"


def open_folder(path: Path) -> bool:
    try:
        if os.name == "nt":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
        return True
    except Exception:
        return False


class App:
    def __init__(self):
        setup_logging()
        self.settings = Settings()
        self.root = None
        if _HAS_DND:
            try:
                self.root = TkinterDnD.Tk()
            except Exception:
                self.root = None
        self.dnd_ok = self.root is not None
        if self.root is None:
            self.root = tk.Tk()
        r = self.root
        r.title(T("app_title"))
        r.geometry("940x790")
        r.minsize(820, 720)
        try:
            r.iconbitmap(str(resource_path("assets/icon.ico")))
        except Exception:
            pass

        self.items: dict[str, Path] = {}       # iid -> yo'l
        self.keys: set[str] = set()
        self.counter = 0
        self.q: queue.Queue = queue.Queue()
        self.cancel = threading.Event()
        self.busy = False
        self.last_output_dir: Path | None = None
        self.output_dir = tk.StringVar(value=self.settings["output_dir"])
        self._themed_text: list[tk.Text] = []

        self.style = ttk.Style(r)
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass
        self.colors = DARK if self.settings["theme"] == "dark" else LIGHT
        self._build()
        self.apply_theme()
        r.protocol("WM_DELETE_WINDOW", self.on_close)
        r.after(100, self._poll)

    # ------------------------------------------------------------ interfeys
    def _build(self) -> None:
        r = self.root
        outer = ttk.Frame(r, padding=12)
        outer.pack(fill="both", expand=True)

        self.drop = ttk.Label(outer, anchor="center", justify="center", style="Drop.TLabel", cursor="hand2",
                              text=T("drop_here") + "\n" + T("drop_hint" if self.dnd_ok else "drop_hint_nodnd"))
        self.drop.pack(fill="x", ipady=16)
        self.drop.bind("<Button-1>", lambda e: self.pick_files())
        if self.dnd_ok:
            for w in (self.drop, r):
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:
                    self.dnd_ok = False

        row = ttk.Frame(outer)
        row.pack(fill="x", pady=(8, 4))
        ttk.Button(row, text=T("btn_pick_files"), command=self.pick_files).pack(side="left")
        ttk.Button(row, text=T("btn_pick_folder"), command=self.pick_folder).pack(side="left", padx=6)
        ttk.Button(row, text=T("btn_remove"), command=self.remove_selected).pack(side="left")
        ttk.Button(row, text=T("btn_clear"), command=self.clear).pack(side="left", padx=6)

        mid = ttk.Frame(outer)
        mid.pack(fill="both", expand=True)
        cols = ("name", "size", "status", "note")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="extended", height=8)
        for c, key, w, anc in (("name", "col_name", 330, "w"), ("size", "col_size", 80, "e"),
                               ("status", "col_status", 110, "center"), ("note", "col_note", 300, "w")):
            self.tree.heading(c, text=T(key))
            self.tree.column(c, width=w, anchor=anc, stretch=(c in ("name", "note")))
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tree.bind("<Double-1>", self._show_row_details)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())

        orow = ttk.Frame(outer)
        orow.pack(fill="x", pady=(8, 0))
        ttk.Label(orow, text=T("output_label")).pack(side="left")
        self.out_entry = ttk.Entry(orow, textvariable=self.output_dir)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(orow, text=T("btn_choose"), command=self.pick_output).pack(side="left")
        ttk.Button(orow, text=T("btn_reset"), command=lambda: self.output_dir.set("")).pack(side="left", padx=(6, 0))
        self.out_hint = ttk.Label(outer, text=T("output_default"), style="Muted.TLabel")
        self.out_hint.pack(anchor="w")

        brow = ttk.Frame(outer)
        brow.pack(fill="x", pady=(10, 6))
        self.btn_convert = ttk.Button(brow, text=T("btn_convert"), style="Big.TButton", command=self.start_convert)
        self.btn_convert.pack(side="left", fill="x", expand=True, ipady=8)
        self.btn_stop = ttk.Button(brow, text=T("btn_stop"), command=self.stop, state="disabled")
        self.btn_stop.pack(side="left", padx=(8, 0), ipady=8)

        srow = ttk.Frame(outer)
        srow.pack(fill="x")
        self.btn_preview = ttk.Button(srow, text=T("btn_preview"), command=self.open_preview)
        self.btn_preview.pack(side="left")
        ttk.Button(srow, text=T("btn_settings"), command=self.open_settings).pack(side="left", padx=6)
        ttk.Button(srow, text=T("btn_about"), command=self.open_about).pack(side="left")

        self.progress = ttk.Progressbar(outer, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(10, 2))
        self.status = ttk.Label(outer, text="", style="Muted.TLabel")
        self.status.pack(anchor="w")

        rep = ttk.Frame(outer)
        rep.pack(fill="x", pady=(6, 0))
        self.report = self._text(rep, height=8)
        self.report.pack(side="left", fill="x", expand=True)
        self.btn_open = ttk.Button(rep, text=T("btn_open_output"), command=self.open_output, state="disabled")
        self.btn_open.pack(side="left", padx=(8, 0), anchor="n")

    def _text(self, parent, height=6, state="disabled") -> tk.Text:
        t = tk.Text(parent, height=height, wrap="word", relief="flat", borderwidth=1, highlightthickness=1)
        self._themed_text.append(t)
        self._style_text(t)
        if state == "disabled":
            t.configure(state="disabled")
        return t

    def _style_text(self, t: tk.Text) -> None:
        c = self.colors
        t.configure(bg=c["card"], fg=c["fg"], insertbackground=c["fg"], highlightbackground=c["border"],
                    highlightcolor=c["accent"], selectbackground=c["sel"], selectforeground=c["fg"])

    def apply_theme(self) -> None:
        self.colors = c = DARK if self.settings["theme"] == "dark" else LIGHT
        r, s = self.root, self.style
        r.configure(bg=c["bg"])
        s.configure(".", background=c["bg"], foreground=c["fg"], fieldbackground=c["card"], bordercolor=c["border"],
                    lightcolor=c["bg"], darkcolor=c["bg"], troughcolor=c["card"])
        s.configure("TFrame", background=c["bg"])
        s.configure("TLabel", background=c["bg"], foreground=c["fg"])
        s.configure("Muted.TLabel", background=c["bg"], foreground=c["muted"])
        s.configure("Drop.TLabel", background=c["card"], foreground=c["fg"], relief="solid", borderwidth=2,
                    font=("Segoe UI", 13, "bold"))
        s.configure("TButton", background=c["card"], foreground=c["fg"], padding=(10, 5), bordercolor=c["border"])
        s.map("TButton", background=[("active", c["sel"]), ("disabled", c["bg"])],
              foreground=[("disabled", c["muted"])])
        s.configure("Big.TButton", background=c["accent"], foreground=c["accent_fg"], font=("Segoe UI", 12, "bold"))
        s.map("Big.TButton", background=[("active", c["accent"]), ("disabled", c["border"])],
              foreground=[("disabled", c["card"])])
        s.configure("TCheckbutton", background=c["bg"], foreground=c["fg"])
        s.configure("TRadiobutton", background=c["bg"], foreground=c["fg"])
        s.configure("TEntry", fieldbackground=c["card"], foreground=c["fg"])
        s.configure("Treeview", background=c["card"], fieldbackground=c["card"], foreground=c["fg"], rowheight=24)
        s.map("Treeview", background=[("selected", c["sel"])], foreground=[("selected", c["fg"])])
        s.configure("Treeview.Heading", background=c["bg"], foreground=c["fg"], font=("Segoe UI", 9, "bold"))
        s.configure("Horizontal.TProgressbar", background=c["accent"], troughcolor=c["card"])
        for t in self._themed_text:
            try:
                self._style_text(t)
            except tk.TclError:
                pass
        self.tree.tag_configure("done", foreground=c["ok"])
        self.tree.tag_configure("error", foreground=c["err"])

    # ------------------------------------------------------------ fayllar ro'yxati
    def _on_drop(self, event) -> None:
        try:
            paths = self.root.tk.splitlist(event.data)
        except tk.TclError:
            paths = [event.data]
        self.add_paths(list(paths))

    def pick_files(self) -> None:
        if self.busy:
            return
        paths = filedialog.askopenfilenames(
            title=T("dlg_pick_files"),
            filetypes=[(T("ft_docx"), "*.docx *.docm *.dotx *.dotm"), (T("ft_all"), "*.*")])
        if paths:
            self.add_paths(list(paths))

    def pick_folder(self) -> None:
        if self.busy:
            return
        d = filedialog.askdirectory(title=T("dlg_pick_folder"))
        if d:
            self.add_paths([d])

    def pick_output(self) -> None:
        d = filedialog.askdirectory(title=T("dlg_pick_output"))
        if d:
            self.output_dir.set(d)

    def add_paths(self, paths) -> None:
        if self.busy:
            messagebox.showinfo(T("app_title"), T("msg_busy"))
            return
        files = collect_files(paths, skip_outputs_suffix=self.settings["suffix"])
        added = 0
        for f in files:
            key = os.path.normcase(str(f.resolve()))
            if key in self.keys:
                continue
            self.keys.add(key)
            self.counter += 1
            iid = str(self.counter)
            try:
                size = _fmt_size(f.stat().st_size)
            except OSError:
                size = "?"
            self.items[iid] = f
            self.tree.insert("", "end", iid=iid, values=(f.name, size, T("st_wait"), str(f.parent)))
            added += 1
        if added:
            self.status.configure(text=T("msg_added", n=added))
            if len(self.items) > 200:
                self.status.configure(text=T("msg_too_many", n=len(self.items)))
        else:
            self.status.configure(text=T("msg_nothing_added"))

    def remove_selected(self) -> None:
        if self.busy:
            return
        for iid in self.tree.selection():
            f = self.items.pop(iid, None)
            if f is not None:
                self.keys.discard(os.path.normcase(str(f.resolve())))
            self.tree.delete(iid)

    def clear(self) -> None:
        if self.busy:
            return
        self.tree.delete(*self.tree.get_children())
        self.items.clear()
        self.keys.clear()
        self._set_report("")
        self.progress["value"] = 0
        self.status.configure(text="")

    def _show_row_details(self, _e) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        messagebox.showinfo(T("msg_error_details"), f"{vals[0]}\n\n{vals[2]}\n{vals[3]}")

    # ------------------------------------------------------------ o'tkazish
    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        self.btn_convert.configure(state="disabled" if busy else "normal")
        self.btn_preview.configure(state="disabled" if busy else "normal")
        self.btn_stop.configure(state="normal" if busy else "disabled")

    def _set_report(self, text: str) -> None:
        self.report.configure(state="normal")
        self.report.delete("1.0", "end")
        self.report.insert("1.0", text)
        self.report.configure(state="disabled")

    def _current_output_dir(self) -> str | None:
        v = self.output_dir.get().strip()
        return v or None

    def start_convert(self) -> None:
        if self.busy:
            messagebox.showinfo(T("app_title"), T("msg_busy"))
            return
        if not self.items:
            messagebox.showinfo(T("app_title"), T("msg_no_files"))
            return
        self.settings["output_dir"] = self.output_dir.get().strip()
        self.settings.save()
        replace = bool(self.settings["replace_original"])
        if replace and not messagebox.askyesno(T("app_title"), T("msg_replace_confirm")):
            return
        files = [self.items[i] for i in self.tree.get_children()]
        for iid in self.tree.get_children():
            self.tree.item(iid, values=(*self.tree.item(iid, "values")[:2], T("st_wait"), ""), tags=())
        self._set_report("")
        self.btn_open.configure(state="disabled")
        self.progress["value"] = 0
        self.cancel.clear()
        self._set_busy(True)
        opt = self.settings.options()
        out_dir, suffix = self._current_output_dir(), self.settings["suffix"]
        order = list(self.tree.get_children())

        def progress(i, total, fr):
            self.q.put(("progress", i, total, fr, order))

        def work():
            try:
                res = run_batch(files, opt, output_dir=out_dir, suffix=suffix, replace_original=replace,
                                progress=progress, cancel=self.cancel)
                self.q.put(("done", res))
            except Exception as e:  # juda kutilmagan holat
                self.q.put(("crash", type(e).__name__))

        threading.Thread(target=work, daemon=True).start()

    def stop(self) -> None:
        self.cancel.set()

    def _poll(self) -> None:
        try:
            while True:
                msg = self.q.get_nowait()
                kind = msg[0]
                if kind == "progress":
                    _, i, total, fr, order = msg
                    self._on_progress(i, total, fr, order)
                elif kind == "done":
                    self._on_done(msg[1])
                elif kind == "crash":
                    self._set_busy(False)
                    messagebox.showerror(T("app_title"), f"Kutilmagan xato: {msg[1]}")
                elif kind == "preview":
                    msg[1](*msg[2:])
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def _iid_for(self, path: Path, order) -> str | None:
        for iid in order:
            if self.items.get(iid) == path:
                return iid
        return None

    def _on_progress(self, i, total, fr, order) -> None:
        iid = self._iid_for(fr.path, order)
        if iid is not None and iid in self.tree.get_children():
            name, size = self.tree.item(iid, "values")[:2]
            label = {ST_RUN: T("st_run"), ST_DONE: T("st_done"), ST_ERR: T("st_err")}.get(fr.status, T("st_wait"))
            note = fr.message if fr.status in (ST_DONE, ST_ERR) else ""
            if fr.status == ST_DONE and fr.output:
                note = f"{fr.output.name}  ({fr.message})"
            tags = ("done",) if fr.status == ST_DONE else (("error",) if fr.status == ST_ERR else ())
            self.tree.item(iid, values=(name, size, label, note), tags=tags)
            self.tree.see(iid)
        self.progress["value"] = (i / total * 100) if total else 0
        self.status.configure(text=T("msg_working", done=i, total=total))

    def _on_done(self, res: BatchResult) -> None:
        self._set_busy(False)
        self.progress["value"] = 100 if not res.cancelled else self.progress["value"]
        self.status.configure(text=T("msg_cancelled") if res.cancelled else T("msg_finished"))
        self._set_report(f"{T('report_title')}\n" + format_report(res))
        outs = [f.output for f in res.files if f.output]
        if outs:
            self.last_output_dir = outs[-1].parent
            self.btn_open.configure(state="normal")

    def open_output(self) -> None:
        if self.last_output_dir and not open_folder(self.last_output_dir):
            messagebox.showwarning(T("app_title"), T("msg_open_failed"))

    # ------------------------------------------------------------ ko'rib chiqish
    def open_preview(self) -> None:
        if self.busy:
            return
        if not self.items:
            messagebox.showinfo(T("app_title"), T("msg_no_files"))
            return
        files = [self.items[i] for i in self.tree.get_children()]
        opt = self.settings.options()
        win = tk.Toplevel(self.root)
        win.title(T("pv_title"))
        win.geometry("700x520")
        win.configure(bg=self.colors["bg"])
        win.transient(self.root)
        frame = ttk.Frame(win, padding=10)
        frame.pack(fill="both", expand=True)
        info = ttk.Label(frame, text=T("pv_loading"))
        info.pack(anchor="w")
        ttk.Label(frame, text=T("pv_hint"), style="Muted.TLabel", wraplength=660).pack(anchor="w", pady=(0, 6))
        tw = ttk.Frame(frame)
        tw.pack(fill="both", expand=True)
        tree = ttk.Treeview(tw, columns=("old", "new", "n"), show="headings", selectmode="extended")
        sortstate: dict[str, bool] = {}

        def sort_by(col):
            rev = sortstate.get(col, False)
            rows = [(tree.set(k, col), k) for k in tree.get_children("")]
            if col == "n":
                rows.sort(key=lambda t: int(t[0]), reverse=rev)
            else:
                rows.sort(key=lambda t: t[0].lower(), reverse=rev)
            for idx, (_, k) in enumerate(rows):
                tree.move(k, "", idx)
            sortstate[col] = not rev

        for col, key, w in (("old", "pv_old", 260), ("new", "pv_new", 260), ("n", "pv_count", 80)):
            tree.heading(col, text=T(key), command=lambda c=col: sort_by(c))
            tree.column(col, width=w, anchor="w" if col != "n" else "e")
        sb = ttk.Scrollbar(tw, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        btns = ttk.Frame(frame)
        btns.pack(fill="x", pady=(8, 0))
        msg = ttk.Label(btns, text="", style="Muted.TLabel")

        def dont_change():
            pairs = {}
            for k in tree.selection():
                old = tree.set(k, "old")
                pairs[old] = old
            if not pairs:
                return
            self.settings.add_exceptions(pairs)
            for k in tree.selection():
                tree.delete(k)
            msg.configure(text=T("pv_added", n=len(pairs)))

        ttk.Button(btns, text=T("pv_dont"), command=dont_change).pack(side="left")
        ttk.Button(btns, text=T("pv_close"), command=win.destroy).pack(side="right")
        msg.pack(side="left", padx=10)
        cancel = threading.Event()
        win.bind("<Destroy>", lambda e: cancel.set() if e.widget is win else None)

        def fill(words, errors):
            if not win.winfo_exists():
                return
            rows = sorted(words.items(), key=lambda kv: (-kv[1], kv[0][0].lower()))[:20000]
            for (old, new), n in rows:
                tree.insert("", "end", values=(old, new, n))
            text = T("pv_summary", n=len(words)) if words else T("pv_empty")
            if errors:
                text += "   " + T("pv_errors", n=len(errors))
            info.configure(text=text)

        def work():
            words, errors = run_preview(files, opt, cancel=cancel)
            self.q.put(("preview", fill, words, errors))

        threading.Thread(target=work, daemon=True).start()

    # ------------------------------------------------------------ sozlamalar
    def open_settings(self) -> None:
        if self.busy:
            return
        st = self.settings
        win = tk.Toplevel(self.root)
        win.title(T("set_title"))
        win.geometry("640x680")
        win.configure(bg=self.colors["bg"])
        win.transient(self.root)
        win.grab_set()
        f = ttk.Frame(win, padding=12)
        f.pack(fill="both", expand=True)

        ttk.Label(f, text=T("set_apos")).pack(anchor="w")
        apos = tk.StringVar(value=st["apostrophe"])
        ar = ttk.Frame(f)
        ar.pack(anchor="w", pady=(2, 8))
        for val, key in (("’", "set_apos_curly"), ("ʼ", "set_apos_mod"), ("'", "set_apos_plain")):
            ttk.Radiobutton(ar, text=T(key), value=val, variable=apos).pack(side="left", padx=(0, 14))

        only_uz = tk.BooleanVar(value=st["only_uzbek_lang"])
        ttk.Checkbutton(f, text=T("set_only_uz"), variable=only_uz).pack(anchor="w", pady=(2, 0))
        ttk.Label(f, text=T("set_only_uz_hint"), style="Muted.TLabel", wraplength=590, justify="left").pack(
            anchor="w", padx=(22, 0), pady=(0, 4))
        cyr = tk.BooleanVar(value=st["cyrillic"])
        ttk.Checkbutton(f, text=T("set_cyr"), variable=cyr).pack(anchor="w", pady=2)
        repl = tk.BooleanVar(value=st["replace_original"])
        ttk.Checkbutton(f, text=T("set_replace"), variable=repl).pack(anchor="w", pady=2)

        sr = ttk.Frame(f)
        sr.pack(fill="x", pady=(8, 4))
        ttk.Label(sr, text=T("set_suffix")).pack(side="left")
        suffix = tk.StringVar(value=st["suffix"])
        ttk.Entry(sr, textvariable=suffix, width=24).pack(side="left", padx=6)

        tr = ttk.Frame(f)
        tr.pack(fill="x", pady=(4, 8))
        ttk.Label(tr, text=T("set_theme")).pack(side="left")
        theme = tk.StringVar(value=st["theme"])
        ttk.Radiobutton(tr, text=T("set_light"), value="light", variable=theme).pack(side="left", padx=8)
        ttk.Radiobutton(tr, text=T("set_dark"), value="dark", variable=theme).pack(side="left")

        ttk.Label(f, text=T("set_stop")).pack(anchor="w")
        stop_t = self._text(f, height=6, state="normal")
        stop_t.pack(fill="x", pady=(2, 8))
        stop_t.insert("1.0", "\n".join(st.stop_list))

        ttk.Label(f, text=T("set_exc")).pack(anchor="w")
        exc_t = self._text(f, height=9, state="normal")
        exc_t.pack(fill="both", expand=True, pady=(2, 8))
        exc_t.insert("1.0", "\n".join(f"{k} => {v}" for k, v in st.exceptions.items()))

        def save():
            st["apostrophe"] = apos.get()
            st["only_uzbek_lang"] = bool(only_uz.get())
            st["cyrillic"] = bool(cyr.get())
            st["replace_original"] = bool(repl.get())
            st["suffix"] = suffix.get().strip() or "_yangi_alifbo"
            st["theme"] = theme.get()
            st.stop_list = parse_stop_text(stop_t.get("1.0", "end"))
            st.exceptions = parse_exceptions_text(exc_t.get("1.0", "end"))
            st.save()
            self._themed_text = [t for t in self._themed_text if t.winfo_exists() and t not in (stop_t, exc_t)]
            self.apply_theme()
            win.destroy()

        def cancel():
            self._themed_text = [t for t in self._themed_text if t not in (stop_t, exc_t)]
            win.destroy()

        b = ttk.Frame(f)
        b.pack(fill="x")
        ttk.Button(b, text=T("btn_save"), style="Big.TButton", command=save).pack(side="right")
        ttk.Button(b, text=T("btn_cancel"), command=cancel).pack(side="right", padx=8)

    def open_about(self) -> None:
        messagebox.showinfo(T("about_title"), T("about_text", name=APP_NAME, version=__version__))

    def on_close(self) -> None:
        self.cancel.set()
        self.settings["output_dir"] = self.output_dir.get().strip()
        self.settings.save()
        self.root.destroy()

    def run(self) -> int:
        self.root.mainloop()
        return 0


def run_gui() -> int:
    return App().run()
