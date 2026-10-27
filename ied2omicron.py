# -*- coding: utf-8 -*-
"""
ied2omicron.py - "IED → OMICRON Auto Config": GUI for the OMICRON files of a General_PB project
(author: German Ievgen).

  python ied2omicron.py

Needs next to it: omicron_configuration.py, omicron_goose_config.py, omicron_xrio.py, ied_doc_manual.py
Needs: Python 3 (with tkinter), openpyxl (pip install openpyxl).

  Project folder          General_PB.xlsm, DOC\\, Library\\, CID\\
  Create Omicron_configurations  -> Omicron\\<IED>\\<IED>_Omicron_AI_Test_<n>.ohc
  Create GOOSE_configurations    -> Omicron\\<IED>\\<IED>_GOOSE.xml
  Create XRIO_test_objects       -> Omicron\\<IED>\\test_object_<IED>_AI_Test_<n>.xrio
                                    (an existing file is kept: new file with date and time)
  Before creating: the list of devices with check boxes.
  Update GOOSE_configurations    -> choose existing *_GOOSE.xml file(s), then the folder with
                                    the new CID files -> <name>_updated_<date>_<time>.xml
"""
import os, re, sys, json, queue, threading, traceback, webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import omicron_configuration as oc            # noqa: E402
import omicron_goose_config as og             # noqa: E402
import omicron_xrio as ox                     # noqa: E402
import ied_doc_manual as man                  # noqa: E402

APP = "IED → OMICRON Auto Config"
VERSION = "v 0.0.4 (dev)"
SETTINGS = os.path.join(os.path.expanduser("~"), ".ied2omicron.json")
OLD_SETTINGS = os.path.join(os.path.expanduser("~"), ".ied_doc_configurator.json")   # v0.0.2 before the rename


JOBS = {"ohc": "Omicron_configurations", "goose": "GOOSE_configurations", "xrio": "XRIO_test_objects"}
REPORTS = {JOBS["ohc"]: oc.REPORT, JOBS["goose"]: og.REPORT, JOBS["xrio"]: ox.REPORT}


def load_settings():
    for p in (SETTINGS, OLD_SETTINGS):
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_settings(d):
    try:
        cur = load_settings()
        cur.update(d)
        with open(SETTINGS, "w", encoding="utf-8") as f:
            json.dump(cur, f)
    except Exception:
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("%s  %s" % (APP, VERSION))
        self.geometry("980x600")
        self.minsize(700, 480)
        self.q = queue.Queue()
        self.busy = False
        self.overlay = None
        st = load_settings()

        # ---- header
        top = ttk.Frame(self, padding=(14, 10, 14, 4))
        top.pack(fill="x")
        ttk.Label(top, text=APP, font=("Segoe UI", 18, "bold")).pack(side="left")
        ttk.Label(top, text="  " + VERSION, font=("Segoe UI", 11)).pack(side="left", pady=(8, 0))
        ttk.Button(top, text="i", width=3, command=self.show_info).pack(side="right")

        # ---- project folder
        pf = ttk.Frame(self, padding=(14, 6))
        pf.pack(fill="x")
        ttk.Label(pf, text="Project folder:").pack(side="left")
        self.prj = tk.StringVar(value=st.get("project", ""))
        ttk.Entry(pf, textvariable=self.prj).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(pf, text="Browse...", command=self.browse).pack(side="left")
        self.b_open = ttk.Button(pf, text="Open Omicron folder", command=self.open_out)
        self.b_open.pack(side="left", padx=(6, 0))

        # ---- actions
        bf = ttk.Frame(self, padding=(14, 6))
        bf.pack(fill="x")
        self.b_ohc = ttk.Button(bf, text="Create Omicron_configurations", command=lambda: self.start("ohc"))
        self.b_ohc.pack(side="left", ipadx=10, ipady=6)
        self.b_goose = ttk.Button(bf, text="Create GOOSE_configurations", command=lambda: self.start("goose"))
        self.b_goose.pack(side="left", padx=10, ipadx=10, ipady=6)
        self.b_upd = ttk.Button(bf, text="Update GOOSE_configurations", command=self.start_update)
        self.b_upd.pack(side="left", ipadx=10, ipady=6)
        self.b_xrio = ttk.Button(bf, text="Create XRIO_test_objects", command=lambda: self.start("xrio"))
        self.b_xrio.pack(side="left", padx=(10, 0), ipadx=10, ipady=6)

        # ---- log
        lf = ttk.Frame(self, padding=(14, 4, 14, 0))
        lf.pack(fill="both", expand=True)
        self.log = ScrolledText(lf, height=18, font=("Consolas", 9), state="disabled", wrap="none")
        self.log.pack(fill="both", expand=True)
        self.status = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.status, padding=(14, 4)).pack(fill="x")
        self.after(100, self.poll)

    # ---------------------------------------------------------------- helpers
    def write(self, s):
        self.log.configure(state="normal")
        self.log.insert("end", s + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def browse(self):
        p = filedialog.askdirectory(title="Project folder (General_PB.xlsm, DOC, Library, CID)",
                                    initialdir=self.prj.get() or None)
        if p:
            self.prj.set(os.path.normpath(p))
            save_settings({"project": self.prj.get()})

    def open_out(self):
        p = os.path.join(self.prj.get(), oc.OUT_DIR)
        if not os.path.isdir(p):
            messagebox.showinfo(APP, "The folder does not exist yet:\n" + p)
            return
        try:
            os.startfile(p)                    # Windows
        except Exception:
            messagebox.showinfo(APP, p)

    def set_busy(self, b):
        self.busy = b
        for w in (self.b_ohc, self.b_goose, self.b_upd, self.b_xrio):
            w.configure(state="disabled" if b else "normal")

    def close_overlay(self):
        if self.overlay is not None:
            self.overlay.destroy()
            self.overlay = None

    def make_overlay(self, title, big=False):
        """a panel inside the main window (not a separate window)"""
        self.close_overlay()
        fr = tk.Frame(self, bd=2, relief="ridge", bg="#f4f4f4")
        if big:
            fr.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.96, relheight=0.94)
        else:
            fr.place(relx=0.5, rely=0.52, anchor="center", relwidth=0.8, relheight=0.78)
        hd = tk.Frame(fr, bg="#dfe6ee")
        hd.pack(fill="x")
        tk.Label(hd, text=title, bg="#dfe6ee", font=("Segoe UI", 11, "bold"), padx=10, pady=6).pack(side="left")
        tk.Button(hd, text="X", relief="flat", bg="#dfe6ee", command=self.close_overlay).pack(side="right", padx=4)
        self.overlay = fr
        return fr

    # ---------------------------------------------------------------- info / manual
    def show_info(self):
        fr = self.make_overlay("%s  %s" % (APP, VERSION), big=True)
        hd = fr.winfo_children()[0]
        lang = tk.StringVar(value=load_settings().get("lang", man.DEFAULT_LANG))
        txt = tk.Text(fr, wrap="word", font=("Segoe UI", 10), padx=16, pady=10, bg="white",
                      relief="flat", cursor="arrow")
        sb = ttk.Scrollbar(fr, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        bb = tk.Frame(fr, bg="#f4f4f4")
        bb.pack(side="bottom", fill="x", pady=6)
        ttk.Button(bb, text="OK", command=self.close_overlay).pack()
        sb.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)
        txt.tag_configure("title", font=("Segoe UI", 15, "bold"), spacing1=4, spacing3=8)
        txt.tag_configure("h2", font=("Segoe UI", 12, "bold"), foreground="#1f4e79", spacing1=12, spacing3=4)
        txt.tag_configure("mono", font=("Consolas", 9), lmargin1=16, lmargin2=40)
        txt.tag_configure("bullet", lmargin1=20, lmargin2=34, spacing3=2)
        txt.tag_configure("p", spacing3=4)
        txt.bind("<MouseWheel>", lambda e: txt.yview_scroll(int(-e.delta / 120), "units"))

        def render(*_):
            code = lang.get()
            save_settings({"lang": code})
            txt.configure(state="normal")
            txt.delete("1.0", "end")
            nlink = 0
            for line in man.MANUAL[code].replace("{version}", VERSION).split("\n"):
                if line.startswith("## "):
                    base, line = ("h2",), line[3:]
                elif line.startswith("# "):
                    base, line = ("title",), line[2:]
                elif re.match(r"^\s{4}(•|\d\.)", line):
                    base, line = ("bullet",), line.strip()
                elif line.startswith("    "):
                    base = ("mono",)
                else:
                    base = ("p",)
                pos = 0
                for m in re.finditer(r"\[\[(.+?)\|(.+?)\]\]", line):
                    txt.insert("end", line[pos:m.start()], base)
                    nlink += 1
                    tag = "link%d" % nlink
                    txt.tag_configure(tag, foreground="#0563c1", underline=True)
                    txt.tag_bind(tag, "<Button-1>", lambda e, u=m.group(2): webbrowser.open(u))
                    txt.tag_bind(tag, "<Enter>", lambda e: txt.configure(cursor="hand2"))
                    txt.tag_bind(tag, "<Leave>", lambda e: txt.configure(cursor="arrow"))
                    txt.insert("end", m.group(1), base + (tag,))
                    pos = m.end()
                txt.insert("end", line[pos:] + "\n", base)
            txt.configure(state="disabled")
            txt.yview_moveto(0)

        for code, name in reversed(man.LANGS):
            tk.Radiobutton(hd, text=name, value=code, variable=lang, command=render, indicatoron=False,
                           bg="#dfe6ee", selectcolor="#b7c9dd", relief="flat", padx=8,
                           font=("Segoe UI", 9)).pack(side="right", padx=2)
        render()

    # ---------------------------------------------------------------- devices
    def start(self, kind):
        if self.busy:
            return
        prj = self.prj.get().strip()
        mod = {"ohc": oc, "goose": og, "xrio": ox}[kind]
        try:
            ieds = mod.list_devices(prj)
        except mod.UserError as ex:
            messagebox.showerror(APP, str(ex))
            return
        if not ieds:
            messagebox.showerror(APP, "No device files in %s" % os.path.join(prj, "DOC"))
            return
        save_settings({"project": prj})
        title = JOBS[kind] + " - devices"
        fr = self.make_overlay(title)

        # master choice (only .ohc, only if there are several)
        master = tk.StringVar()
        if kind == "ohc":
            ms = oc.list_masters(prj)
            if not ms:
                self.close_overlay()
                messagebox.showerror(APP, "No master .ohc file in %s" % os.path.join(prj, "Library"))
                return
            master.set(ms[0])
            if len(ms) > 1:
                mf = tk.Frame(fr, bg="#f4f4f4", padx=10, pady=4)
                mf.pack(fill="x")
                tk.Label(mf, text="Master file:", bg="#f4f4f4").pack(side="left")
                ttk.Combobox(mf, textvariable=master, values=ms, state="readonly", width=40).pack(side="left", padx=6)

        # check boxes (scrollable)
        body = tk.Frame(fr, bg="#f4f4f4")
        body.pack(fill="both", expand=True, padx=10, pady=4)
        cv = tk.Canvas(body, bg="white", highlightthickness=1, highlightbackground="#bbb")
        sb = ttk.Scrollbar(body, orient="vertical", command=cv.yview)
        inner = tk.Frame(cv, bg="white")
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.create_window((0, 0), window=inner, anchor="nw")
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(int(-e.delta / 120), "units") if self.overlay else None)
        vars_ = []
        cols = 3
        for i, ied in enumerate(ieds):
            v = tk.BooleanVar(value=False)
            tk.Checkbutton(inner, text=ied, variable=v, bg="white", anchor="w").grid(
                row=i // cols, column=i % cols, sticky="w", padx=8, pady=1)
            vars_.append((ied, v))

        bb = tk.Frame(fr, bg="#f4f4f4", pady=8)
        bb.pack(fill="x", padx=10)
        ttk.Button(bb, text="Select all", command=lambda: [v.set(True) for _, v in vars_]).pack(side="left")
        ttk.Button(bb, text="None", command=lambda: [v.set(False) for _, v in vars_]).pack(side="left", padx=6)

        def ok():
            sel = [i for i, v in vars_ if v.get()]
            if not sel:
                messagebox.showinfo(APP, "Select at least one device.", parent=self)
                return
            m = master.get()
            self.close_overlay()
            self.run_job(kind, prj, sel, m)
        ttk.Button(bb, text="Cancel", command=self.close_overlay).pack(side="right")
        ttk.Button(bb, text="Create", command=ok).pack(side="right", padx=6)

    # ---------------------------------------------------------------- update GOOSE
    def start_update(self):
        if self.busy:
            return
        st = load_settings()
        prj = self.prj.get().strip()
        start_xml = st.get("upd_xml") or (os.path.join(prj, oc.OUT_DIR) if prj else None)
        files = filedialog.askopenfilenames(
            title="GOOSE configuration file(s) to update (*.xml)",
            initialdir=start_xml if start_xml and os.path.isdir(start_xml) else None,
            filetypes=[("GOOSE configuration", "*.xml"), ("All files", "*.*")])
        if not files:
            return
        files = [os.path.normpath(f) for f in files]
        start_cid = st.get("upd_cid") or (os.path.join(prj, og.CID_SUB) if prj else None)
        cid = filedialog.askdirectory(title="Folder with the NEW CID files",
                                      initialdir=start_cid if start_cid and os.path.isdir(start_cid) else None)
        if not cid:
            return
        cid = os.path.normpath(cid)
        save_settings({"upd_xml": os.path.dirname(files[0]), "upd_cid": cid})
        self.set_busy(True)
        name = "Update GOOSE_configurations"
        self.status.set("Working: %s (%d file(s)) ..." % (name, len(files)))
        self.write("=" * 70)
        self.write("%s: %s" % (name, ", ".join(os.path.basename(f) for f in files)))
        self.write("New CID: " + cid)

        def log(s):
            self.q.put(("log", s))

        def work():
            try:
                made, lines = og.update_files(files, cid, log=log)
                self.q.put(("done", (name, made, lines)))
            except og.UserError as ex:
                self.q.put(("error", str(ex)))
            except Exception:
                self.q.put(("error", "Unexpected error:\n" + traceback.format_exc()))
        threading.Thread(target=work, daemon=True).start()

    # ---------------------------------------------------------------- work (thread)
    def run_job(self, kind, prj, sel, master):
        self.set_busy(True)
        name = JOBS[kind]
        self.status.set("Working: %s (%d device(s)) ..." % (name, len(sel)))
        self.write("=" * 70)
        self.write("%s: %s" % (name, ", ".join(sel)))

        def log(s):
            self.q.put(("log", s))

        def work():
            try:
                probs = []
                if kind == "ohc":
                    made, lines = oc.run(prj, sel, master or None, log=log)
                elif kind == "xrio":
                    made, lines, probs = ox.run(prj, sel, log=log)
                else:
                    made, lines, probs = og.run(prj, sel, log=log)
                self.q.put(("done", (name, made, lines, probs)))
            except (oc.UserError, og.UserError) as ex:
                self.q.put(("error", str(ex)))
            except Exception:
                self.q.put(("error", "Unexpected error:\n" + traceback.format_exc()))
        threading.Thread(target=work, daemon=True).start()

    def poll(self):
        try:
            while True:
                what, val = self.q.get_nowait()
                if what == "log":
                    self.write(val)
                elif what == "done":
                    name, made, lines = val[:3]
                    probs = val[3] if len(val) > 3 else []
                    self.write("")
                    for l in lines:
                        self.write(l)
                    self.set_busy(False)
                    if name.startswith("Update"):
                        self.status.set("%s: %d file(s) written next to the originals" % (name, made))
                        messagebox.showinfo(APP, "%s: %d file(s) written next to the originals.\n\n"
                                                 "Report: _goose_update_report_<date>_<time>.txt" % (name, made))
                    elif probs:
                        # XRIO: tests whose SV streams have different / empty nominal values;
                        # GOOSE: devices without the own CID file of a publisher / subscriber IED
                        self.status.set("%s: %d file(s) written, %d problem(s)" % (name, made, len(probs)))
                        more = "\n  ... (%d more)" % (len(probs) - 12) if len(probs) > 12 else ""
                        messagebox.showwarning(APP, "%s: %d file(s) written.\n\nNOT created:\n  %s%s\n\nReport: %s\\%s" %
                                               (name, made, "\n  ".join(probs[:12]), more, oc.OUT_DIR, REPORTS[name]))
                    else:
                        self.status.set("%s: %d file(s) written - folder %s" % (name, made, oc.OUT_DIR))
                        messagebox.showinfo(APP, "%s: %d file(s) written.\n\nReport: %s\\%s" %
                                            (name, made, oc.OUT_DIR, REPORTS[name]))
                elif what == "error":
                    self.write(val)
                    self.set_busy(False)
                    self.status.set("Stopped")
                    messagebox.showerror(APP, val)
        except queue.Empty:
            pass
        self.after(100, self.poll)


if __name__ == "__main__":
    App().mainloop()
