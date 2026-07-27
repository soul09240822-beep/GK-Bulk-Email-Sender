#!/usr/bin/env python3
"""
GK Bulk Email Sender  v2.0
Golden Key POS
- Large UI fonts (2x)
- Rich text editor (Bold / Italic / Underline / Size / Color / Align / Image / Link)
- Gmail + Custom SMTP
- Mail-merge variables from ALL Excel columns  {{COLUMN}}
- Signature with logo
- Newsletter-ready HTML email
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, colorchooser
import smtplib, threading, subprocess, platform, time, json, sys, base64
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def _require_pandas():
    try:
        import pandas as pd
        return pd
    except ImportError:
        messagebox.showerror("Missing Package",
            "pandas not found.\n\nRun:  pip install pandas openpyxl\nthen restart the app.")
        return None

# ══════════════════════════════════════════════════════════
#  CONSTANTS
# ══════════════════════════════════════════════════════════
APP  = "GK Bulk Email Sender"
VER  = "v2.0"

DARK   = "#1d1d1f"
GREEN  = "#34c759"
RED    = "#ff3b30"
BLUE   = "#0066cc"
LIGHT  = "#f5f5f5"
WHITE  = "#ffffff"
MUTED  = "#888888"
BORDER = "#dddddd"
WARN_BG= "#fffde7"

# ── 2× font sizes ──────────────────────────────────────
F_XS    = ("Arial", 13)
F_SM    = ("Arial", 15)
F_NORM  = ("Arial", 17)
F_BOLD  = ("Arial", 17, "bold")
F_LG    = ("Arial", 20)
F_LGB   = ("Arial", 20, "bold")
F_TITLE = ("Arial", 28, "bold")
F_MONO  = ("Courier", 15)
F_EDIT  = ("Arial", 16)   # email body editor default

# ══════════════════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════════════════
DEFAULT_CFG = {
    "server_type": "gmail",
    "gmail_email": "", "gmail_pass": "",
    "smtp_host": "", "smtp_port": "465", "smtp_ssl": True,
    "smtp_user": "", "smtp_pass": "",
    "from_name": "", "test_email": "",
    "excel_file": "", "pdf_col": "",
    "pdf_folder": "", "faq_file": "",
    "subject": "", "body": "",
    "sig_on": False, "sig_logo": "", "sig_text": "",
    "test_mode": True, "delay": "2.0",
}

def _cfg_path():
    base = Path(sys.executable).parent if getattr(sys,"frozen",False) else Path(__file__).parent
    return base / "gk_config.json"

def load_cfg():
    p = _cfg_path()
    if p.exists():
        try:
            with open(p, encoding="utf-8") as f:
                return {**DEFAULT_CFG, **json.load(f)}
        except: pass
    return dict(DEFAULT_CFG)

def save_cfg(c):
    try:
        with open(_cfg_path(), "w", encoding="utf-8") as f:
            json.dump(c, f, indent=2, ensure_ascii=False)
    except Exception as e:
        messagebox.showwarning("Warning", f"Could not save:\n{e}")

# ══════════════════════════════════════════════════════════
#  EMAIL UTILITIES
# ══════════════════════════════════════════════════════════
def load_merchants(excel_path, pdf_col):
    """Returns (rows:list[dict], columns:list[str], error:str|None)"""
    pd = _require_pandas()
    if pd is None: return None, [], "pandas not available"
    try:
        df = pd.read_excel(excel_path, dtype=str)
    except Exception as e:
        return None, [], str(e)
    df.columns = df.columns.str.strip().str.upper()
    if "EMAIL" not in df.columns:
        return None, list(df.columns), "EMAIL column not found"
    df = df.dropna(subset=["EMAIL"])
    df = df[df["EMAIL"].str.contains("@", na=False)]
    rows = [dict(r) for _, r in df.iterrows()]
    return rows, list(df.columns), None

def apply_vars(template, row):
    """Replace {{COLUMN}} with values from row dict"""
    result = template
    for col, val in row.items():
        safe = str(val) if (val is not None and str(val) != "nan") else ""
        result = result.replace(f"{{{{{col}}}}}", safe)
        result = result.replace(f"[{col}]", safe)   # legacy [NAME] style
    return result

def load_rtfd(path):
    if platform.system() != "Darwin":
        return None, ".rtfd is macOS only.\nUse 'Load .txt' instead."
    import tempfile
    tmp = Path(tempfile.gettempdir()) / "_gk_tmp.txt"
    r = subprocess.run(["textutil","-convert","txt","-output",str(tmp),path],
                       capture_output=True, text=True)
    if r.returncode != 0: return None, r.stderr
    text = tmp.read_text(encoding="utf-8")
    try: tmp.unlink()
    except: pass
    return text, None

def _attach(msg, path, name):
    with open(path,"rb") as f:
        part = MIMEBase("application","octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{name}"')
    msg.attach(part)

def build_msg(cfg, row, to_email, plain, html):
    from_addr = cfg["gmail_email"] if cfg["server_type"]=="gmail" else cfg["smtp_user"]
    subject   = apply_vars(cfg["subject"], row)

    outer = MIMEMultipart("mixed")
    outer["From"]    = f'{cfg["from_name"]} <{from_addr}>'
    outer["To"]      = to_email
    outer["Subject"] = subject

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(plain, "plain", "utf-8"))
    alt.attach(MIMEText(html,  "html",  "utf-8"))
    outer.attach(alt)

    # Individual PDF
    if cfg["pdf_folder"]:
        name = str(row.get(cfg["pdf_col"].upper(),"")).strip()
        if name:
            p = Path(cfg["pdf_folder"]) / f"{name}.pdf"
            if p.exists(): _attach(outer, p, f"{name}.pdf")

    # Common attachment
    if cfg["faq_file"]:
        p = Path(cfg["faq_file"])
        if p.exists(): _attach(outer, p, p.name)

    return outer

def smtp_connect(cfg):
    if cfg["server_type"] == "gmail":
        s = smtplib.SMTP("smtp.gmail.com", 587)
        s.starttls()
        s.login(cfg["gmail_email"], cfg["gmail_pass"].replace(" ",""))
        return s, cfg["gmail_email"]
    else:
        port = int(cfg["smtp_port"])
        if cfg["smtp_ssl"]:
            s = smtplib.SMTP_SSL(cfg["smtp_host"], port)
        else:
            s = smtplib.SMTP(cfg["smtp_host"], port)
            s.starttls()
        user = cfg["smtp_user"] or cfg["gmail_email"]
        s.login(user, cfg["smtp_pass"])
        return s, user

def send_one(cfg, row, to_email, plain, html, retry=3):
    msg = build_msg(cfg, row, to_email, plain, html)
    _, from_addr = smtp_connect(cfg)   # test connect only once outside loop
    for attempt in range(1, retry+1):
        try:
            s, fa = smtp_connect(cfg)
            s.sendmail(fa, to_email, msg.as_string())
            s.quit()
            return True
        except Exception:
            if attempt < retry: time.sleep(3)
    return False

# ══════════════════════════════════════════════════════════
#  HTML BUILDER
# ══════════════════════════════════════════════════════════
def widget_to_html(widget):
    """Convert tkinter Text with tags → inline-styled HTML"""
    end = widget.index("end-1c")
    if widget.compare("1.0",">=",end): return ""
    parts = []
    idx = "1.0"
    while widget.compare(idx,"<",end):
        ch   = widget.get(idx)
        tags = set(widget.tag_names(idx))
        sty  = []
        if "bold"      in tags: sty.append("font-weight:bold")
        if "italic"    in tags: sty.append("font-style:italic")
        if "underline" in tags: sty.append("text-decoration:underline")
        for t in tags:
            if t.startswith("sz_"):
                sty.append(f"font-size:{t[3:]}px")
            elif t.startswith("clr_"):
                sty.append(f"color:#{t[4:]}")
        if ch == "\n":
            parts.append("<br>\n")
        else:
            safe = ch.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
            parts.append(f'<span style="{";".join(sty)}">{safe}</span>' if sty else safe)
        idx = widget.index(f"{idx}+1c")
    return "".join(parts)

def wrap_html(body_html, sig_html=""):
    sig = (f'<br><hr style="border:none;border-top:1px solid #ddd;margin:24px 0">'
           f'<div style="color:#555;font-size:13px">{sig_html}</div>') if sig_html else ""
    return (f'<!DOCTYPE html><html><head><meta charset="UTF-8">'
            f'<style>body{{font-family:Arial,sans-serif;font-size:14px;'
            f'line-height:1.7;color:#333;max-width:680px;margin:0 auto;padding:24px}}'
            f'</style></head><body>{body_html}{sig}</body></html>')

# ══════════════════════════════════════════════════════════
#  RICH TEXT EDITOR WIDGET
# ══════════════════════════════════════════════════════════
class RichEditor(tk.Frame):
    def __init__(self, parent, height=14, **kw):
        super().__init__(parent, bg=WHITE, **kw)
        self._imgs = []   # keep image refs alive
        self._build_toolbar()
        self._build_text(height)

    # ── Toolbar ───────────────────────────────────────────
    def _build_toolbar(self):
        tb = tk.Frame(self, bg="#eeeeee", bd=1, relief="solid", pady=5)
        tb.pack(fill="x")

        def btn(text, cmd, **kw):
            font = kw.pop("font", F_SM)
            fg   = kw.pop("fg", DARK)
            b = tk.Button(tb, text=text, font=font, fg=fg, bd=1, relief="raised",
                         bg=WHITE, cursor="hand2", padx=8, pady=3, command=cmd, **kw)
            b.pack(side="left", padx=2)
            return b

        def sep():
            tk.Frame(tb, bg=BORDER, width=2, height=30).pack(side="left",fill="y",padx=6,pady=2)

        btn("B",  self.bold,   font=("Arial",17,"bold"))
        btn("I",  self.italic, font=("Arial",17,"italic"))
        btn("U̲", self.under)
        sep()

        tk.Label(tb, text="Size:", font=F_XS, bg="#eeeeee").pack(side="left", padx=(2,0))
        self._sz = tk.StringVar(value="14")
        cb = ttk.Combobox(tb, textvariable=self._sz, width=5, state="readonly",
                          font=F_XS,
                          values=["10","11","12","13","14","16","18","20","22","24","28","32","36","48"])
        cb.pack(side="left", padx=3)
        cb.bind("<<ComboboxSelected>>", self._size)
        sep()

        btn("≡L", lambda: self._align("al_left"))
        btn("≡C", lambda: self._align("al_center"))
        btn("≡R", lambda: self._align("al_right"))
        sep()

        btn("A🎨", self._color, fg=RED, font=("Arial",16,"bold"))
        sep()

        btn("🖼 Image", self._img)
        btn("🔗 Link",  self._link)

    def _build_text(self, height):
        frame = tk.Frame(self, bd=1, relief="solid")
        frame.pack(fill="both", expand=True, pady=(4,0))

        self.text = tk.Text(frame, font=F_EDIT, wrap="word", height=height,
                            padx=14, pady=12, undo=True, spacing1=3, spacing3=3,
                            bd=0, relief="flat")
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=sb.set)
        self.text.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # Tag definitions
        self.text.tag_config("bold",      font=(F_EDIT[0], F_EDIT[1], "bold"))
        self.text.tag_config("italic",    font=(F_EDIT[0], F_EDIT[1], "italic"))
        self.text.tag_config("underline", underline=True)
        self.text.tag_config("al_left",   justify="left")
        self.text.tag_config("al_center", justify="center")
        self.text.tag_config("al_right",  justify="right")

        # Shortcuts
        for key, fn in [("<Control-b>",self.bold),("<Command-b>",self.bold),
                        ("<Control-i>",self.italic),("<Command-i>",self.italic),
                        ("<Control-u>",self.under),("<Command-u>",self.under)]:
            self.text.bind(key, lambda e, f=fn: f() or "break")

    # ── Formatting actions ─────────────────────────────────
    def _toggle(self, tag):
        try:
            s,e = self.text.index("sel.first"), self.text.index("sel.last")
            (self.text.tag_remove if tag in self.text.tag_names(s) else self.text.tag_add)(tag,s,e)
        except tk.TclError: pass

    def bold(self):   self._toggle("bold")
    def italic(self): self._toggle("italic")
    def under(self):  self._toggle("underline")

    def _size(self, _=None):
        sz = self._sz.get(); tag = f"sz_{sz}"
        self.text.tag_config(tag, font=("Arial", int(sz)))
        try: self.text.tag_add(tag,"sel.first","sel.last")
        except tk.TclError: pass

    def _align(self, tag):
        ls = self.text.index("insert linestart")
        le = self.text.index("insert lineend+1c")
        for t in ["al_left","al_center","al_right"]:
            self.text.tag_remove(t,ls,le)
        self.text.tag_add(tag,ls,le)

    def _color(self):
        clr = colorchooser.askcolor(title="Pick Text Color")[1]
        if clr:
            tag = f"clr_{clr.replace('#','')}"
            self.text.tag_config(tag, foreground=clr)
            try: self.text.tag_add(tag,"sel.first","sel.last")
            except tk.TclError: pass

    def _img(self):
        path = filedialog.askopenfilename(title="Insert Image",
                filetypes=[("Images","*.png *.jpg *.jpeg *.gif *.webp"),("All","*.*")])
        if not path: return
        try:
            from PIL import Image, ImageTk
            img   = Image.open(path); img.thumbnail((500,400))
            photo = ImageTk.PhotoImage(img)
            self.text.image_create("insert", image=photo)
            self._imgs.append(photo)
        except ImportError:
            self.text.insert("insert", f"[Image:{Path(path).name}]")
            self._imgs.append(path)

    def _link(self):
        win = tk.Toplevel(self); win.title("Insert Link")
        win.geometry("480x200"); win.configure(bg=WHITE); win.grab_set()
        tk.Label(win, text="Display Text:", font=F_SM, bg=WHITE).pack(padx=24,pady=(18,4),anchor="w")
        t = tk.Entry(win, font=F_SM, bd=1, relief="solid"); t.pack(fill="x",padx=24)
        tk.Label(win, text="URL:", font=F_SM, bg=WHITE).pack(padx=24,pady=(10,4),anchor="w")
        u = tk.Entry(win, font=F_SM, bd=1, relief="solid"); u.pack(fill="x",padx=24)
        u.insert(0,"https://")
        def _do():
            disp = t.get() or u.get(); url=u.get()
            tag = f"lnk_{len(self._imgs)}"
            self.text.tag_config(tag, foreground=BLUE, underline=True)
            self.text.insert("insert", disp, tag)
            self._imgs.append(("lnk",url)); win.destroy()
        tk.Button(win, text="Insert", font=F_BOLD, bg=DARK, fg=WHITE, relief="flat",
                 command=_do, pady=10).pack(fill="x",padx=24,pady=14)

    # ── Public API ──────────────────────────────────────────
    def get_plain(self): return self.text.get("1.0","end-1c")
    def get_html(self):  return widget_to_html(self.text)
    def set_text(self, txt):
        self.text.delete("1.0","end"); self.text.insert("1.0", txt)
    def insert_at_cursor(self, txt):
        self.text.insert("insert", txt); self.text.focus()

# ══════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ══════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP}  {VER}")
        self.geometry("1160x880")
        self.minsize(900,720)
        self.configure(bg=LIGHT)
        self.cfg  = load_cfg()
        self._run = False
        self._cols: list[str] = []
        self._html_override: str | None = None   # HTML fetched from URL
        self._build_header()
        self._build_nb()
        self._populate()

    # ── Header ────────────────────────────────────────────
    def _build_header(self):
        h = tk.Frame(self, bg=DARK, height=64)
        h.pack(fill="x"); h.pack_propagate(False)
        tk.Label(h, text=f"⚡  {APP}", font=F_TITLE, fg=WHITE, bg=DARK
                 ).pack(side="left", padx=24, pady=16)
        tk.Label(h, text=f"{VER}  ·  {platform.system()}", font=F_SM,
                 fg=MUTED, bg=DARK).pack(side="right", padx=24)

    # ── Notebook ──────────────────────────────────────────
    def _build_nb(self):
        s = ttk.Style()
        s.configure("TNotebook",     background=LIGHT, borderwidth=0)
        s.configure("TNotebook.Tab", font=F_LG, padding=[18,10])
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=16, pady=16)
        tabs = [
            ("  ⚙  Setup  ",             self._build_setup),
            ("  📂  Files & Recipients  ",self._build_files),
            ("  ✉  Email Content  ",      self._build_content),
            ("  🚀  Send  ",              self._build_send),
        ]
        for title, builder in tabs:
            f = tk.Frame(self.nb, bg=LIGHT)
            self.nb.add(f, text=title)
            builder(f)

    # ── helpers ───────────────────────────────────────────
    def _card(self, parent, title):
        f = tk.LabelFrame(parent, text=f"  {title}  ", font=F_LGB,
                          bg=WHITE, fg=DARK, bd=1, relief="solid", padx=20, pady=14)
        f.pack(fill="x", padx=18, pady=(0,16))
        return f

    def _lbl(self, p, t):
        tk.Label(p, text=t, font=F_SM, bg=WHITE, fg=MUTED, anchor="w"
                 ).pack(fill="x", pady=(12,3))

    def _ent(self, p, **kw):
        e = tk.Entry(p, font=F_NORM, bd=1, relief="solid", bg=WHITE, **kw)
        e.pack(fill="x", pady=(0,4)); return e

    def _brow_row(self, parent, lbl, var, cmd):
        self._lbl(parent, lbl)
        r = tk.Frame(parent, bg=WHITE); r.pack(fill="x")
        tk.Entry(r, textvariable=var, font=F_NORM, bd=1, relief="solid", bg=WHITE
                 ).pack(side="left", fill="x", expand=True, padx=(0,8))
        tk.Button(r, text="Browse", font=F_SM, bd=1, relief="solid",
                  bg=LIGHT, cursor="hand2", padx=12, command=cmd).pack(side="left")

    def _pw_row(self, parent, entry_ref_name):
        r = tk.Frame(parent, bg=WHITE); r.pack(fill="x", pady=(0,4))
        e = tk.Entry(r, font=F_NORM, bd=1, relief="solid", bg=WHITE, show="•")
        e.pack(side="left", fill="x", expand=True, padx=(0,8))
        setattr(self, entry_ref_name, e)
        tk.Button(r, text="👁", font=F_SM, bd=1, relief="solid", bg=LIGHT,
                  cursor="hand2", width=3,
                  command=lambda: e.config(show="" if e.cget("show")=="•" else "•")
                  ).pack(side="left")

    # ════════════════════════════════════════════════════════
    #  TAB 1 — Setup
    # ════════════════════════════════════════════════════════
    def _build_setup(self, tab):
        inner = tk.Frame(tab, bg=LIGHT)
        inner.pack(fill="both", expand=True, pady=16)

        # Server type toggle
        c0 = self._card(inner, "📡 Email Server")
        self._srv = tk.StringVar(value="gmail")
        r0 = tk.Frame(c0, bg=WHITE); r0.pack(fill="x", pady=8)
        tk.Radiobutton(r0, text="Gmail / Google Workspace",
                       variable=self._srv, value="gmail",
                       font=F_NORM, bg=WHITE, command=self._srv_toggle
                       ).pack(side="left", padx=(0,30))
        tk.Radiobutton(r0, text="Custom SMTP",
                       variable=self._srv, value="custom",
                       font=F_NORM, bg=WHITE, command=self._srv_toggle
                       ).pack(side="left")

        # Gmail card
        self._gcard = self._card(inner, "🔑 Gmail / Google Workspace")
        self._lbl(self._gcard, "Gmail Address")
        self._g_email = self._ent(self._gcard)
        self._lbl(self._gcard, "App Password  (16-character — NOT your regular password)")
        self._pw_row(self._gcard, "_g_pass")

        info = tk.Text(self._gcard, font=F_XS, height=6, bd=1, relief="solid",
                       bg=WARN_BG, padx=12, pady=8, wrap="word")
        info.pack(fill="x", pady=(10,0))
        info.insert("1.0",
            "📋 How to create an App Password:\n"
            "1. Go to myaccount.google.com\n"
            "2. Security → Turn ON 2-Step Verification (required first)\n"
            "3. Security → search 'App passwords' → Create → Copy the 16-char code\n"
            "4. Paste it above (spaces are fine)")
        info.config(state="disabled", fg="#555555")

        # Custom SMTP card
        self._scard = self._card(inner, "🖥 Custom SMTP Server")
        self._lbl(self._scard, "SMTP Host")
        self._smtp_host = self._ent(self._scard)

        pr = tk.Frame(self._scard, bg=WHITE); pr.pack(fill="x")
        lf = tk.Frame(pr, bg=WHITE); lf.pack(side="left", fill="x", expand=True, padx=(0,20))
        tk.Label(lf, text="Port", font=F_SM, bg=WHITE, fg=MUTED, anchor="w"
                 ).pack(fill="x", pady=(12,3))
        self._smtp_port = tk.Entry(lf, font=F_NORM, bd=1, relief="solid", bg=WHITE)
        self._smtp_port.pack(fill="x")
        rf = tk.Frame(pr, bg=WHITE); rf.pack(side="left", anchor="s", pady=(0,4))
        self._ssl_var = tk.BooleanVar(value=True)
        tk.Checkbutton(rf, text="Use SSL (port 465)", variable=self._ssl_var,
                       font=F_NORM, bg=WHITE).pack()

        self._lbl(self._scard, "Username  (usually same as email address)")
        self._smtp_user = self._ent(self._scard)
        self._lbl(self._scard, "Password")
        self._pw_row(self._scard, "_smtp_pass")

        # General
        c3 = self._card(inner, "👤 General")
        self._lbl(c3, "From Name  (displayed to recipient)")
        self._from_name = self._ent(c3)
        self._lbl(c3, "Test Email  (receives all mail in Test Mode)")
        self._test_email = self._ent(c3)

        tk.Button(inner, text="💾   Save Settings", font=F_LGB,
                  bg=DARK, fg=WHITE, relief="flat", cursor="hand2", pady=14,
                  command=self._save_cfg).pack(fill="x", padx=18)

    def _srv_toggle(self):
        is_gmail = self._srv.get() == "gmail"
        for child in self._gcard.winfo_children():
            child.configure(state="normal" if is_gmail else "disabled") if isinstance(child, (tk.Entry, tk.Button)) else None
        for child in self._scard.winfo_children():
            child.configure(state="normal" if not is_gmail else "disabled") if isinstance(child, (tk.Entry, tk.Button)) else None

    def _save_cfg(self):
        self._sync(); save_cfg(self.cfg)
        messagebox.showinfo("Saved", "Settings saved ✓\nThey'll be remembered next time.")

    # ════════════════════════════════════════════════════════
    #  TAB 2 — Files & Recipients
    # ════════════════════════════════════════════════════════
    def _build_files(self, tab):
        inner = tk.Frame(tab, bg=LIGHT)
        inner.pack(fill="both", expand=True, pady=16)

        c1 = self._card(inner, "📊 Excel File  (Recipients List)")

        ef = tk.Frame(c1, bg=WHITE); ef.pack(fill="x")
        self._xls = tk.StringVar()
        tk.Entry(ef, textvariable=self._xls, font=F_NORM, bd=1, relief="solid", bg=WHITE
                 ).pack(side="left", fill="x", expand=True, padx=(0,8))
        tk.Button(ef, text="Browse", font=F_SM, bd=1, relief="solid",
                  bg=LIGHT, cursor="hand2", padx=12, command=self._browse_xls).pack(side="left")
        tk.Button(c1, text="🔄  Load Columns from File", font=F_SM, bd=1, relief="solid",
                 bg=LIGHT, cursor="hand2", pady=6,
                 command=self._load_cols).pack(anchor="w", pady=(10,0))

        self._lbl(c1, "Mail-Merge Variables  (click any to insert into Subject or Body):")
        self._pframe = tk.Frame(c1, bg=WHITE)
        self._pframe.pack(fill="x", pady=6)
        tk.Label(self._pframe, text="← Browse and load an Excel file to see columns",
                font=F_XS, bg=WHITE, fg=MUTED).pack(anchor="w")

        self._lbl(c1, "PDF Filename Column  (value must match the PDF filenames in your PDF folder)")
        self._pdf_col = tk.StringVar(value="")
        self._pdf_col_cb = ttk.Combobox(c1, textvariable=self._pdf_col, font=F_NORM,
                                         state="readonly", width=26)
        self._pdf_col_cb.pack(anchor="w", pady=(0,4))

        c2 = self._card(inner, "📎 Attachments  (optional)")
        self._pdf_dir = tk.StringVar()
        self._brow_row(c2, "PDF Folder  (individual PDFs named by PDF column value, e.g. ABC_STORE.pdf)",
                       self._pdf_dir, lambda: self._pdf_dir.set(
                           filedialog.askdirectory(title="Select PDF Folder") or self._pdf_dir.get()))
        self._faq    = tk.StringVar()
        self._brow_row(c2, "Common Attachment  (same file sent to ALL recipients)",
                       self._faq, lambda: self._faq.set(
                           filedialog.askopenfilename(title="Select File",
                               filetypes=[("PDF","*.pdf"),("All","*.*")]) or self._faq.get()))

    def _browse_xls(self):
        p = filedialog.askopenfilename(title="Select Excel File",
                filetypes=[("Excel","*.xlsx *.xls"),("All","*.*")])
        if p:
            self._xls.set(p); self._load_cols()

    def _load_cols(self):
        path = self._xls.get()
        if not path: return
        pd = _require_pandas()
        if pd is None: return
        try:
            df = pd.read_excel(path, dtype=str, nrows=1)
            cols = [c.strip().upper() for c in df.columns]
            self._cols = cols
            self._pdf_col_cb["values"] = cols
            if self._pdf_col.get() not in cols:
                self._pdf_col.set("")
            self._refresh_pills()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _refresh_pills(self):
        frames = [self._pframe]
        if hasattr(self, "_cpills"):
            frames.append(self._cpills)
        for frame in frames:
            for w in frame.winfo_children(): w.destroy()
        for col in self._cols:
            v = f"{{{{{col}}}}}"
            for frame in frames:
                tk.Button(frame, text=v, font=("Courier",14,"bold"),
                         bd=1, relief="solid", bg="#e8f4fd", cursor="hand2",
                         padx=6, pady=3,
                         command=lambda s=v: self._editor.insert_at_cursor(s)
                         ).pack(side="left", padx=3, pady=3)

    # ════════════════════════════════════════════════════════
    #  TAB 3 — Email Content
    # ════════════════════════════════════════════════════════
    def _build_content(self, tab):
        outer = tk.Frame(tab, bg=LIGHT)
        outer.pack(fill="both", expand=True, padx=18, pady=16)

        # Subject row
        sr = tk.Frame(outer, bg=LIGHT); sr.pack(fill="x", pady=(0,12))
        tk.Label(sr, text="Subject:", font=F_LGB, bg=LIGHT).pack(side="left", padx=(0,12))
        self._subj = tk.StringVar()
        tk.Entry(sr, textvariable=self._subj, font=F_NORM, bd=1, relief="solid"
                 ).pack(side="left", fill="x", expand=True)

        # Load buttons
        br = tk.Frame(outer, bg=LIGHT); br.pack(fill="x", pady=(0,6))
        for lbl, cmd in [
            ("📄 Load .rtfd (macOS only)", self._ld_rtfd),
            ("📄 Load .txt",               self._ld_txt),
            ("🌐 Load from URL",           self._ld_url),
            ("🗑  Clear Body",             self._clear_body),
        ]:
            tk.Button(br, text=lbl, font=F_SM, bd=1, relief="solid",
                     bg=WHITE, cursor="hand2", padx=12, pady=5, command=cmd
                     ).pack(side="left", padx=(0,8))

        # URL status bar — shows when HTML is loaded from URL
        self._url_bar = tk.Frame(outer, bg="#e3f2fd", bd=1, relief="solid")
        self._url_lbl = tk.Label(self._url_bar, text="", font=F_XS,
                                  bg="#e3f2fd", fg="#1565c0", anchor="w", padx=10, pady=6)
        self._url_lbl.pack(side="left", fill="x", expand=True)
        tk.Button(self._url_bar, text="✕ Clear", font=F_XS, bd=0, relief="flat",
                 bg="#e3f2fd", fg="#1565c0", cursor="hand2",
                 command=self._clear_body).pack(side="right", padx=8)

        # Variable pills — auto-populated when Excel is loaded
        vrow = tk.Frame(outer, bg=LIGHT); vrow.pack(fill="x", pady=(0,10))
        tk.Label(vrow, text="Insert Variable:", font=F_SM, bg=LIGHT,
                fg=MUTED).pack(side="left", padx=(0,10))
        self._cpills = tk.Frame(vrow, bg=LIGHT)
        self._cpills.pack(side="left", fill="x", expand=True)
        tk.Label(self._cpills,
                text="← Go to Files & Recipients tab, load your Excel file to see column variables here",
                font=F_XS, bg=LIGHT, fg=MUTED).pack(anchor="w")

        # Body label + shortcut hint on same row
        bl = tk.Frame(outer, bg=LIGHT); bl.pack(fill="x")
        tk.Label(bl, text="Body:", font=F_LGB, bg=LIGHT).pack(side="left")
        tk.Label(bl, text="  ⌘/Ctrl+B = Bold  ·  ⌘/Ctrl+I = Italic  ·  ⌘/Ctrl+U = Underline",
                font=F_XS, bg=LIGHT, fg=MUTED).pack(side="left", padx=12)

        # Signature — pinned to bottom, always visible
        sig = tk.LabelFrame(outer, text="  ✍  Email Signature  ", font=F_LGB,
                            bg=WHITE, fg=DARK, bd=1, relief="solid", padx=18, pady=12)
        sig.pack(fill="x", side="bottom", pady=(10,0))

        # Rich text editor — sits above signature
        self._editor = RichEditor(outer, height=12)
        self._editor.pack(fill="both", expand=True, pady=(6,4))

        r0 = tk.Frame(sig, bg=WHITE); r0.pack(fill="x", pady=(0,6))
        self._sig_on = tk.BooleanVar(value=False)
        tk.Checkbutton(r0, text="Include signature at bottom of every email",
                      variable=self._sig_on, font=F_NORM, bg=WHITE,
                      command=self._sig_toggle).pack(side="left")

        # Details frame — hidden until checkbox checked
        self._sig_details = tk.Frame(sig, bg=WHITE)

        tk.Label(self._sig_details, bg="#e8f5e9", fg="#2e7d32", font=F_SM,
                anchor="w", padx=12, pady=10, justify="left",
                text="💡  Logo: upload your company logo image (PNG / JPG)\n"
                     "    Signature Text: company name, phone, address, website, etc.\n"
                     "    Both are optional — leave blank if not needed."
                ).pack(fill="x", pady=(0,12))

        lr = tk.Frame(self._sig_details, bg=WHITE); lr.pack(fill="x", pady=(0,8))
        tk.Label(lr, text="Logo Image:", font=F_SM, bg=WHITE, fg=MUTED,
                 width=14, anchor="w").pack(side="left")
        self._sig_logo = tk.StringVar()
        tk.Entry(lr, textvariable=self._sig_logo, font=F_NORM, bd=1, relief="solid"
                 ).pack(side="left", fill="x", expand=True, padx=(0,8))
        tk.Button(lr, text="Browse", font=F_SM, bd=1, relief="solid", bg=LIGHT,
                  cursor="hand2", padx=12,
                  command=lambda: self._sig_logo.set(
                      filedialog.askopenfilename(title="Select Logo Image",
                          filetypes=[("Images","*.png *.jpg *.jpeg"),("All","*.*")]
                      ) or self._sig_logo.get()
                  )).pack(side="left")

        tk.Label(self._sig_details,
                text="Signature Text:  (e.g. company name · phone · website)",
                font=F_SM, bg=WHITE, fg=MUTED, anchor="w").pack(fill="x", pady=(6,3))
        self._sig_txt = tk.Text(self._sig_details, font=F_EDIT, height=7,
                                bd=1, relief="solid", padx=10, pady=8)
        self._sig_txt.insert("1.0",
            "Golden Key POS\n"
            "Tel: (xxx) xxx-xxxx\n"
            "Email: info@goldenkeypos.com\n"
            "www.goldenkeypos.com\n"
            "\n"
            "─────────────────────────")
        self._sig_txt.pack(fill="x")

    def _ld_rtfd(self):
        p = filedialog.askopenfilename(title="Select .rtfd",
                filetypes=[("RTFD","*.rtfd"),("All","*.*")])
        if not p: return
        txt, err = load_rtfd(p)
        if err: messagebox.showerror("Error", err); return
        self._editor.set_text(txt)

    def _ld_txt(self):
        p = filedialog.askopenfilename(title="Select .txt",
                filetypes=[("Text","*.txt"),("All","*.*")])
        if not p: return
        try: self._editor.set_text(Path(p).read_text(encoding="utf-8"))
        except Exception as e: messagebox.showerror("Error", str(e))

    def _ld_url(self):
        """Fetch HTML from URL and use as email body"""
        win = tk.Toplevel(self)
        win.title("Load Content from URL")
        win.geometry("620x200")
        win.configure(bg=WHITE)
        win.grab_set()

        tk.Label(win, bg=WHITE, font=F_SM, anchor="w",
                text="Enter the URL of your newsletter / webpage:"
                ).pack(fill="x", padx=24, pady=(20,6))

        url_var = tk.StringVar()
        url_entry = tk.Entry(win, textvariable=url_var, font=F_NORM, bd=1, relief="solid")
        url_entry.pack(fill="x", padx=24)
        url_entry.focus()

        info = tk.Label(win, bg=WHITE, fg=MUTED, font=F_XS, anchor="w",
                       text="The page's HTML will be used directly as the email body. "
                            "{{VARIABLES}} inside the HTML will still be replaced per recipient.")
        info.pack(fill="x", padx=24, pady=(6,0))

        def fetch():
            url = url_var.get().strip()
            if not url.startswith("http"):
                messagebox.showwarning("Invalid URL","URL must start with http:// or https://")
                return
            try:
                import urllib.request
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=15) as r:
                    html = r.read().decode("utf-8", errors="replace")
                self._html_override = html
                self._editor.set_text(
                    f"✅ HTML content loaded from:\n{url}\n\n"
                    "This page's HTML will be sent as the email body.\n"
                    "{{VARIABLES}} in the HTML will still be replaced per recipient.\n\n"
                    "Click '✕ Clear' or '🗑 Clear Body' to go back to the text editor.")
                self._url_bar.pack(fill="x", pady=(4,0))
                self._url_lbl.config(text=f"🌐 HTML from: {url}")
                win.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Could not fetch URL:\n{e}")

        tk.Button(win, text="🌐  Fetch & Load into Email Body", font=F_BOLD,
                 bg=BLUE, fg=WHITE, relief="flat", pady=11, cursor="hand2",
                 command=fetch).pack(fill="x", padx=24, pady=14)

    def _clear_body(self):
        self._html_override = None
        self._editor.set_text("")
        self._url_bar.pack_forget()

    def _sig_toggle(self):
        if self._sig_on.get():
            self._sig_details.pack(fill="x", pady=(8,0))
        else:
            self._sig_details.pack_forget()

    # ════════════════════════════════════════════════════════
    #  TAB 4 — Send
    # ════════════════════════════════════════════════════════
    def _build_send(self, tab):
        outer = tk.Frame(tab, bg=LIGHT)
        outer.pack(fill="both", expand=True, padx=18, pady=16)

        # Options
        opt = tk.LabelFrame(outer, text="  🛠 Send Options  ", font=F_LGB,
                            bg=WHITE, fg=DARK, bd=1, relief="solid", padx=20, pady=14)
        opt.pack(fill="x", pady=(0,16))

        mr = tk.Frame(opt, bg=WHITE); mr.pack(fill="x", pady=6)
        tk.Label(mr, text="Mode:", font=F_NORM, bg=WHITE, fg=MUTED, width=8, anchor="w").pack(side="left")
        self._test_mode = tk.BooleanVar(value=True)
        tk.Radiobutton(mr, text="Test Mode  (send only to test email)",
                       variable=self._test_mode, value=True,
                       font=F_NORM, bg=WHITE).pack(side="left", padx=(0,24))
        tk.Radiobutton(mr, text="⚠  Live  (real send to all recipients)",
                       variable=self._test_mode, value=False,
                       font=F_NORM, bg=WHITE, fg=RED).pack(side="left")

        dr = tk.Frame(opt, bg=WHITE); dr.pack(fill="x", pady=6)
        tk.Label(dr, text="Delay:", font=F_NORM, bg=WHITE, fg=MUTED, width=8, anchor="w").pack(side="left")
        self._delay = tk.StringVar(value="2.0")
        ttk.Combobox(dr, textvariable=self._delay,
                     values=["1.0","1.5","2.0","3.0"], width=7,
                     state="readonly", font=F_NORM).pack(side="left")
        tk.Label(dr, text="seconds between emails  (spam prevention)",
                font=F_SM, bg=WHITE, fg=MUTED).pack(side="left", padx=12)

        # Action buttons
        br = tk.Frame(outer, bg=LIGHT); br.pack(fill="x", pady=(0,16))
        tk.Button(br, text="👁  Preview First Email", font=F_LGB,
                  bd=1, relief="solid", bg=WHITE, cursor="hand2", pady=14,
                  command=self._preview).pack(side="left", fill="x", expand=True, padx=(0,12))
        self._sbtn = tk.Button(br, text="✉  Start Sending", font=F_LGB,
                               bg=GREEN, fg=WHITE, relief="flat", cursor="hand2", pady=14,
                               command=self._start)
        self._sbtn.pack(side="left", fill="x", expand=True)

        # Progress
        prg = tk.LabelFrame(outer, text="  📊 Progress  ", font=F_LGB,
                            bg=WHITE, fg=DARK, bd=1, relief="solid", padx=20, pady=14)
        prg.pack(fill="x", pady=(0,16))

        self._slbl = tk.Label(prg, text="Ready.", font=F_NORM, bg=WHITE, fg=MUTED, anchor="w")
        self._slbl.pack(fill="x", pady=(0,8))
        self._pvar = tk.DoubleVar()
        ttk.Progressbar(prg, variable=self._pvar, maximum=100).pack(fill="x", pady=(0,14))

        sr2 = tk.Frame(prg, bg=WHITE); sr2.pack(fill="x")
        self._st  = self._sbox(sr2, "Total",   DARK)
        self._ss  = self._sbox(sr2, "Success", GREEN)
        self._sf  = self._sbox(sr2, "Failed",  RED)

        # Log
        self._log = scrolledtext.ScrolledText(outer, height=9, font=F_MONO,
                                               bg=DARK, fg="#a8ff78",
                                               bd=0, state="disabled", wrap="word")
        self._log.pack(fill="both", expand=True)
        self._log.tag_config("ok",   foreground="#a8ff78")
        self._log.tag_config("err",  foreground="#ff6b6b")
        self._log.tag_config("info", foreground="#87ceeb")

    def _sbox(self, parent, label, color):
        b = tk.Frame(parent, bg=LIGHT, bd=1, relief="solid")
        b.pack(side="left", fill="x", expand=True, padx=4)
        v = tk.Label(b, text="0", font=("Arial",28,"bold"), bg=LIGHT, fg=color)
        v.pack(pady=(12,3))
        tk.Label(b, text=label, font=F_SM, bg=LIGHT, fg=MUTED).pack(pady=(0,10))
        return v

    # ── Config sync ───────────────────────────────────────
    def _populate(self):
        c = self.cfg
        self._srv.set(c["server_type"])
        self._g_email.insert(0, c["gmail_email"])
        self._g_pass.insert(0,  c["gmail_pass"])
        self._smtp_host.insert(0, c["smtp_host"])
        self._smtp_port.insert(0, c["smtp_port"])
        self._ssl_var.set(c["smtp_ssl"])
        self._smtp_user.insert(0, c["smtp_user"])
        self._smtp_pass.insert(0, c["smtp_pass"])
        self._from_name.insert(0, c["from_name"])
        self._test_email.insert(0, c["test_email"])
        self._xls.set(c["excel_file"])
        self._pdf_col.set(c["pdf_col"])
        self._pdf_dir.set(c["pdf_folder"])
        self._faq.set(c["faq_file"])
        self._subj.set(c["subject"])
        self._editor.set_text(c["body"])
        self._sig_on.set(c["sig_on"])
        self._sig_logo.set(c["sig_logo"])
        self._sig_txt.insert("1.0", c["sig_text"])
        self._test_mode.set(c["test_mode"])
        self._delay.set(c["delay"])
        if c["excel_file"]:
            try: self._load_cols()
            except: pass

    def _sync(self):
        self.cfg.update({
            "server_type": self._srv.get(),
            "gmail_email": self._g_email.get().strip(),
            "gmail_pass":  self._g_pass.get(),
            "smtp_host":   self._smtp_host.get().strip(),
            "smtp_port":   self._smtp_port.get().strip(),
            "smtp_ssl":    self._ssl_var.get(),
            "smtp_user":   self._smtp_user.get().strip(),
            "smtp_pass":   self._smtp_pass.get(),
            "from_name":   self._from_name.get().strip(),
            "test_email":  self._test_email.get().strip(),
            "excel_file":  self._xls.get().strip(),
            "pdf_col":     self._pdf_col.get().strip(),
            "pdf_folder":  self._pdf_dir.get().strip(),
            "faq_file":    self._faq.get().strip(),
            "subject":     self._subj.get().strip(),
            "body":        self._editor.get_plain(),
            "sig_on":      self._sig_on.get(),
            "sig_logo":    self._sig_logo.get().strip(),
            "sig_text":    self._sig_txt.get("1.0","end-1c"),
            "test_mode":   self._test_mode.get(),
            "delay":       self._delay.get(),
        })

    def _ok(self):
        self._sync(); c = self.cfg
        checks = [
            (c["server_type"]=="gmail" and not c["gmail_email"], "Enter Gmail address in Setup."),
            (c["server_type"]=="gmail" and not c["gmail_pass"],  "Enter Gmail App Password in Setup."),
            (c["server_type"]=="custom" and not c["smtp_host"],  "Enter SMTP Host in Setup."),
            (c["server_type"]=="custom" and not c["smtp_pass"],  "Enter SMTP Password in Setup."),
            (not c["excel_file"],  "Select Excel file in Files & Recipients."),
            (not c["subject"],     "Enter Subject in Email Content."),
            (not c["body"],        "Enter email body in Email Content."),
        ]
        for cond, msg in checks:
            if cond: messagebox.showwarning("Missing", msg); return False
        return True

    # ── Preview ──────────────────────────────────────────
    def _preview(self):
        if not self._ok(): return
        rows, cols, err = load_merchants(self.cfg["excel_file"], self.cfg["pdf_col"])
        if err: messagebox.showerror("Error", err); return
        if not rows: messagebox.showwarning("Empty","No recipients found."); return

        row  = rows[0]
        subj = apply_vars(self.cfg["subject"], row)
        body = apply_vars(self.cfg["body"], row) if not self._html_override else \
               f"[HTML from URL — opens correctly in email client]\n\nFirst 500 chars:\n" + \
               self._html_override[:500] + "..."

        win = tk.Toplevel(self); win.title("Preview"); win.geometry("720x640")
        win.configure(bg=WHITE); win.grab_set()
        tk.Label(win, text="Preview — First Recipient", font=F_LGB, bg=WHITE
                 ).pack(fill="x", padx=24, pady=(20,8))
        for k,v in [("To:", row.get("EMAIL","")),
                    ("Total:", f"{len(rows)} recipients"),
                    ("Subject:", subj)]:
            fr = tk.Frame(win, bg=LIGHT); fr.pack(fill="x")
            tk.Label(fr, text=k, font=F_BOLD, bg=LIGHT, width=10, anchor="w",
                    padx=14, pady=7).pack(side="left")
            tk.Label(fr, text=v, font=F_NORM, bg=LIGHT, anchor="w").pack(side="left")
        txt = scrolledtext.ScrolledText(win, font=F_NORM, wrap="word",
                                         bd=1, relief="solid", padx=14, pady=12)
        txt.pack(fill="both", expand=True, padx=24, pady=14)
        txt.insert("1.0", body); txt.config(state="disabled")
        tk.Button(win, text="Close", font=F_LGB, command=win.destroy,
                  bg=DARK, fg=WHITE, relief="flat", pady=12
                  ).pack(fill="x", padx=24, pady=(0,20))

    # ── Log helper ───────────────────────────────────────
    def _logmsg(self, msg, tag="ok"):
        self._log.config(state="normal")
        self._log.insert("end", msg+"\n", tag)
        self._log.see("end")
        self._log.config(state="disabled")

    # ── Send ─────────────────────────────────────────────
    def _start(self):
        if self._run: messagebox.showinfo("Busy","Already sending."); return
        if not self._ok(): return
        rows, _, err = load_merchants(self.cfg["excel_file"], self.cfg["pdf_col"])
        if err: messagebox.showerror("Error", err); return
        if not rows: messagebox.showwarning("Empty","No recipients."); return

        mode = "TEST MODE" if self.cfg["test_mode"] else "⚠ LIVE — REAL SEND"
        if not messagebox.askyesno("Confirm",
                f"[{mode}]\n\nReady to send to {len(rows)} recipients.\n\nContinue?"):
            return

        self._run = True
        self._sbtn.config(state="disabled", text="Sending…")
        self._log.config(state="normal"); self._log.delete("1.0","end")
        self._log.config(state="disabled")
        self._pvar.set(0)
        self._st.config(text=str(len(rows)))
        self._ss.config(text="0"); self._sf.config(text="0")

        # Build signature HTML
        sig_html = ""
        if self._sig_on.get():
            logo_path = self._sig_logo.get()
            if logo_path and Path(logo_path).exists():
                try:
                    ext = Path(logo_path).suffix.lower().strip(".")
                    with open(logo_path,"rb") as f:
                        b64 = base64.b64encode(f.read()).decode()
                    sig_html += (f'<img src="data:image/{ext};base64,{b64}" '
                                f'style="max-height:80px;max-width:300px;display:block;margin-bottom:10px">')
                except: pass
            st = self._sig_txt.get("1.0","end-1c")
            if st: sig_html += st.replace("\n","<br>")

        html_tpl  = self._html_override if self._html_override else self._editor.get_html()
        plain_tpl = self._editor.get_plain() if not self._html_override else "Please view this email in an HTML-capable email client."

        threading.Thread(target=self._worker,
                        args=(rows, dict(self.cfg), html_tpl, plain_tpl, sig_html),
                        daemon=True).start()

    def _worker(self, rows, cfg, html_tpl, plain_tpl, sig_html):
        total = len(rows); success = 0; failed = []
        delay = float(cfg.get("delay","2.0"))
        test  = cfg["test_mode"]; test_to = cfg["test_email"]

        if test:
            self.after(0, self._logmsg, f"★ TEST MODE → {test_to}", "info")

        for i, row in enumerate(rows, 1):
            to    = test_to if test else row.get("EMAIL","")
            plain = apply_vars(plain_tpl, row)
            if self._html_override:
                html = apply_vars(self._html_override, row)   # use fetched HTML as-is
            else:
                html  = wrap_html(apply_vars(html_tpl, row), sig_html)
            name  = str(row.get(cfg["pdf_col"].upper(),"")).strip()

            ok = send_one(cfg, row, to, plain, html)

            if ok:
                success += 1
                self.after(0, self._logmsg,
                          f"✓  [{i:>3}/{total}]  {name}  <{to}>", "ok")
            else:
                failed.append(row)
                self.after(0, self._logmsg,
                          f"✗  [{i:>3}/{total}]  {name}  <{row.get('EMAIL','')}> — FAILED", "err")

            self.after(0, self._upd, i, total, success, len(failed))
            if i < total: time.sleep(delay)

        self.after(0, self._done, success, failed, total)

    def _upd(self, cur, tot, succ, fail):
        pct = cur/tot*100
        self._pvar.set(pct)
        self._slbl.config(text=f"Sending…  {cur}/{tot}  ({pct:.0f}%)")
        self._ss.config(text=str(succ)); self._sf.config(text=str(fail))

    def _done(self, succ, failed, tot):
        self._run = False
        self._sbtn.config(state="normal", text="✉  Start Sending")
        self._pvar.set(100)
        msg = f"✅ Done!  {succ} sent  /  {len(failed)} failed  (total {tot})"
        self._slbl.config(text=msg)
        self._logmsg("─"*50, "info"); self._logmsg(msg, "info")
        if failed:
            for r in failed:
                self._logmsg(f"  ✗  {r.get(self.cfg['pdf_col'].upper(),'')}  <{r.get('EMAIL','')}>", "err")
        messagebox.showinfo("Complete",
            f"Sending complete!\n\n✓ Success:  {succ}\n✗ Failed:   {len(failed)}")


if __name__ == "__main__":
    App().mainloop()
