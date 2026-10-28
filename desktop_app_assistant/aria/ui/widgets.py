"""Theme and small reusable widgets for the minimal Tk UI."""
from __future__ import annotations

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

BG, CARD, INK, MUTED = "#f4f5f7", "#ffffff", "#1d2330", "#6b7385"
ACCENT, ACCENT_DK, OK, WARN, BAD = "#2f6df6", "#2557cc", "#1f9d63", "#d98a00", "#d64545"
LINE = "#dfe3ea"

_FAMILIES = ("Segoe UI", "SF Pro Text", "Helvetica Neue", "Inter", "DejaVu Sans", "Arial")


def pick_font(root) -> str:
    have = set(tkfont.families(root))
    return next((f for f in _FAMILIES if f in have),
                tkfont.nametofont("TkDefaultFont").actual("family"))


def setup_style(root: tk.Misc) -> str:
    fam = pick_font(root)
    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
        try:
            tkfont.nametofont(name).configure(family=fam, size=10)
        except tk.TclError:
            pass
    root.configure(bg=BG)
    st = ttk.Style(root)
    st.theme_use("clam")
    st.configure(".", background=BG, foreground=INK, fieldbackground=CARD, bordercolor=LINE,
                 lightcolor=BG, darkcolor=BG, troughcolor=LINE, font=(fam, 10))
    st.configure("TFrame", background=BG)
    st.configure("Card.TFrame", background=CARD)
    st.configure("TLabel", background=BG, foreground=INK)
    st.configure("Muted.TLabel", foreground=MUTED)
    st.configure("H1.TLabel", font=(fam, 15, "bold"))
    st.configure("H2.TLabel", font=(fam, 11, "bold"))
    st.configure("TLabelframe", background=BG, bordercolor=LINE)
    st.configure("TLabelframe.Label", background=BG, foreground=MUTED, font=(fam, 9, "bold"))
    st.configure("TButton", padding=(12, 6), background=CARD, bordercolor=LINE, focuscolor=CARD)
    st.map("TButton", background=[("active", "#eef1f6"), ("disabled", BG)], foreground=[("disabled", MUTED)])
    st.configure("Accent.TButton", background=ACCENT, foreground="white", bordercolor=ACCENT)
    st.map("Accent.TButton", background=[("active", ACCENT_DK), ("disabled", LINE)])
    st.configure("Danger.TButton", foreground=BAD)
    st.configure("TEntry", padding=6, bordercolor=LINE)
    st.configure("TCombobox", padding=5)
    st.configure("TSpinbox", padding=5)
    st.configure("TCheckbutton", background=BG)
    st.configure("TNotebook", background=BG, borderwidth=0, tabmargins=(0, 4, 0, 0))
    st.configure("TNotebook.Tab", padding=(18, 8), background=BG, foreground=MUTED, borderwidth=0)
    st.map("TNotebook.Tab", background=[("selected", CARD)], foreground=[("selected", INK)])
    st.configure("Treeview", background=CARD, fieldbackground=CARD, rowheight=28, bordercolor=LINE, borderwidth=1)
    st.configure("Treeview.Heading", background=BG, foreground=MUTED, font=(fam, 9, "bold"), padding=6)
    st.map("Treeview", background=[("selected", "#dce7ff")], foreground=[("selected", INK)])
    return fam


def make_tree(parent, columns: list[tuple[str, str, int]], height: int = 10, stretch: str | None = None):
    """columns: (id, heading, width).  Returns (frame, tree)."""
    frame = ttk.Frame(parent)
    tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings", height=height,
                        selectmode="browse")
    for cid, head, width in columns:
        tree.heading(cid, text=head, anchor="w")
        tree.column(cid, width=width, minwidth=40, anchor="w", stretch=(cid == (stretch or columns[0][0])))
    sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    tree.tag_configure("off", foreground=MUTED)
    return frame, tree


def make_text(parent, height: int = 4, font=None, **kw) -> tk.Text:
    t = tk.Text(parent, height=height, wrap="word", relief="flat", bd=0, padx=8, pady=6, bg=CARD, fg=INK,
                insertbackground=INK, highlightthickness=1, highlightbackground=LINE, highlightcolor=ACCENT,
                font=font, undo=True, **kw)
    return t


def get_text(t: tk.Text) -> str:
    return t.get("1.0", "end").strip()


def set_text(t: tk.Text, value: str) -> None:
    t.delete("1.0", "end")
    t.insert("1.0", value)


def pad(widget, **kw):
    widget.grid_configure(padx=kw.get("padx", 4), pady=kw.get("pady", 4))
    return widget
