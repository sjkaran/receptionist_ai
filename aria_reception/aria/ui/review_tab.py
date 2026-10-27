"""Review tab: the learning loop.  Questions Aria could not answer from training (or answered with
the local AI) wait here; approve them with a good answer and she knows them next time."""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from . import widgets as w


class ReviewTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=14)
        self.app, self.ctx = app, app.ctx
        self.sel: dict | None = None
        self.items: dict[str, dict] = {}
        self.columnconfigure(0, weight=5, uniform="c")
        self.columnconfigure(1, weight=6, uniform="c")
        self.rowconfigure(0, weight=1)

        left = ttk.Frame(self)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)
        ttk.Label(left, text="Visitors asked these — most frequent first", style="Muted.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 8))
        frame, self.tree = w.make_tree(left, [("n", "Asked", 52), ("q", "Question", 300), ("ai", "AI replied", 70)],
                                       stretch="q")
        frame.grid(row=1, column=0, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.empty = tk.StringVar()
        ttk.Label(left, textvariable=self.empty, style="Muted.TLabel").grid(row=2, column=0, sticky="w", pady=(6, 0))

        right = ttk.Frame(self)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        ttk.Label(right, text="Teach Aria", style="H2.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(right, text="Question (edit to a clean, general form)", style="Muted.TLabel").grid(
            row=1, column=0, sticky="w", pady=(8, 2))
        self.q_entry = ttk.Entry(right)
        self.q_entry.grid(row=2, column=0, sticky="ew", ipady=3)
        ttk.Label(right, text="Answer  (pre-filled with what the AI said — check it carefully)",
                  style="Muted.TLabel").grid(row=3, column=0, sticky="w", pady=(10, 2))
        self.a_text = w.make_text(right, height=6, font=(app.fam, 10))
        self.a_text.grid(row=4, column=0, sticky="ew")
        crow = ttk.Frame(right)
        crow.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(crow, text="Category", style="Muted.TLabel").pack(side="left")
        self.cat = ttk.Entry(crow, width=24)
        self.cat.pack(side="left", padx=8, ipady=2)
        btns = ttk.Frame(right)
        btns.grid(row=6, column=0, sticky="ew", pady=(14, 0))
        ttk.Button(btns, text="Approve and train", style="Accent.TButton", command=self.approve).pack(side="left")
        ttk.Button(btns, text="Dismiss", command=self.dismiss).pack(side="left", padx=8)
        ttk.Button(btns, text="Refresh", command=self.refresh).pack(side="right")

    def on_show(self) -> None:
        self.refresh()

    def refresh(self) -> None:
        rows = self.ctx.db.list_review()
        self.items = {str(r["id"]): r for r in rows}
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert("", "end", iid=str(r["id"]),
                             values=(r["count"], r["question"], "yes" if r["suggested"] else "—"))
        self.empty.set("" if rows else "Nothing waiting. Unanswered questions will appear here.")
        self._clear()
        if rows:
            self.tree.selection_set(str(rows[0]["id"]))
        self.app.refresh_badges()

    def _clear(self) -> None:
        self.sel = None
        self.q_entry.delete(0, "end")
        self.cat.delete(0, "end")
        w.set_text(self.a_text, "")

    def _on_select(self, _e=None) -> None:
        s = self.tree.selection()
        if not s or s[0] not in self.items:
            return
        self.sel = self.items[s[0]]
        self.q_entry.delete(0, "end")
        self.q_entry.insert(0, self.sel["question"])
        w.set_text(self.a_text, self.sel["suggested"])

    def approve(self) -> None:
        if not self.sel:
            return
        q, a = self.q_entry.get().strip(), w.get_text(self.a_text)
        if not q or not a:
            messagebox.showwarning("Missing text", "Please provide the question and the answer to teach.", parent=self)
            return
        try:
            self.ctx.db.approve_review(self.sel["id"], q, a, self.cat.get().strip())
        except ValueError as exc:
            messagebox.showwarning("Cannot save", str(exc), parent=self)
            return
        self.refresh()

    def dismiss(self) -> None:
        if self.sel:
            self.ctx.db.dismiss_review(self.sel["id"])
            self.refresh()
