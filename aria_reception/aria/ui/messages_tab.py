"""Messages tab: visitor messages Aria collected (name, contact, message)."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from . import widgets as w


class MessagesTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=14)
        self.app, self.ctx = app, app.ctx
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(top, text="Messages left for the team", style="H2.TLabel").pack(side="left")
        self.show_all = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Show handled", variable=self.show_all, command=self.refresh).pack(side="right")

        frame, self.tree = w.make_tree(
            self, [("when", "Received", 140), ("who", "Visitor", 130), ("contact", "Contact", 170),
                   ("body", "Message", 400), ("done", "Done", 50)], stretch="body")
        frame.grid(row=1, column=0, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        self.detail = tk.Text(self, height=4, wrap="word", state="disabled", relief="flat", bd=0, padx=10, pady=8,
                              bg=w.CARD, fg=w.INK, highlightthickness=1, highlightbackground=w.LINE,
                              font=(app.fam, 10))
        self.detail.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        btns = ttk.Frame(self)
        btns.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(btns, text="Mark handled", style="Accent.TButton", command=lambda: self.mark(True)).pack(side="left")
        ttk.Button(btns, text="Reopen", command=lambda: self.mark(False)).pack(side="left", padx=8)
        ttk.Button(btns, text="Copy contact", command=self.copy_contact).pack(side="left")
        ttk.Button(btns, text="Refresh", command=self.refresh).pack(side="right")
        self.rows: dict[str, dict] = {}
        self.refresh()

    def on_show(self) -> None:
        self.refresh()

    def refresh(self) -> None:
        rows = self.ctx.db.list_messages(self.show_all.get())
        self.rows = {str(r["id"]): r for r in rows}
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert("", "end", iid=str(r["id"]), tags=("off",) if r["handled"] else (),
                             values=(r["created"], r["visitor"], r["contact"], r["body"], "yes" if r["handled"] else ""))
        self._show("")
        self.app.refresh_badges()

    def _show(self, text: str) -> None:
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        self.detail.insert("1.0", text)
        self.detail.configure(state="disabled")

    def _selected(self) -> dict | None:
        s = self.tree.selection()
        return self.rows.get(s[0]) if s else None

    def _on_select(self, _e=None) -> None:
        r = self._selected()
        if r:
            self._show(f"{r['visitor']}  ·  {r['contact']}\n\n{r['body']}")

    def mark(self, handled: bool) -> None:
        r = self._selected()
        if r:
            self.ctx.db.set_message_handled(r["id"], handled)
            self.refresh()

    def copy_contact(self) -> None:
        r = self._selected()
        if r:
            self.clipboard_clear()
            self.clipboard_append(r["contact"])
