"""Knowledge tab: office facts (ground truth for the AI), greetings, import / export."""
from __future__ import annotations

import json
from tkinter import filedialog, messagebox, ttk

from . import widgets as w


class KnowledgeTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=14)
        self.app, self.ctx = app, app.ctx
        self.fact_id: int | None = None
        self.columnconfigure(0, weight=3, uniform="c")
        self.columnconfigure(1, weight=2, uniform="c")
        self.rowconfigure(0, weight=1)

        # ---- facts
        fbox = ttk.Frame(self)
        fbox.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        fbox.rowconfigure(2, weight=1)
        fbox.columnconfigure(0, weight=1)
        ttk.Label(fbox, text="Office facts", style="H2.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(fbox, text="Short true statements. The AI may use these (and only these) to answer new questions.",
                  style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(0, 8))
        frame, self.ftree = w.make_tree(fbox, [("l", "Label", 150), ("v", "Value", 300)], stretch="v")
        frame.grid(row=2, column=0, sticky="nsew")
        self.ftree.bind("<<TreeviewSelect>>", self._pick_fact)
        form = ttk.Frame(fbox)
        form.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        form.columnconfigure(1, weight=1)
        ttk.Label(form, text="Label", style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        self.f_label = ttk.Entry(form)
        self.f_label.grid(row=0, column=1, sticky="ew", padx=8, pady=2, ipady=2)
        ttk.Label(form, text="Value", style="Muted.TLabel").grid(row=1, column=0, sticky="w")
        self.f_value = ttk.Entry(form)
        self.f_value.grid(row=1, column=1, sticky="ew", padx=8, pady=2, ipady=2)
        fb = ttk.Frame(fbox)
        fb.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(fb, text="Add", style="Accent.TButton", command=self.add_fact).pack(side="left")
        ttk.Button(fb, text="Update selected", command=self.update_fact).pack(side="left", padx=6)
        ttk.Button(fb, text="Delete", style="Danger.TButton", command=self.del_fact).pack(side="right")

        # ---- greetings + data
        gbox = ttk.Frame(self)
        gbox.grid(row=0, column=1, sticky="nsew")
        gbox.rowconfigure(2, weight=1)
        gbox.columnconfigure(0, weight=1)
        ttk.Label(gbox, text="Greetings", style="H2.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(gbox, text="One is picked at random each time reception starts.",
                  style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(0, 8))
        gframe, self.gtree = w.make_tree(gbox, [("t", "Greeting", 300)], height=6)
        gframe.grid(row=2, column=0, sticky="nsew")
        self.g_text = w.make_text(gbox, height=3, font=(app.fam, 10))
        self.g_text.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        gb = ttk.Frame(gbox)
        gb.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(gb, text="Add greeting", style="Accent.TButton", command=self.add_greeting).pack(side="left")
        ttk.Button(gb, text="Delete selected", style="Danger.TButton", command=self.del_greeting).pack(side="right")

        dbox = ttk.LabelFrame(gbox, text="BACKUP", padding=10)
        dbox.grid(row=5, column=0, sticky="ew", pady=(18, 0))
        ttk.Button(dbox, text="Export training…", command=self.export).pack(side="left")
        ttk.Button(dbox, text="Import training…", command=self.import_).pack(side="left", padx=8)

        self.refresh()

    def on_show(self) -> None:
        self.refresh()

    def refresh(self) -> None:
        db = self.ctx.db
        self.ftree.delete(*self.ftree.get_children())
        for f in db.list_facts():
            self.ftree.insert("", "end", iid=str(f["id"]), values=(f["label"], f["value"]))
        self.gtree.delete(*self.gtree.get_children())
        for g in db.list_greetings():
            self.gtree.insert("", "end", iid=str(g["id"]), values=(g["text"],))

    # ------------------------------------------------------------------ facts
    def _pick_fact(self, _e=None) -> None:
        s = self.ftree.selection()
        if not s:
            return
        self.fact_id = int(s[0])
        label, value = self.ftree.item(s[0], "values")
        self.f_label.delete(0, "end")
        self.f_label.insert(0, label)
        self.f_value.delete(0, "end")
        self.f_value.insert(0, value)

    def add_fact(self) -> None:
        try:
            self.ctx.db.add_fact(self.f_label.get(), self.f_value.get())
        except ValueError as exc:
            messagebox.showwarning("Cannot save", str(exc), parent=self)
            return
        self.f_label.delete(0, "end")
        self.f_value.delete(0, "end")
        self.refresh()

    def update_fact(self) -> None:
        if self.fact_id is None or not self.f_label.get().strip() or not self.f_value.get().strip():
            return
        self.ctx.db.update_fact(self.fact_id, self.f_label.get(), self.f_value.get())
        self.refresh()

    def del_fact(self) -> None:
        s = self.ftree.selection()
        if s and messagebox.askyesno("Delete", "Delete this fact?", parent=self):
            self.ctx.db.delete_fact(int(s[0]))
            self.refresh()

    # --------------------------------------------------------------- greetings
    def add_greeting(self) -> None:
        try:
            self.ctx.db.add_greeting(w.get_text(self.g_text))
        except ValueError as exc:
            messagebox.showwarning("Cannot save", str(exc), parent=self)
            return
        w.set_text(self.g_text, "")
        self.refresh()

    def del_greeting(self) -> None:
        s = self.gtree.selection()
        if s:
            self.ctx.db.delete_greeting(int(s[0]))
            self.refresh()

    # ------------------------------------------------------------ import/export
    def export(self) -> None:
        path = filedialog.asksaveasfilename(parent=self, defaultextension=".json", initialfile="aria_training.json",
                                            filetypes=[("JSON", "*.json")])
        if path:
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(self.ctx.db.export_data(), fh, indent=2, ensure_ascii=False)
            messagebox.showinfo("Exported", f"Saved to {path}", parent=self)

    def import_(self) -> None:
        path = filedialog.askopenfilename(parent=self, filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as fh:
                res = self.ctx.db.import_data(json.load(fh))
        except (OSError, ValueError) as exc:
            messagebox.showerror("Import failed", str(exc), parent=self)
            return
        self.refresh()
        self.app.training.refresh()
        messagebox.showinfo("Imported", f"Added {res['added']} Q&A, skipped {res['skipped']} duplicates.", parent=self)
