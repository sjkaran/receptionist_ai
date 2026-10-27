"""Training tab: create/edit Q&A with alternate phrasings, AI helpers, and a match tester."""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from ..textutils import normalize
from . import widgets as w


class TrainingTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=14)
        self.app, self.ctx = app, app.ctx
        self.sel_id: int | None = None
        self.columnconfigure(0, weight=5, uniform="c")
        self.columnconfigure(1, weight=6, uniform="c")
        self.rowconfigure(0, weight=1)

        # ---- left: list
        left = ttk.Frame(self)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)
        self.search = tk.StringVar()
        self.search.trace_add("write", lambda *_: self.refresh())
        srow = ttk.Frame(left)
        srow.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        srow.columnconfigure(1, weight=1)
        ttk.Label(srow, text="Search", style="Muted.TLabel").grid(row=0, column=0, padx=(0, 8))
        ttk.Entry(srow, textvariable=self.search).grid(row=0, column=1, sticky="ew", ipady=3)
        frame, self.tree = w.make_tree(left, [("q", "Question", 260), ("cat", "Category", 80),
                                              ("hits", "Used", 46), ("on", "On", 34)], stretch="q")
        frame.grid(row=1, column=0, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.count_var = tk.StringVar()
        ttk.Label(left, textvariable=self.count_var, style="Muted.TLabel").grid(row=2, column=0, sticky="w", pady=(6, 0))

        # ---- right: form
        right = ttk.Frame(self)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)

        ttk.Label(right, text="When a visitor asks…", style="H2.TLabel").grid(row=0, column=0, sticky="w")
        qrow = ttk.Frame(right)
        qrow.grid(row=1, column=0, sticky="ew", pady=(4, 10))
        qrow.columnconfigure(0, weight=1)
        self.q_entry = ttk.Entry(qrow)
        self.q_entry.grid(row=0, column=0, sticky="ew", ipady=3)
        self.q_dict = ttk.Button(qrow, text="Dictate", command=lambda: self.dictate("q"))
        self.q_dict.grid(row=0, column=1, padx=(6, 0))

        ttk.Label(right, text="Aria answers", style="H2.TLabel").grid(row=2, column=0, sticky="w")
        arow = ttk.Frame(right)
        arow.grid(row=3, column=0, sticky="ew", pady=(4, 10))
        arow.columnconfigure(0, weight=1)
        self.a_text = w.make_text(arow, height=4, font=(app.fam, 10))
        self.a_text.grid(row=0, column=0, rowspan=2, sticky="ew")
        self.a_dict = ttk.Button(arow, text="Dictate", command=lambda: self.dictate("a"))
        self.a_dict.grid(row=0, column=1, padx=(6, 0), sticky="n")
        self.polish_btn = ttk.Button(arow, text="Polish", command=self.polish)
        self.polish_btn.grid(row=1, column=1, padx=(6, 0), sticky="s")

        crow = ttk.Frame(right)
        crow.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(crow, text="Category", style="Muted.TLabel").pack(side="left")
        self.cat_entry = ttk.Entry(crow, width=24)
        self.cat_entry.pack(side="left", padx=8, ipady=2)

        vhead = ttk.Frame(right)
        vhead.grid(row=5, column=0, sticky="ew")
        ttk.Label(vhead, text="Other ways to ask it  (one per line)", style="H2.TLabel").pack(side="left")
        self.var_btn = ttk.Button(vhead, text="Suggest with AI", command=self.suggest)
        self.var_btn.pack(side="right")
        self.v_text = w.make_text(right, height=6, font=(app.fam, 10))
        self.v_text.grid(row=6, column=0, sticky="ew", pady=(4, 10))

        btns = ttk.Frame(right)
        btns.grid(row=7, column=0, sticky="ew")
        self.save_btn = ttk.Button(btns, text="Save", style="Accent.TButton", command=self.save)
        self.save_btn.pack(side="left")
        ttk.Button(btns, text="New", command=self.new).pack(side="left", padx=6)
        self.toggle_btn = ttk.Button(btns, text="Disable", command=self.toggle)
        self.toggle_btn.pack(side="left")
        ttk.Button(btns, text="Delete", style="Danger.TButton", command=self.delete).pack(side="right")
        self.note = tk.StringVar()
        ttk.Label(right, textvariable=self.note, style="Muted.TLabel").grid(row=8, column=0, sticky="w", pady=(6, 0))

        tbox = ttk.LabelFrame(right, text="TEST MATCHING", padding=8)
        tbox.grid(row=9, column=0, sticky="ew", pady=(10, 0))
        tbox.columnconfigure(0, weight=1)
        self.t_entry = ttk.Entry(tbox)
        self.t_entry.grid(row=0, column=0, sticky="ew", ipady=3)
        self.t_entry.bind("<Return>", lambda _e: self.test())
        ttk.Button(tbox, text="Test", command=self.test).grid(row=0, column=1, padx=(6, 0))
        self.t_out = tk.Text(tbox, height=6, wrap="word", state="disabled", relief="flat", bd=0, bg=w.BG,
                             fg=w.INK, font=(app.fam, 9))
        self.t_out.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        self.refresh()
        self.new()

    # ---------------------------------------------------------------- list
    def on_show(self) -> None:
        self.refresh()

    def refresh(self) -> None:
        rows = self.ctx.db.list_qa(self.search.get())
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert("", "end", iid=str(r["id"]), tags=() if r["enabled"] else ("off",),
                             values=(r["question"], r["category"], r["hits"], "yes" if r["enabled"] else "no"))
        self.count_var.set(f"{len(rows)} trained question{'s' if len(rows) != 1 else ''}")
        if self.sel_id is not None and self.tree.exists(str(self.sel_id)):
            self.tree.selection_set(str(self.sel_id))

    def _on_select(self, _e=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        qa = self.ctx.db.get_qa(int(sel[0]))
        if not qa:
            return
        self.sel_id = qa["id"]
        self.q_entry.delete(0, "end")
        self.q_entry.insert(0, qa["question"])
        w.set_text(self.a_text, qa["answer"])
        self.cat_entry.delete(0, "end")
        self.cat_entry.insert(0, qa["category"])
        w.set_text(self.v_text, "\n".join(qa["variants"]))
        self.toggle_btn.configure(text="Disable" if qa["enabled"] else "Enable")
        self.save_btn.configure(text="Update")
        self.note.set(f"Used {qa['hits']} time{'s' if qa['hits'] != 1 else ''}.")

    def new(self) -> None:
        self.sel_id = None
        self.tree.selection_remove(*self.tree.selection())
        self.q_entry.delete(0, "end")
        self.cat_entry.delete(0, "end")
        w.set_text(self.a_text, "")
        w.set_text(self.v_text, "")
        self.save_btn.configure(text="Save")
        self.toggle_btn.configure(text="Disable")
        self.note.set("")
        self.q_entry.focus_set()

    # ---------------------------------------------------------------- actions
    def _form(self):
        variants = [ln.strip() for ln in w.get_text(self.v_text).splitlines() if ln.strip()]
        return self.q_entry.get().strip(), w.get_text(self.a_text), self.cat_entry.get().strip(), variants

    def save(self) -> None:
        q, a, cat, variants = self._form()
        if not q or not a:
            messagebox.showwarning("Missing text", "Please enter both a question and an answer.", parent=self)
            return
        db = self.ctx.db
        try:
            if self.sel_id is None:
                near = self.ctx.retriever.best(q)
                if near and near.score >= 0.8 and normalize(near.question) != normalize(q):
                    if not messagebox.askyesno(
                            "Very similar entry",
                            f"This looks like an existing entry:\n\n“{near.question}”\n→ {near.answer}\n\n"
                            "Add it anyway?\n(Tip: add your wording as an alternate phrasing instead.)",
                            parent=self):
                        return
                self.sel_id = db.add_qa(q, a, cat, variants)
                self.note.set("Saved.")
            else:
                db.update_qa(self.sel_id, q, a, cat, variants)
                self.note.set("Updated.")
        except ValueError as exc:
            messagebox.showwarning("Cannot save", str(exc), parent=self)
            return
        self.save_btn.configure(text="Update")
        self.refresh()

    def delete(self) -> None:
        if self.sel_id is None:
            return
        if messagebox.askyesno("Delete", "Delete this trained entry?", parent=self):
            self.ctx.db.delete_qa(self.sel_id)
            self.new()
            self.refresh()

    def toggle(self) -> None:
        if self.sel_id is None:
            return
        qa = self.ctx.db.get_qa(self.sel_id)
        self.ctx.db.set_enabled(self.sel_id, not qa["enabled"])
        self.refresh()
        self._on_select()

    # ------------------------------------------------------------- AI helpers
    def _ai_error(self, exc) -> None:
        self.var_btn.state(["!disabled"])
        self.polish_btn.state(["!disabled"])
        messagebox.showerror("Local AI unavailable", str(exc), parent=self)

    def suggest(self) -> None:
        q, a, _, _ = self._form()
        if not q:
            messagebox.showinfo("Need a question", "Type the main question first.", parent=self)
            return
        self.var_btn.state(["disabled"])
        self.note.set("Asking the local model…")

        def ok(found):
            self.var_btn.state(["!disabled"])
            have = {normalize(x) for x in w.get_text(self.v_text).splitlines()}
            fresh = [f for f in found if normalize(f) not in have]
            cur = w.get_text(self.v_text)
            w.set_text(self.v_text, (cur + "\n" if cur else "") + "\n".join(fresh))
            self.note.set(f"Added {len(fresh)} suggestion(s) — review, then Save.")

        self.app.run_bg(lambda: self.ctx.brain.suggest_variants(q, a), ok, self._ai_error)

    def polish(self) -> None:
        a = w.get_text(self.a_text)
        if not a:
            return
        self.polish_btn.state(["disabled"])
        self.note.set("Polishing…")

        def ok(text):
            self.polish_btn.state(["!disabled"])
            w.set_text(self.a_text, text or a)
            self.note.set("Polished — review, then Save.")

        self.app.run_bg(lambda: self.ctx.brain.polish_answer(a), ok, self._ai_error)

    def dictate(self, which: str) -> None:
        lis = self.ctx.listener
        if not lis.available:
            messagebox.showerror("Voice unavailable", lis.error or "No microphone.", parent=self)
            return
        btn = self.q_dict if which == "q" else self.a_dict
        btn.configure(text="Listening…")
        btn.state(["disabled"])

        def finish(text=""):
            btn.configure(text="Dictate")
            btn.state(["!disabled"])
            if not text:
                self.note.set("Heard nothing — try again.")
                return
            if which == "q":
                self.q_entry.delete(0, "end")
                self.q_entry.insert(0, text)
            else:
                w.set_text(self.a_text, text)

        self.app.run_bg(lambda: lis.listen(timeout=8, phrase_limit=30), finish,
                        lambda e: (finish(), messagebox.showerror("Voice error", str(e), parent=self)))

    # ------------------------------------------------------------------- test
    def test(self) -> None:
        text = self.t_entry.get().strip()
        if not text:
            return
        st = self.ctx.settings

        def ok(matches):
            lines = []
            if not matches:
                lines.append("No trained entries yet.")
            for m in matches:
                sem = f" · meaning {m.semantic:.2f}" if m.semantic is not None else ""
                lines.append(f"{m.score:.2f}   {m.question}   (via “{m.phrase}”{sem})")
            top = matches[0].score if matches else 0
            verdict = ("Result: answered directly from training" if top >= st.match_threshold else
                       "Result: local AI, using the closest entries as reference" if top >= st.suggest_threshold else
                       "Result: local AI, using office facts only")
            lines.append("")
            lines.append(verdict)
            self.t_out.configure(state="normal")
            self.t_out.delete("1.0", "end")
            self.t_out.insert("1.0", "\n".join(lines))
            self.t_out.configure(state="disabled")

        self.app.run_bg(lambda: self.ctx.retriever.search(text, 3), ok,
                        lambda e: self.note.set(f"Test failed: {e}"))
