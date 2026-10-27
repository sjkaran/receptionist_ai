"""Main window: header status, tabbed layout, PIN gate, thread-safe UI queue."""
from __future__ import annotations

import queue
import threading
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from . import widgets as w
from .knowledge_tab import KnowledgeTab
from .messages_tab import MessagesTab
from .reception_tab import ReceptionTab
from .review_tab import ReviewTab
from .settings_tab import SettingsTab
from .training_tab import TrainingTab


class App(tk.Tk):
    def __init__(self, ctx):
        super().__init__()
        self.ctx = ctx
        self.unlocked = False
        self._q: queue.Queue = queue.Queue()
        self.title(f"{ctx.settings.assistant_name} — {ctx.settings.office_name}")
        self.geometry("1120x740")
        self.minsize(940, 620)
        self.fam = w.setup_style(self)

        self._build_header()
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        self.reception = ReceptionTab(self.nb, self)
        self.nb.add(self.reception, text="Reception")
        self._protected: dict[str, tuple[ttk.Frame, str]] = {}
        for key, label, cls in (("training", "Training", TrainingTab), ("review", "Review", ReviewTab),
                                ("knowledge", "Knowledge", KnowledgeTab), ("messages", "Messages", MessagesTab),
                                ("settings", "Settings", SettingsTab)):
            tab = cls(self.nb, self)
            self.nb.add(tab, text=label)
            self._protected[str(tab)] = (tab, label)
            setattr(self, key, tab)

        self.nb.bind("<<NotebookTabChanged>>", self._on_tab)
        self.protocol("WM_DELETE_WINDOW", self._close)
        self.after(40, self._pump)
        self.refresh_status()
        self.refresh_badges()
        self.after(15000, self._poll)

    # ------------------------------------------------------------------ header
    def _build_header(self) -> None:
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=18, pady=(14, 8))
        s = self.ctx.settings
        self.title_var = tk.StringVar(value=f"{s.assistant_name}  ·  {s.office_name}")
        ttk.Label(bar, textvariable=self.title_var, style="H1.TLabel").pack(side="left")
        self.lock_btn = ttk.Button(bar, text="Lock", command=self.lock)
        self.lock_btn.pack(side="right")
        self.lock_btn.state(["disabled"])
        self.status_var = tk.StringVar(value="Checking Ollama…")
        self.status_lbl = tk.Label(bar, textvariable=self.status_var, bg=w.BG, fg=w.MUTED, font=(self.fam, 9))
        self.status_lbl.pack(side="right", padx=14)

    def refresh_titles(self) -> None:
        s = self.ctx.settings
        self.title(f"{s.assistant_name} — {s.office_name}")
        self.title_var.set(f"{s.assistant_name}  ·  {s.office_name}")

    # ---------------------------------------------------------- thread plumbing
    def ui(self, fn) -> None:
        """Schedule `fn` on the Tk thread (safe from any thread)."""
        self._q.put(fn)

    def _pump(self) -> None:
        try:
            while True:
                self._q.get_nowait()()
        except queue.Empty:
            pass
        except Exception as exc:             # a UI callback must never kill the pump
            print(f"[ui] callback error: {exc!r}")
            self.after(40, self._pump)
            return
        self.after(40, self._pump)

    def run_bg(self, work, ok=None, err=None) -> None:
        def runner():
            try:
                result = work()
            except Exception as exc:
                if err:
                    self.ui(lambda e=exc: err(e))
                else:
                    print(f"[bg] {exc!r}")
                return
            if ok:
                self.ui(lambda r=result: ok(r))
        threading.Thread(target=runner, daemon=True).start()

    # ------------------------------------------------------------------- status
    def refresh_status(self) -> None:
        llm = self.ctx.llm

        def work():
            llm.reset()
            return llm.resolve_model()

        def ok(model):
            sem = "semantic on" if self.ctx.retriever.semantic_active else "keyword + fuzzy"
            self.status_var.set(f"●  Ollama · {model} · {sem}")
            self.status_lbl.configure(fg=w.OK)

        def bad(exc):
            self.status_var.set("●  Ollama offline — trained answers only")
            self.status_lbl.configure(fg=w.BAD)
            self._last_llm_error = str(exc)

        self.run_bg(work, ok, bad)

    def _poll(self) -> None:
        self.refresh_status()
        self.refresh_badges()
        self.after(15000, self._poll)

    def refresh_badges(self) -> None:
        def work():
            return self.ctx.db.pending_count(), self.ctx.db.unhandled_count()

        def ok(counts):
            pending, msgs = counts
            for tab, base in self._protected.values():
                if base == "Review":
                    self.nb.tab(tab, text=f"Review ({pending})" if pending else "Review")
                elif base == "Messages":
                    self.nb.tab(tab, text=f"Messages ({msgs})" if msgs else "Messages")
        self.run_bg(work, ok)

    # ------------------------------------------------------------------ PIN gate
    def unlock(self) -> bool:
        if self.unlocked:
            return True
        pin = simpledialog.askstring("Trainer access", "Enter trainer PIN:", show="•", parent=self)
        if pin is not None and self.ctx.settings.check_pin(pin.strip()):
            self.unlocked = True
            self.lock_btn.state(["!disabled"])
            return True
        if pin is not None:
            messagebox.showerror("Access denied", "Incorrect PIN.", parent=self)
        return False

    def lock(self) -> None:
        self.unlocked = False
        self.lock_btn.state(["disabled"])
        self.nb.select(self.reception)

    def _on_tab(self, _event=None) -> None:
        tab_id = self.nb.select()
        if tab_id in self._protected:
            if not self.unlock():
                self.nb.select(self.reception)
                return
            tab = self._protected[tab_id][0]
            if hasattr(tab, "on_show"):
                tab.on_show()

    # ---------------------------------------------------------------------- exit
    def _close(self) -> None:
        try:
            self.ctx.session.stop()
            self.ctx.speaker.stop()
        finally:
            self.destroy()
