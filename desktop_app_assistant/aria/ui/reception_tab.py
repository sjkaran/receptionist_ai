"""Reception tab: live transcript, typed input, hands-free voice mode."""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from . import widgets as w

SOURCE_LABEL = {"trained": "trained answer", "ai": "local AI", "smalltalk": "small talk",
                "dialog": "taking a message", "fallback": "not known"}


class ReceptionTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=14)
        self.app, self.ctx = app, app.ctx
        self._aria_open = False

        # bottom rows are packed first so a small window shrinks the transcript, never the controls
        ctl = ttk.Frame(self)
        ctl.pack(side="bottom", fill="x", pady=(10, 0))
        self.voice_btn = ttk.Button(ctl, text="Start voice reception", command=self.toggle_voice)
        self.voice_btn.pack(side="left")
        ttk.Button(ctl, text="Stop speaking", command=self.ctx.speaker.stop).pack(side="left", padx=8)
        ttk.Button(ctl, text="New visitor", command=self.new_visitor).pack(side="left")
        self.state_var = tk.StringVar(value="Idle")
        ttk.Label(ctl, textvariable=self.state_var, style="Muted.TLabel").pack(side="right")

        row = ttk.Frame(self)
        row.pack(side="bottom", fill="x", pady=(12, 0))
        self.entry = ttk.Entry(row, font=(app.fam, 11))
        self.entry.pack(side="left", fill="x", expand=True, ipady=3)
        self.entry.bind("<Return>", lambda _e: self.send())
        ttk.Button(row, text="Send", style="Accent.TButton", command=self.send).pack(side="left", padx=(8, 0))

        card = tk.Frame(self, bg=w.CARD, highlightthickness=1, highlightbackground=w.LINE)
        card.pack(side="top", fill="both", expand=True)
        self.log = tk.Text(card, wrap="word", state="disabled", relief="flat", bd=0, padx=18, pady=14,
                           bg=w.CARD, fg=w.INK, font=(app.fam, 11), spacing1=2, spacing3=6, cursor="arrow")
        sb = ttk.Scrollbar(card, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(side="left", fill="both", expand=True)
        self.log.tag_configure("who_v", foreground=w.ACCENT, font=(app.fam, 9, "bold"))
        self.log.tag_configure("who_a", foreground=w.OK, font=(app.fam, 9, "bold"))
        self.log.tag_configure("meta", foreground=w.MUTED, font=(app.fam, 8), lmargin1=0)
        self.log.tag_configure("err", foreground=w.BAD)

        self._hello()

    # ------------------------------------------------------------- transcript
    def _write(self, text: str, *tags: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text, tags)
        self.log.configure(state="disabled")
        self.log.see("end")

    def _hello(self) -> None:
        s = self.ctx.settings
        self._write(f"{s.assistant_name} is ready. Type below, or start voice reception.\n", "meta")
        if not self.ctx.listener.available:
            self._write("Voice input unavailable on this machine — typing still works. "
                        f"({self.ctx.listener.error})\n", "meta")

    def on_visitor(self, text: str) -> None:
        self._aria_open = False
        self._write("\nVISITOR\n", "who_v")
        self._write(text + "\n")

    def on_partial(self, sentence: str) -> None:
        if not self._aria_open:
            self._write(f"\n{self.ctx.settings.assistant_name.upper()}\n", "who_a")
            self._aria_open = True
        else:
            self._write(" ")
        self._write(sentence)

    def on_reply(self, reply) -> None:
        if self._aria_open:
            self._write("\n")
        self._aria_open = False
        if reply is None:
            return
        label = SOURCE_LABEL.get(reply.source, reply.source)
        bits = [label]
        if reply.source in ("trained", "ai", "fallback") and reply.score:
            bits.append(f"match {reply.score:.2f}")
        if reply.queued:
            bits.append("sent to Review")
        self._write("  " + "  ·  ".join(bits) + "\n", "meta")
        self.app.refresh_badges()

    def on_error(self, exc) -> None:
        self._aria_open = False
        self._write(f"\n{exc}\n", "err")

    # ------------------------------------------------------------------ typing
    def send(self) -> None:
        text = self.entry.get().strip()
        if not text:
            return
        self.entry.delete(0, "end")
        self.on_visitor(text)
        self.state_var.set("Thinking…")

        def say(sentence: str) -> None:
            self.ctx.speaker.say(sentence)
            self.app.ui(lambda s=sentence: self.on_partial(s))

        def done(reply):
            self.state_var.set("Listening…" if self.ctx.session.running else "Idle")
            self.on_reply(reply)

        self.app.run_bg(lambda: self.ctx.brain.respond(text, on_sentence=say), done,
                        lambda e: (self.state_var.set("Idle"), self.on_error(e)))

    def new_visitor(self) -> None:
        self.ctx.speaker.stop()
        self.ctx.brain.reset()
        self._write("\n— new visitor —\n", "meta")

    # ------------------------------------------------------------------- voice
    def toggle_voice(self) -> None:
        session = self.ctx.session
        if session.running:
            session.stop()
            self.voice_btn.state(["disabled"])       # re-enabled by the "stopped" event
            return
        if not self.ctx.listener.available:
            messagebox.showerror("Voice unavailable", self.ctx.listener.error or "Microphone not available.",
                                 parent=self)
            return
        self.app.lock()                              # kiosk mode: trainer tabs close
        self.voice_btn.configure(text="Stop voice reception")
        session.start(lambda kind, payload: self.app.ui(lambda: self._session_event(kind, payload)))

    def _session_event(self, kind: str, payload) -> None:
        if kind == "status":
            self.state_var.set(str(payload))
        elif kind == "visitor":
            self.on_visitor(payload)
        elif kind == "partial":
            self.on_partial(payload)
        elif kind == "reply":
            self.on_reply(payload)
        elif kind == "error":
            self.on_error(payload)
        elif kind == "stopped":
            self.voice_btn.configure(text="Start voice reception")
            self.voice_btn.state(["!disabled"])
            self.state_var.set("Idle")
