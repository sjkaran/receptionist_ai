"""Settings tab: identity, Ollama connection + model pickers, thresholds, voice, PIN."""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk



class SettingsTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=14)
        self.app, self.ctx = app, app.ctx
        s = self.ctx.settings
        self.v = {
            "office_name": tk.StringVar(value=s.office_name),
            "assistant_name": tk.StringVar(value=s.assistant_name),
            "ollama_url": tk.StringVar(value=s.ollama_url),
            "ollama_model": tk.StringVar(value=s.ollama_model),
            "embed_model": tk.StringVar(value=s.embed_model),
            "temperature": tk.StringVar(value=str(s.temperature)),
            "match_threshold": tk.StringVar(value=str(s.match_threshold)),
            "suggest_threshold": tk.StringVar(value=str(s.suggest_threshold)),
            "speak_replies": tk.BooleanVar(value=s.speak_replies),
            "voice_hint": tk.StringVar(value=s.voice_hint),
            "stt_language": tk.StringVar(value=s.stt_language),
        }
        self.new_pin, self.new_pin2 = tk.StringVar(), tk.StringVar()

        self.columnconfigure(0, weight=1, uniform="c")
        self.columnconfigure(1, weight=1, uniform="c")

        # ---------------- left column
        left = ttk.Frame(self)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        left.columnconfigure(1, weight=1)
        r = 0
        ttk.Label(left, text="Identity", style="H2.TLabel").grid(row=r, column=0, columnspan=2, sticky="w")
        r += 1
        self._row(left, r, "Office name", ttk.Entry(left, textvariable=self.v["office_name"])); r += 1
        self._row(left, r, "Assistant name", ttk.Entry(left, textvariable=self.v["assistant_name"])); r += 1

        ttk.Label(left, text="Local AI (Ollama)", style="H2.TLabel").grid(row=r, column=0, columnspan=2, sticky="w", pady=(16, 0))
        r += 1
        self._row(left, r, "Server URL", ttk.Entry(left, textvariable=self.v["ollama_url"])); r += 1
        self.model_cb = ttk.Combobox(left, textvariable=self.v["ollama_model"])
        self._row(left, r, "Chat model", self.model_cb); r += 1
        self.embed_cb = ttk.Combobox(left, textvariable=self.v["embed_model"])
        self._row(left, r, "Embedding model", self.embed_cb); r += 1
        ttk.Label(left, text="Leave chat model blank to auto-pick Mistral. Embedding model is optional "
                             "(e.g. nomic-embed-text) and adds paraphrase-level matching.",
                  style="Muted.TLabel", wraplength=420, justify="left").grid(row=r, column=0, columnspan=2, sticky="w", pady=(2, 6))
        r += 1
        row = ttk.Frame(left)
        row.grid(row=r, column=0, columnspan=2, sticky="w")
        ttk.Button(row, text="Test connection", command=self.test).pack(side="left")
        self.test_var = tk.StringVar()
        ttk.Label(row, textvariable=self.test_var, style="Muted.TLabel", wraplength=300).pack(side="left", padx=10)

        # ---------------- right column
        right = ttk.Frame(self)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(1, weight=1)
        r = 0
        ttk.Label(right, text="Behaviour", style="H2.TLabel").grid(row=r, column=0, columnspan=2, sticky="w"); r += 1
        self._row(right, r, "Answer directly at ≥", ttk.Spinbox(right, from_=0.4, to=0.98, increment=0.01, width=8,
                                                                 textvariable=self.v["match_threshold"], format="%.2f")); r += 1
        self._row(right, r, "Use as AI reference at ≥", ttk.Spinbox(right, from_=0.1, to=0.8, increment=0.01, width=8,
                                                                     textvariable=self.v["suggest_threshold"], format="%.2f")); r += 1
        self._row(right, r, "AI creativity", ttk.Spinbox(right, from_=0.0, to=1.0, increment=0.05, width=8,
                                                           textvariable=self.v["temperature"], format="%.2f")); r += 1

        ttk.Label(right, text="Voice", style="H2.TLabel").grid(row=r, column=0, columnspan=2, sticky="w", pady=(16, 0)); r += 1
        self._row(right, r, "Speak replies", ttk.Checkbutton(right, variable=self.v["speak_replies"])); r += 1
        self._row(right, r, "Voice (name hints)", ttk.Entry(right, textvariable=self.v["voice_hint"])); r += 1
        self._row(right, r, "Speech language", ttk.Entry(right, textvariable=self.v["stt_language"])); r += 1

        ttk.Label(right, text="Trainer PIN", style="H2.TLabel").grid(row=r, column=0, columnspan=2, sticky="w", pady=(16, 0)); r += 1
        self._row(right, r, "New PIN", ttk.Entry(right, textvariable=self.new_pin, show="•")); r += 1
        self._row(right, r, "Repeat PIN", ttk.Entry(right, textvariable=self.new_pin2, show="•")); r += 1

        bottom = ttk.Frame(self)
        bottom.grid(row=1, column=0, columnspan=2, sticky="w", pady=(18, 0))
        ttk.Button(bottom, text="Save settings", style="Accent.TButton", command=self.save).pack(side="left")
        self.save_var = tk.StringVar()
        ttk.Label(bottom, textvariable=self.save_var, style="Muted.TLabel").pack(side="left", padx=12)
        self.after(300, self.load_models)

    @staticmethod
    def _row(parent, r, label, widget) -> None:
        ttk.Label(parent, text=label, style="Muted.TLabel").grid(row=r, column=0, sticky="w", pady=4, padx=(0, 12))
        widget.grid(row=r, column=1, sticky="ew" if not isinstance(widget, (ttk.Spinbox, ttk.Checkbutton)) else "w", pady=4)

    def on_show(self) -> None:
        self.load_models()

    # ----------------------------------------------------------------- models
    def load_models(self) -> None:
        def ok(models):
            self.model_cb.configure(values=[""] + models)
            self.embed_cb.configure(values=[""] + models)

        self.app.run_bg(self.ctx.llm.models, ok, lambda e: None)

    def test(self) -> None:
        self._apply_connection()
        self.test_var.set("Testing…")

        def work():
            llm = self.ctx.llm
            llm.reset()
            model = llm.resolve_model()
            return model, llm.complete("Reply with one short word.", "Say hello.", num_predict=8)

        self.app.run_bg(work,
                        lambda res: (self.test_var.set(f"OK — {res[0]} replied “{res[1][:30]}”"), self.load_models(),
                                     self.app.refresh_status()),
                        lambda e: self.test_var.set(str(e)))

    def _apply_connection(self) -> None:
        s = self.ctx.settings
        s.ollama_url = self.v["ollama_url"].get().strip() or "http://localhost:11434"
        s.ollama_model = self.v["ollama_model"].get().strip()
        s.embed_model = self.v["embed_model"].get().strip()

    # ------------------------------------------------------------------- save
    def save(self) -> None:
        s = self.ctx.settings
        try:
            temp = float(self.v["temperature"].get())
            match = float(self.v["match_threshold"].get())
            sugg = float(self.v["suggest_threshold"].get())
        except ValueError:
            messagebox.showwarning("Invalid number", "Thresholds and creativity must be numbers.", parent=self)
            return
        if not 0 < sugg < match <= 1:
            messagebox.showwarning("Invalid thresholds",
                                   "“AI reference” must be lower than “Answer directly”, and both between 0 and 1.",
                                   parent=self)
            return
        pin, pin2 = self.new_pin.get().strip(), self.new_pin2.get().strip()
        if pin or pin2:
            if pin != pin2 or len(pin) < 4:
                messagebox.showwarning("PIN", "PINs must match and be at least 4 characters.", parent=self)
                return
            s.set_pin(pin)
        s.office_name = self.v["office_name"].get().strip() or s.office_name
        s.assistant_name = self.v["assistant_name"].get().strip() or s.assistant_name
        s.temperature, s.match_threshold, s.suggest_threshold = temp, match, sugg
        s.speak_replies = bool(self.v["speak_replies"].get())
        s.voice_hint = self.v["voice_hint"].get().strip()
        s.stt_language = self.v["stt_language"].get().strip() or "en-IN"
        self._apply_connection()
        s.save()
        self.ctx.llm.reset()
        self.new_pin.set("")
        self.new_pin2.set("")
        self.app.refresh_titles()
        self.app.refresh_status()
        self.save_var.set("Saved.")
        self.after(2500, lambda: self.save_var.set(""))
