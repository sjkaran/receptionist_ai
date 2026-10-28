"""Hands-free reception loop: greet -> listen -> think -> speak, until stopped.

Runs on its own thread and reports progress through a single `emit(kind, payload)`
callback (the UI forwards these onto the Tk thread).

Events:  status(str) · visitor(str) · partial(str) · reply(Reply) · error(str) · stopped(None)
"""
from __future__ import annotations

import re
import threading
import time
from typing import Callable

from .speech import SpeechError

EXIT_PHRASES = re.compile(r"\b(exit reception|quit assistant|stop assistant|shut down assistant)\b", re.I)
IDLE_RESET_SECONDS = 90          # forget the previous visitor after this much quiet


class VoiceSession:
    def __init__(self, brain, speaker, listener, settings):
        self.brain, self.speaker, self.listener, self.s = brain, speaker, listener, settings
        self._thread: threading.Thread | None = None
        self._stop = threading.Event()
        self._emit: Callable[[str, object], None] = lambda k, p: None

    @property
    def running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def start(self, emit: Callable[[str, object], None]) -> None:
        if self.running:
            return
        self._emit = emit
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, name="aria-session", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()
        self.brain.cancel.set()
        self.speaker.stop()

    # ------------------------------------------------------------------ loop
    def _say(self, sentence: str) -> None:
        self.speaker.say(sentence)
        self._emit("partial", sentence)

    def _run(self) -> None:
        emit = self._emit
        try:
            self.brain.reset()
            greeting = self.brain.greeting()
            emit("status", "Greeting")
            self._say(greeting)
            emit("reply", None)
            silent_rounds, last_activity = 0, time.monotonic()

            while not self._stop.is_set():
                emit("status", "Listening…")
                try:
                    text = self.listener.listen(cancel=self._stop)
                except SpeechError as exc:
                    emit("error", str(exc))
                    if "unavailable" in str(exc).lower():
                        break
                    self._stop.wait(3)
                    continue
                if self._stop.is_set():
                    break

                if not text:
                    silent_rounds += 1
                    if silent_rounds >= 3:
                        self._say("I am here whenever you need help. Please feel free to ask me anything.")
                        emit("reply", None)
                        silent_rounds = 0
                    if self.brain.history and time.monotonic() - last_activity > IDLE_RESET_SECONDS:
                        self.brain.reset()
                    continue

                silent_rounds, last_activity = 0, time.monotonic()
                emit("visitor", text)
                if EXIT_PHRASES.search(text):
                    self._say("Thank you for visiting. Have a wonderful day!")
                    emit("reply", None)
                    self.speaker.wait_idle(15)
                    break

                emit("status", "Thinking…")
                reply = self.brain.respond(text, on_sentence=self._say)
                emit("reply", reply)
                if reply.intent == "bye":
                    self.speaker.wait_idle(15)
                    self.brain.reset()          # visitor said goodbye – next person starts fresh
        except Exception as exc:               # never let the kiosk thread die silently
            emit("error", f"Voice session crashed: {exc}")
        finally:
            self.speaker.stop()
            emit("status", "Idle")
            emit("stopped", None)
