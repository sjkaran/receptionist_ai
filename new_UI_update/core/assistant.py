"""The reception engine: state machine, event feed, listen -> match -> (Ollama) -> speak loop."""
import threading
import time

from config import settings
from core import database as db
from core.llm import LLMError, OllamaClient
from core.matcher import find_best_answer
from core.stt import STTError, listen
from core.tts import VoiceEngine

EXIT_PHRASES = ("exit reception", "quit assistant", "stop assistant")
FALLBACK_REPLY = "I'm sorry, I'm having trouble with that right now. Please speak to our team directly."


class ReceptionAssistant:
    def __init__(self):
        self.voice = VoiceEngine(settings)
        self.llm = OllamaClient(settings)
        self.state = "idle"            # idle | listening | thinking | speaking
        self.running = False
        self.history = []
        self.events = []
        self._eid = 0
        self._lock = threading.Lock()
        self._mic = threading.Lock()
        self._stop = threading.Event()
        self._thread = None

    # ---- event feed (polled by the UI) ----
    def emit(self, kind: str, text: str, **extra):
        with self._lock:
            self._eid += 1
            self.events.append({"id": self._eid, "kind": kind, "text": text,
                                "time": time.strftime("%H:%M:%S"), **extra})
            self.events = self.events[-200:]

    def events_after(self, after: int):
        with self._lock:
            return [e for e in self.events if e["id"] > after], self._eid

    # ---- core answer logic (shared by voice loop and typed test box) ----
    def respond(self, text: str) -> dict:
        answer, qid, score = find_best_answer(text, float(settings["similarity_threshold"]))
        if answer:
            db.increment_hit(qid)
            source = "trained"
        else:
            try:
                answer, source = self.llm.ask(text, self.history), "ollama"
            except LLMError as e:
                self.emit("system", str(e))
                answer, source = FALLBACK_REPLY, "fallback"

        self.history.append({"user": text, "assistant": answer})
        self.history = self.history[-20:]
        db.log_conversation(text, answer, source, score)
        return {"answer": answer, "source": source, "score": round(score, 2)}

    def say(self, text: str):
        self.state = "speaking"
        try:
            self.voice.speak(text)
        finally:
            self.state = "listening" if self.running else "idle"

    # ---- reception loop ----
    def start(self):
        if self.running:
            return
        self._stop.clear()
        self.running = True
        self.history = []
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def stop(self):
        self._stop.set()

    def _greeting(self) -> str:
        greetings = db.list_greetings()
        if greetings:
            return greetings[0]["text"]
        return (f"Welcome to {settings['office_name']}! I am {settings['assistant_name']}, "
                "your virtual receptionist. How may I help you today?")

    def _loop(self):
        self.emit("system", "Reception started")
        try:
            greeting = self._greeting()
            self.emit("assistant", greeting, source="greeting")
            self.say(greeting)

            silent = 0
            while not self._stop.is_set():
                self.state = "listening"
                try:
                    heard = listen(settings, stop_event=self._stop)
                except STTError as e:
                    self.emit("system", str(e))
                    time.sleep(3)
                    continue
                if self._stop.is_set():
                    break

                if not heard:
                    silent += 1
                    if silent >= 3:
                        nudge = "I am here to help you. Please feel free to ask me anything."
                        self.emit("assistant", nudge, source="nudge")
                        self.say(nudge)
                        silent = 0
                    continue

                silent = 0
                self.emit("visitor", heard)

                if any(p in heard.lower() for p in EXIT_PHRASES):
                    farewell = "Thank you for visiting. Have a wonderful day!"
                    self.emit("assistant", farewell, source="farewell")
                    self.say(farewell)
                    break

                self.state = "thinking"
                result = self.respond(heard)
                self.emit("assistant", result["answer"], source=result["source"], score=result["score"])
                self.say(result["answer"])
        except Exception as e:                       # keep the server alive no matter what
            self.emit("system", f"Reception error: {e}")
        finally:
            self.running = False
            self.state = "idle"
            self.emit("system", "Reception stopped")

    # ---- one-shot dictation for training forms ----
    def capture_once(self) -> str:
        if self.running:
            raise RuntimeError("Stop reception mode before using the microphone for training.")
        if not self._mic.acquire(blocking=False):
            raise RuntimeError("The microphone is already in use.")
        try:
            self.state = "listening"
            return listen(settings)
        finally:
            self.state = "idle"
            self._mic.release()


assistant = ReceptionAssistant()
