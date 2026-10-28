"""Speech I/O.

Speaker  – one persistent worker thread with a queue, so replies can be spoken
           sentence-by-sentence while the LLM is still generating, and can be
           interrupted instantly (Stop button).  Windows SAPI via pywin32; falls
           back to `say` (macOS) / `espeak` (Linux) / silent.
Listener – sounddevice capture with adaptive noise-floor calibration, then Google
           STT via SpeechRecognition.  Waits for the Speaker to finish first so
           Aria never transcribes her own voice.
"""
from __future__ import annotations

import queue
import shutil
import subprocess
import sys
import threading
import time


class SpeechError(RuntimeError):
    pass


# =========================================================================== TTS
class Speaker:
    def __init__(self, settings):
        self.s = settings
        self._q: queue.Queue = queue.Queue()
        self._gen = 0                       # bumped by stop(); items from older generations are dropped
        self._idle = threading.Event()
        self._idle.set()
        self._pending = 0
        self._plock = threading.Lock()
        self._backend = self._detect()
        self._thread = threading.Thread(target=self._worker, name="aria-tts", daemon=True)
        self._thread.start()

    @staticmethod
    def _detect() -> str:
        if sys.platform == "win32":
            try:
                import pythoncom  # noqa: F401  (availability probe)
                import win32com.client  # noqa: F401
                return "sapi"
            except ImportError:
                return "none"
        if sys.platform == "darwin" and shutil.which("say"):
            return "say"
        if shutil.which("espeak-ng") or shutil.which("espeak"):
            return "espeak"
        return "none"

    @property
    def backend(self) -> str:
        return self._backend

    # -------------------------------------------------------------------- API
    def say(self, text: str) -> None:
        text = (text or "").strip()
        if not text or not self.s.speak_replies or self._backend == "none":
            return
        with self._plock:
            self._pending += 1
            self._idle.clear()
            gen = self._gen
        self._q.put((gen, text))

    def stop(self) -> None:
        """Silence immediately and drop everything queued (the next say() works normally)."""
        with self._plock:
            self._gen += 1
        try:
            while True:
                self._q.get_nowait()
                self._done_one()
        except queue.Empty:
            pass

    def wait_idle(self, timeout: float | None = None) -> bool:
        return self._idle.wait(timeout)

    @property
    def speaking(self) -> bool:
        return not self._idle.is_set()

    # ---------------------------------------------------------------- internals
    def _done_one(self) -> None:
        with self._plock:
            self._pending = max(0, self._pending - 1)
            if self._pending == 0:
                self._idle.set()

    def _worker(self) -> None:
        com = None
        if self._backend == "sapi":
            import pythoncom
            pythoncom.CoInitialize()
            com = pythoncom
        try:
            while True:
                gen, text = self._q.get()
                if gen != self._gen:            # stopped after this was queued
                    self._done_one()
                    continue
                try:
                    self._speak(text, gen)
                except Exception as exc:
                    print(f"[tts] {exc}")
                finally:
                    self._done_one()
        finally:
            if com:
                com.CoUninitialize()

    def _speak(self, text: str, gen: int) -> None:
        if self._backend == "sapi":
            self._speak_sapi(text, gen)
        elif self._backend in ("say", "espeak"):
            self._speak_cmd(text, gen)

    def _speak_sapi(self, text: str, gen: int) -> None:
        import win32com.client
        sp = win32com.client.Dispatch("SAPI.SpVoice")   # fresh voice per utterance: avoids the "silent after first use" bug
        hints = [h.strip().lower() for h in self.s.voice_hint.split(",") if h.strip()]
        voices = sp.GetVoices()
        for i in range(voices.Count):
            desc = voices.Item(i).GetDescription().lower()
            if any(h in desc for h in hints):
                sp.Voice = voices.Item(i)
                break
        sp.Rate = int(self.s.voice_rate)
        sp.Volume = 100
        sp.Speak(text, 1)                               # async
        while not sp.WaitUntilDone(80):
            if gen != self._gen:
                sp.Speak("", 3)                         # async + purge
                break

    def _speak_cmd(self, text: str, gen: int) -> None:
        exe = {"say": ["say"], "espeak": [shutil.which("espeak-ng") or "espeak"]}[self._backend]
        proc = subprocess.Popen(exe + [text], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        while proc.poll() is None:
            if gen != self._gen:
                proc.terminate()
                break
            time.sleep(0.05)


# =========================================================================== STT
class Listener:
    SAMPLE_RATE = 16000
    CHUNK = 1024

    def __init__(self, settings, speaker: Speaker | None = None):
        self.s = settings
        self.speaker = speaker
        self._sd = self._np = self._sr = None
        self._error: str | None = None
        self._load()

    def _load(self) -> None:
        try:
            import numpy as np
            import sounddevice as sd
            import speech_recognition as sr
            self._np, self._sd, self._sr = np, sd, sr
        except (ImportError, OSError) as e:      # OSError: PortAudio library not found
            self._error = f"Voice input unavailable ({e}). Install: pip install sounddevice SpeechRecognition numpy"

    @property
    def available(self) -> bool:
        return self._error is None

    @property
    def error(self) -> str | None:
        return self._error

    def listen(self, cancel: threading.Event | None = None, timeout: float | None = None,
               phrase_limit: float | None = None) -> str:
        """Block until one utterance has been captured and transcribed.  Returns '' on silence/cancel."""
        if not self.available:
            raise SpeechError(self._error)
        if self.speaker is not None:
            self.speaker.wait_idle(30)
            time.sleep(0.25)                     # let the room echo die down
        audio = self._record(cancel, timeout or self.s.mic_timeout, phrase_limit or self.s.phrase_limit)
        if audio is None:
            return ""
        sr = self._sr
        rec = sr.Recognizer()
        data = sr.AudioData(audio.tobytes(), self.SAMPLE_RATE, 2)
        try:
            return rec.recognize_google(data, language=self.s.stt_language).strip()
        except sr.UnknownValueError:
            return ""
        except sr.RequestError as e:
            raise SpeechError(f"Speech service unreachable: {e}") from None

    def _record(self, cancel, timeout: float, phrase_limit: float):
        np, sd = self._np, self._sd
        frames: list = []
        state = {"noise": [], "speech": False, "last_voice": 0.0, "start": time.monotonic(),
                 "speech_start": 0.0, "thr": float(self.s.min_rms)}
        calibrate_for = 0.45

        def cb(indata, n, t, status):
            now = time.monotonic()
            frames.append(indata.copy())
            rms = float(np.sqrt(np.mean(indata.astype(np.float32) ** 2)))
            if now - state["start"] < calibrate_for:
                state["noise"].append(rms)
                return
            if state["noise"]:
                floor = float(np.mean(state["noise"]))
                state["thr"] = max(float(self.s.min_rms), floor * 2.8)
                state["noise"] = []
            if rms > state["thr"]:
                if not state["speech"]:
                    state["speech"] = True
                    state["speech_start"] = now
                state["last_voice"] = now

        try:
            stream = sd.InputStream(samplerate=self.SAMPLE_RATE, channels=1, dtype="int16",
                                    blocksize=self.CHUNK, callback=cb)
        except Exception as e:
            raise SpeechError(f"Microphone unavailable: {e}") from None
        with stream:
            while True:
                time.sleep(0.02)
                now = time.monotonic()
                if cancel is not None and cancel.is_set():
                    return None
                if not state["speech"] and now - state["start"] > timeout:
                    return None
                if state["speech"]:
                    if now - state["last_voice"] > self.s.end_silence:
                        break
                    if now - state["speech_start"] > phrase_limit:
                        break
        if not state["speech"] or not frames:
            return None
        return np.concatenate(frames, axis=0)
