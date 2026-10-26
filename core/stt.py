"""Microphone capture (sounddevice) + Google speech recognition."""
import threading
import time
from typing import Optional

SAMPLE_RATE = 16000
CHUNK = 1024


class STTError(RuntimeError):
    pass


def record_until_silence(timeout: int, phrase_limit: int, silence_rms: float,
                         stop_event: Optional[threading.Event] = None):
    """Record until the speaker pauses. Returns int16 numpy array, or None (timeout / stopped)."""
    import numpy as np
    import sounddevice as sd

    frames = []
    state = {"started": False, "silence": 0, "total": 0}
    timeout_chunks = int(SAMPLE_RATE / CHUNK * timeout)
    phrase_chunks = int(SAMPLE_RATE / CHUNK * phrase_limit)
    silence_limit = int(SAMPLE_RATE / CHUNK * 1.2)

    def callback(indata, frame_count, time_info, status):
        frames.append(indata.copy())
        rms = np.sqrt(np.mean(indata.astype(np.float32) ** 2))
        if rms > silence_rms:
            state["started"], state["silence"] = True, 0
        elif state["started"]:
            state["silence"] += 1
        state["total"] += 1

    try:
        stream = sd.InputStream(samplerate=SAMPLE_RATE, channels=1, dtype="int16",
                                blocksize=CHUNK, callback=callback)
    except Exception as e:
        raise STTError(f"Microphone unavailable: {e}")

    with stream:
        while True:
            time.sleep(0.01)
            if stop_event is not None and stop_event.is_set():
                return None
            if not state["started"] and state["total"] > timeout_chunks:
                return None
            if state["started"] and state["silence"] >= silence_limit:
                break
            if state["total"] > phrase_chunks:
                break

    if not state["started"] or not frames:
        return None
    return np.concatenate(frames, axis=0)


def listen(settings, timeout=None, phrase_limit=None, stop_event=None) -> str:
    """Listen once and return recognised text ('' if nothing understood)."""
    import speech_recognition as sr

    audio = record_until_silence(
        timeout or int(settings["listen_timeout"]),
        phrase_limit or int(settings["phrase_limit"]),
        float(settings["silence_rms"]),
        stop_event,
    )
    if audio is None:
        return ""

    data = sr.AudioData(audio.tobytes(), SAMPLE_RATE, 2)
    try:
        return sr.Recognizer().recognize_google(data).strip()
    except sr.UnknownValueError:
        return ""
    except sr.RequestError as e:
        raise STTError(f"Speech service error (internet needed): {e}")
