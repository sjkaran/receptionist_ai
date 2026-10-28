"""Text-to-speech via Windows SAPI. A fresh COM thread per utterance avoids the 'silent after first use' bug."""
import threading

try:
    import pythoncom
    import win32com.client
    HAVE_SAPI = True
except ImportError:          # non-Windows dev machines: fall back to console only
    HAVE_SAPI = False


class VoiceEngine:
    def __init__(self, settings):
        self.settings = settings
        self._lock = threading.Lock()

    def speak(self, text: str):
        """Blocking speak."""
        if not text:
            return
        if not HAVE_SAPI:
            print(f"[TTS unavailable] {text}")
            return

        hints = [h.strip().lower() for h in str(self.settings["voice_hint"]).split(",") if h.strip()]
        rate = int(self.settings["speech_rate"])

        def _run():
            pythoncom.CoInitialize()
            try:
                speaker = win32com.client.Dispatch("SAPI.SpVoice")
                voices = speaker.GetVoices()
                found = False
                for hint in hints:
                    for i in range(voices.Count):
                        if hint in voices.Item(i).GetDescription().lower():
                            speaker.Voice = voices.Item(i)
                            found = True
                            break
                    if found:
                        break
                speaker.Rate = max(-10, min(10, rate))
                speaker.Volume = 100
                speaker.Speak(text)
            finally:
                pythoncom.CoUninitialize()

        with self._lock:
            t = threading.Thread(target=_run, daemon=True)
            t.start()
            t.join()
