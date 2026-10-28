"""Central configuration. Settings persist in data/settings.json and are editable from the UI."""
import json
import threading
from pathlib import Path

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
DB_PATH = DATA_DIR / "reception.db"
SETTINGS_PATH = DATA_DIR / "settings.json"

DEFAULTS = {
    # Identity
    "office_name": "Karan's Office",
    "assistant_name": "Aria",
    "office_info": "",            # facts the AI may use: hours, address, departments...
    # Local AI (Ollama)
    "ollama_url": "http://localhost:11434",
    "ollama_model": "mistral",
    "ollama_timeout": 90,
    # Matching
    "similarity_threshold": 0.55,
    # Voice
    "voice_hint": "zira,hazel,female",   # comma list, first installed match wins
    "speech_rate": 1,                    # SAPI range -10..10
    "silence_rms": 500,                  # mic level treated as speech
    "listen_timeout": 10,
    "phrase_limit": 20,
    # Admin
    "trainer_pin": "1234",
}


class Settings:
    def __init__(self):
        self._lock = threading.Lock()
        self.data = dict(DEFAULTS)
        self.load()

    def load(self):
        if SETTINGS_PATH.exists():
            try:
                self.data.update(json.loads(SETTINGS_PATH.read_text(encoding="utf-8")))
            except (json.JSONDecodeError, OSError):
                pass

    def save(self):
        SETTINGS_PATH.write_text(json.dumps(self.data, indent=2), encoding="utf-8")

    def update(self, patch: dict):
        with self._lock:
            for key, default in DEFAULTS.items():
                if key in patch and patch[key] is not None:
                    try:
                        self.data[key] = type(default)(patch[key])
                    except (TypeError, ValueError):
                        pass
            self.save()

    def __getitem__(self, key):
        return self.data[key]

    def public(self) -> dict:
        d = dict(self.data)
        d.pop("trainer_pin", None)
        return d


settings = Settings()
