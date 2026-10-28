"""Persistent settings (JSON) shared by every module.

A single ``Settings`` instance is passed around by reference, so changes made in
the UI take effect immediately everywhere without restarting.
"""
from __future__ import annotations

import hashlib
import json
from dataclasses import asdict, dataclass, fields
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
SETTINGS_PATH = DATA_DIR / "settings.json"


def hash_pin(pin: str) -> str:
    return hashlib.sha256(("aria-pin:" + pin).encode("utf-8")).hexdigest()


@dataclass
class Settings:
    # identity
    office_name: str = "Karan's Office"
    assistant_name: str = "Aria"

    # local LLM (Ollama)
    ollama_url: str = "http://localhost:11434"
    ollama_model: str = ""          # blank = auto-detect a model whose name contains "mistral"
    embed_model: str = ""           # optional, e.g. "nomic-embed-text" -> semantic matching
    temperature: float = 0.3
    num_predict: int = 160          # max tokens per spoken answer
    llm_timeout: int = 90           # seconds (first call may load the model)

    # matching
    match_threshold: float = 0.70   # >= : answer straight from trained data (no LLM)
    suggest_threshold: float = 0.40  # >= : give the LLM the closest trained answers as context

    # speech
    speak_replies: bool = True
    voice_hint: str = "zira,hazel,female"   # comma list; first matching Windows voice wins
    voice_rate: int = 1
    stt_language: str = "en-IN"
    mic_timeout: int = 10
    phrase_limit: int = 20
    end_silence: float = 1.1
    min_rms: int = 350              # floor for the adaptive silence threshold

    # security / storage
    pin_hash: str = hash_pin("1234")
    db_path: str = str(DATA_DIR / "reception.db")

    # ------------------------------------------------------------------
    @classmethod
    def load(cls, path: Path = SETTINGS_PATH) -> "Settings":
        s = cls()
        try:
            raw = json.loads(Path(path).read_text(encoding="utf-8"))
            known = {f.name for f in fields(cls)}
            for k, v in raw.items():
                if k in known:
                    setattr(s, k, v)
        except FileNotFoundError:
            pass
        except (OSError, ValueError) as exc:
            print(f"[config] could not read settings ({exc}); using defaults")
        s._path = Path(path)
        return s

    def save(self, path: Path | None = None) -> None:
        target = Path(path or getattr(self, "_path", SETTINGS_PATH))
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp = target.with_suffix(".tmp")
        tmp.write_text(json.dumps(asdict(self), indent=2), encoding="utf-8")
        tmp.replace(target)

    def check_pin(self, pin: str) -> bool:
        return hash_pin(pin) == self.pin_hash

    def set_pin(self, pin: str) -> None:
        self.pin_hash = hash_pin(pin)
