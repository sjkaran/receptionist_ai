"""Wires every component together once; the UI and the CLI both start from here."""
from __future__ import annotations

from dataclasses import dataclass

from .brain import Receptionist
from .config import Settings
from .database import Database
from .llm import OllamaClient
from .retriever import Retriever
from .session import VoiceSession
from .speech import Listener, Speaker


@dataclass
class Context:
    settings: Settings
    db: Database
    llm: OllamaClient
    retriever: Retriever
    brain: Receptionist
    speaker: Speaker
    listener: Listener
    session: VoiceSession

    @classmethod
    def create(cls, settings: Settings | None = None, seed: bool = True) -> "Context":
        settings = settings or Settings.load()
        db = Database(settings.db_path)
        if seed:
            db.seed_sample_data(settings.office_name, settings.assistant_name)
        llm = OllamaClient(settings)
        retriever = Retriever(db, settings, embedder=llm.embed)
        brain = Receptionist(db, settings, retriever, llm)
        speaker = Speaker(settings)
        listener = Listener(settings, speaker)
        session = VoiceSession(brain, speaker, listener, settings)
        return cls(settings, db, llm, retriever, brain, speaker, listener, session)
