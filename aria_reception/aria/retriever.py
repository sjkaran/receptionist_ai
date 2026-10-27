"""Hybrid retrieval over trained Q&A.

Score = blend of
  * TF-IDF cosine on stemmed, synonym-folded terms
  * IDF-weighted coverage of the *trained* phrase by the visitor's words
    (so "hey, could you tell me what time you guys close today" still hits
    "What are your opening hours?")
  * character-trigram Dice (robust to speech-to-text typos)
  * optional dense embeddings from an Ollama embedding model (semantic paraphrase)

Every Q&A is indexed under its main question *and* all alternate phrasings, and
the best phrasing wins. The index rebuilds itself whenever the database changes.
"""
from __future__ import annotations

import math
import struct
import threading
import time
from dataclasses import dataclass
from typing import Callable

try:
    import numpy as np
except ImportError:  # semantic matching is simply disabled without numpy
    np = None

from .textutils import dice, term_list, trigrams

Embedder = Callable[[list[str]], list[list[float]]]


@dataclass
class Match:
    qa_id: int
    question: str
    answer: str
    score: float
    phrase: str
    lexical: float = 0.0
    semantic: float | None = None


class _Entry:
    __slots__ = ("qa_id", "phrase", "question", "answer", "weights", "norm", "tri", "emb")


class Retriever:
    def __init__(self, db, settings, embedder: Embedder | None = None):
        self.db = db
        self.settings = settings
        self.embedder = embedder
        self._entries: list[_Entry] = []
        self._idf: dict[str, float] = {}
        self._default_idf = 1.0
        self._built_version = -1
        self._built_embed_model = None
        self._lock = threading.RLock()
        self._embed_fail_until = 0.0
        self.semantic_active = False

    # ------------------------------------------------------------------ build
    def _ensure_index(self) -> None:
        model = self.settings.embed_model
        if self._built_version == self.db.version and self._built_embed_model == model:
            return
        with self._lock:
            if self._built_version == self.db.version and self._built_embed_model == model:
                return
            rows = self.db.all_phrases()
            docs = [set(term_list(r["phrase"])) for r in rows]
            n = max(len(rows), 1)
            df: dict[str, int] = {}
            for d in docs:
                for t in d:
                    df[t] = df.get(t, 0) + 1
            self._idf = {t: math.log((n + 1) / (c + 1)) + 1.0 for t, c in df.items()}
            self._default_idf = math.log(n + 1) + 1.0

            entries = []
            for r in rows:
                e = _Entry()
                e.qa_id, e.phrase, e.question, e.answer = r["qa_id"], r["phrase"], r["question"], r["answer"]
                tl = term_list(r["phrase"])
                counts: dict[str, int] = {}
                for t in tl:
                    counts[t] = counts.get(t, 0) + 1
                e.weights = {t: (1 + math.log(c)) * self._idf.get(t, self._default_idf) for t, c in counts.items()}
                e.norm = math.sqrt(sum(w * w for w in e.weights.values())) or 1.0
                e.tri = trigrams(r["phrase"])
                e.emb = None
                entries.append(e)
            self._entries = entries
            self._attach_embeddings(entries)
            self._built_version = self.db.version
            self._built_embed_model = model

    def _attach_embeddings(self, entries: list[_Entry]) -> None:
        model = self.settings.embed_model
        self.semantic_active = False
        if not (model and self.embedder and np is not None and entries) or time.time() < self._embed_fail_until:
            return
        try:
            missing, cached = [], {}
            for e in entries:
                blob = self.db.get_embedding(e.phrase, model)
                if blob is None:
                    missing.append(e.phrase)
                else:
                    cached[e.phrase] = blob
            if missing:
                uniq = list(dict.fromkeys(missing))
                vecs = self.embedder(uniq)
                for text, vec in zip(uniq, vecs):
                    blob = struct.pack(f"{len(vec)}f", *vec)
                    self.db.put_embedding(text, model, blob)
                    cached[text] = blob
            for e in entries:
                v = np.frombuffer(cached[e.phrase], dtype=np.float32)
                n = np.linalg.norm(v)
                e.emb = v / n if n else v
            self.semantic_active = True
        except Exception as exc:  # Ollama down / model missing → lexical only, retry in a minute
            print(f"[retriever] semantic matching off: {exc}")
            self._embed_fail_until = time.time() + 60
            for e in entries:
                e.emb = None

    # ----------------------------------------------------------------- search
    def search(self, query: str, k: int = 3) -> list[Match]:
        self._ensure_index()
        entries = self._entries
        if not entries or not query.strip():
            return []

        q_terms = term_list(query)
        q_counts: dict[str, int] = {}
        for t in q_terms:
            q_counts[t] = q_counts.get(t, 0) + 1
        q_w = {t: (1 + math.log(c)) * self._idf.get(t, self._default_idf) for t, c in q_counts.items()}
        q_norm = math.sqrt(sum(w * w for w in q_w.values())) or 1.0
        q_tri = trigrams(query)

        q_emb = None
        if self.semantic_active:
            try:
                q_emb = self._embed_query(query)
            except Exception as exc:
                print(f"[retriever] query embedding failed: {exc}")

        best: dict[int, Match] = {}
        for e in entries:
            shared = q_w.keys() & e.weights.keys()
            dot = sum(q_w[t] * e.weights[t] for t in shared)
            cos = dot / (q_norm * e.norm)
            # coverage of the trained phrase's own terms (idf-weighted)
            total = sum(self._idf.get(t, self._default_idf) for t in e.weights)
            have = sum(self._idf.get(t, self._default_idf) for t in shared)
            coverage = have / total if total else 0.0
            tri = dice(q_tri, e.tri)
            lex = 0.45 * cos + 0.35 * coverage + 0.20 * tri
            if not shared:
                lex = min(lex, 0.30 + 0.5 * tri * 0.2)   # no shared concept → can't be a confident hit

            sem = None
            score = lex
            if q_emb is not None and e.emb is not None:
                cs = float(np.dot(q_emb, e.emb))
                sem = max(0.0, min(1.0, (cs - 0.55) / 0.35))
                score = max(lex, 0.35 * lex + 0.65 * sem)
            m = best.get(e.qa_id)
            if m is None or score > m.score:
                best[e.qa_id] = Match(e.qa_id, e.question, e.answer, round(score, 4), e.phrase, lex, sem)
        out = sorted(best.values(), key=lambda m: m.score, reverse=True)
        return out[:k]

    def _embed_query(self, query: str):
        model = self.settings.embed_model
        blob = self.db.get_embedding(query, model)
        if blob is None:
            vec = self.embedder([query])[0]
            blob = struct.pack(f"{len(vec)}f", *vec)
            self.db.put_embedding(query, model, blob)
        v = np.frombuffer(blob, dtype=np.float32)
        n = np.linalg.norm(v)
        return v / n if n else v

    def best(self, query: str) -> Match | None:
        r = self.search(query, 1)
        return r[0] if r else None
