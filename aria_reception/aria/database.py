"""SQLite storage: Q&A, alternate phrasings, facts, greetings, review queue,
visitor messages, conversation log and an embedding cache.

A fresh connection per operation keeps the class safe to use from the UI thread,
the voice thread and LLM worker threads at the same time (WAL mode).
"""
from __future__ import annotations

import hashlib
import os
import random
import sqlite3
import threading
from contextlib import contextmanager
from datetime import datetime

from .textutils import normalize

SCHEMA = """
CREATE TABLE IF NOT EXISTS greetings(
    id INTEGER PRIMARY KEY AUTOINCREMENT, text TEXT NOT NULL, created TEXT);
CREATE TABLE IF NOT EXISTS qa(
    id INTEGER PRIMARY KEY AUTOINCREMENT, question TEXT NOT NULL, answer TEXT NOT NULL,
    category TEXT NOT NULL DEFAULT '', hits INTEGER NOT NULL DEFAULT 0,
    enabled INTEGER NOT NULL DEFAULT 1, created TEXT, updated TEXT);
CREATE TABLE IF NOT EXISTS qa_variants(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    qa_id INTEGER NOT NULL REFERENCES qa(id) ON DELETE CASCADE, phrase TEXT NOT NULL);
CREATE INDEX IF NOT EXISTS idx_variants_qa ON qa_variants(qa_id);
CREATE TABLE IF NOT EXISTS facts(
    id INTEGER PRIMARY KEY AUTOINCREMENT, label TEXT NOT NULL, value TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS review_queue(
    id INTEGER PRIMARY KEY AUTOINCREMENT, question TEXT NOT NULL, norm TEXT NOT NULL UNIQUE,
    count INTEGER NOT NULL DEFAULT 1, suggested TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT 'pending', first_seen TEXT, last_seen TEXT);
CREATE TABLE IF NOT EXISTS messages(
    id INTEGER PRIMARY KEY AUTOINCREMENT, visitor TEXT, contact TEXT, body TEXT NOT NULL,
    created TEXT, handled INTEGER NOT NULL DEFAULT 0);
CREATE TABLE IF NOT EXISTS log(
    id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, role TEXT, text TEXT, source TEXT, score REAL);
CREATE TABLE IF NOT EXISTS embeddings(
    hash TEXT, model TEXT, vec BLOB, PRIMARY KEY(hash, model));
"""


def _now() -> str:
    return datetime.now().isoformat(sep=" ", timespec="seconds")


class Database:
    def __init__(self, path: str):
        self.path = path
        self.version = 0          # bumped whenever Q&A content changes (retriever watches this)
        self._vlock = threading.Lock()
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        with self._tx() as c:
            c.execute("PRAGMA journal_mode=WAL")
            c.executescript(SCHEMA)

    # ------------------------------------------------------------------ infra
    @contextmanager
    def _tx(self):
        conn = sqlite3.connect(self.path, timeout=15)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys=ON")
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _bump(self) -> None:
        with self._vlock:
            self.version += 1

    @staticmethod
    def _rows(rows) -> list[dict]:
        return [dict(r) for r in rows]

    # -------------------------------------------------------------------- Q&A
    def add_qa(self, question: str, answer: str, category: str = "", variants=()) -> int:
        question, answer = question.strip(), answer.strip()
        if not question or not answer:
            raise ValueError("Both a question and an answer are required.")
        with self._tx() as c:
            cur = c.execute(
                "INSERT INTO qa(question, answer, category, created, updated) VALUES(?,?,?,?,?)",
                (question, answer, category.strip(), _now(), _now()),
            )
            qa_id = cur.lastrowid
            self._write_variants(c, qa_id, question, variants)
        self._bump()
        return qa_id

    def update_qa(self, qa_id: int, question: str, answer: str, category: str = "", variants=None) -> None:
        question, answer = question.strip(), answer.strip()
        if not question or not answer:
            raise ValueError("Both a question and an answer are required.")
        with self._tx() as c:
            c.execute(
                "UPDATE qa SET question=?, answer=?, category=?, updated=? WHERE id=?",
                (question, answer, category.strip(), _now(), qa_id),
            )
            if variants is not None:
                c.execute("DELETE FROM qa_variants WHERE qa_id=?", (qa_id,))
                self._write_variants(c, qa_id, question, variants)
        self._bump()

    @staticmethod
    def _write_variants(c, qa_id: int, question: str, variants) -> None:
        seen = {normalize(question)}
        for v in variants or ():
            n = normalize(v)
            if n and n not in seen:
                seen.add(n)
                c.execute("INSERT INTO qa_variants(qa_id, phrase) VALUES(?,?)", (qa_id, v.strip()))

    def delete_qa(self, qa_id: int) -> None:
        with self._tx() as c:
            c.execute("DELETE FROM qa WHERE id=?", (qa_id,))
        self._bump()

    def set_enabled(self, qa_id: int, enabled: bool) -> None:
        with self._tx() as c:
            c.execute("UPDATE qa SET enabled=? WHERE id=?", (int(enabled), qa_id))
        self._bump()

    def get_qa(self, qa_id: int) -> dict | None:
        with self._tx() as c:
            row = c.execute("SELECT * FROM qa WHERE id=?", (qa_id,)).fetchone()
            if not row:
                return None
            d = dict(row)
            d["variants"] = [r["phrase"] for r in c.execute(
                "SELECT phrase FROM qa_variants WHERE qa_id=? ORDER BY id", (qa_id,))]
        return d

    def list_qa(self, search: str = "", enabled_only: bool = False) -> list[dict]:
        sql, args = "SELECT * FROM qa", []
        where = []
        if enabled_only:
            where.append("enabled=1")
        if search.strip():
            where.append("(question LIKE ? OR answer LIKE ? OR category LIKE ?)")
            args += [f"%{search.strip()}%"] * 3
        if where:
            sql += " WHERE " + " AND ".join(where)
        sql += " ORDER BY hits DESC, id DESC"
        with self._tx() as c:
            return self._rows(c.execute(sql, args))

    def all_phrases(self) -> list[dict]:
        """Every phrasing (main question + variants) of every enabled Q&A, for the index."""
        with self._tx() as c:
            return self._rows(c.execute(
                """SELECT q.id AS qa_id, q.question AS phrase, q.question AS question, q.answer AS answer
                   FROM qa q WHERE q.enabled=1
                   UNION ALL
                   SELECT q.id, v.phrase, q.question, q.answer
                   FROM qa_variants v JOIN qa q ON q.id=v.qa_id WHERE q.enabled=1"""))

    def hit(self, qa_id: int) -> None:
        with self._tx() as c:
            c.execute("UPDATE qa SET hits=hits+1 WHERE id=?", (qa_id,))

    # -------------------------------------------------------------- greetings
    def add_greeting(self, text: str) -> int:
        if not text.strip():
            raise ValueError("Greeting is empty.")
        with self._tx() as c:
            return c.execute("INSERT INTO greetings(text, created) VALUES(?,?)", (text.strip(), _now())).lastrowid

    def list_greetings(self) -> list[dict]:
        with self._tx() as c:
            return self._rows(c.execute("SELECT * FROM greetings ORDER BY id"))

    def delete_greeting(self, gid: int) -> None:
        with self._tx() as c:
            c.execute("DELETE FROM greetings WHERE id=?", (gid,))

    def pick_greeting(self) -> str | None:
        g = self.list_greetings()
        return random.choice(g)["text"] if g else None

    # ------------------------------------------------------------------ facts
    def add_fact(self, label: str, value: str) -> int:
        if not label.strip() or not value.strip():
            raise ValueError("Both a label and a value are required.")
        with self._tx() as c:
            return c.execute("INSERT INTO facts(label, value) VALUES(?,?)", (label.strip(), value.strip())).lastrowid

    def update_fact(self, fid: int, label: str, value: str) -> None:
        with self._tx() as c:
            c.execute("UPDATE facts SET label=?, value=? WHERE id=?", (label.strip(), value.strip(), fid))

    def delete_fact(self, fid: int) -> None:
        with self._tx() as c:
            c.execute("DELETE FROM facts WHERE id=?", (fid,))

    def list_facts(self) -> list[dict]:
        with self._tx() as c:
            return self._rows(c.execute("SELECT * FROM facts ORDER BY id"))

    # ----------------------------------------------------------- review queue
    def queue_question(self, question: str, suggested: str = "") -> None:
        norm = normalize(question)
        if not norm:
            return
        with self._tx() as c:
            c.execute(
                """INSERT INTO review_queue(question, norm, count, suggested, first_seen, last_seen)
                   VALUES(?,?,1,?,?,?)
                   ON CONFLICT(norm) DO UPDATE SET
                     count = count + 1, last_seen = excluded.last_seen,
                     suggested = CASE WHEN review_queue.suggested = '' THEN excluded.suggested
                                      ELSE review_queue.suggested END""",
                (question.strip(), norm, suggested.strip(), _now(), _now()),
            )

    def list_review(self, status: str = "pending") -> list[dict]:
        with self._tx() as c:
            return self._rows(c.execute(
                "SELECT * FROM review_queue WHERE status=? ORDER BY count DESC, last_seen DESC", (status,)))

    def pending_count(self) -> int:
        with self._tx() as c:
            return c.execute("SELECT COUNT(*) FROM review_queue WHERE status='pending'").fetchone()[0]

    def approve_review(self, rid: int, question: str, answer: str, category: str = "") -> int:
        """Turn a reviewed visitor question into trained Q&A (the original wording becomes a variant)."""
        with self._tx() as c:
            row = c.execute("SELECT question FROM review_queue WHERE id=?", (rid,)).fetchone()
            original = row["question"] if row else question
        qa_id = self.add_qa(question, answer, category, variants=[original])
        with self._tx() as c:
            c.execute("UPDATE review_queue SET status='approved' WHERE id=?", (rid,))
        return qa_id

    def dismiss_review(self, rid: int) -> None:
        with self._tx() as c:
            c.execute("UPDATE review_queue SET status='dismissed' WHERE id=?", (rid,))

    # --------------------------------------------------------------- messages
    def add_message(self, visitor: str, contact: str, body: str) -> int:
        with self._tx() as c:
            return c.execute(
                "INSERT INTO messages(visitor, contact, body, created) VALUES(?,?,?,?)",
                (visitor, contact, body, _now())).lastrowid

    def list_messages(self, include_handled: bool = False) -> list[dict]:
        sql = "SELECT * FROM messages" + ("" if include_handled else " WHERE handled=0") + " ORDER BY id DESC"
        with self._tx() as c:
            return self._rows(c.execute(sql))

    def set_message_handled(self, mid: int, handled: bool = True) -> None:
        with self._tx() as c:
            c.execute("UPDATE messages SET handled=? WHERE id=?", (int(handled), mid))

    def unhandled_count(self) -> int:
        with self._tx() as c:
            return c.execute("SELECT COUNT(*) FROM messages WHERE handled=0").fetchone()[0]

    # -------------------------------------------------------------------- log
    def log_turn(self, role: str, text: str, source: str = "", score: float | None = None) -> None:
        with self._tx() as c:
            c.execute("INSERT INTO log(ts, role, text, source, score) VALUES(?,?,?,?,?)",
                      (_now(), role, text, source, score))

    def stats(self) -> dict:
        with self._tx() as c:
            one = lambda q: c.execute(q).fetchone()[0]  # noqa: E731
            return {
                "qa": one("SELECT COUNT(*) FROM qa"),
                "facts": one("SELECT COUNT(*) FROM facts"),
                "pending_review": one("SELECT COUNT(*) FROM review_queue WHERE status='pending'"),
                "messages_open": one("SELECT COUNT(*) FROM messages WHERE handled=0"),
                "answered_trained": one("SELECT COUNT(*) FROM log WHERE role='aria' AND source='trained'"),
                "answered_ai": one("SELECT COUNT(*) FROM log WHERE role='aria' AND source='ai'"),
            }

    # -------------------------------------------------------- embedding cache
    @staticmethod
    def _ehash(text: str) -> str:
        return hashlib.sha1(normalize(text).encode()).hexdigest()

    def get_embedding(self, text: str, model: str) -> bytes | None:
        with self._tx() as c:
            r = c.execute("SELECT vec FROM embeddings WHERE hash=? AND model=?",
                          (self._ehash(text), model)).fetchone()
        return r["vec"] if r else None

    def put_embedding(self, text: str, model: str, blob: bytes) -> None:
        with self._tx() as c:
            c.execute("INSERT OR REPLACE INTO embeddings(hash, model, vec) VALUES(?,?,?)",
                      (self._ehash(text), model, blob))

    # ---------------------------------------------------------- import/export
    def export_data(self) -> dict:
        qa = []
        for row in self.list_qa():
            full = self.get_qa(row["id"])
            qa.append({k: full[k] for k in ("question", "answer", "category", "enabled", "variants")})
        return {
            "version": 2,
            "greetings": [g["text"] for g in self.list_greetings()],
            "facts": [{"label": f["label"], "value": f["value"]} for f in self.list_facts()],
            "qa": qa,
        }

    def import_data(self, data: dict) -> dict:
        existing = {normalize(q["question"]) for q in self.list_qa()}
        added = skipped = 0
        for item in data.get("qa", []):
            q, a = (item.get("question") or "").strip(), (item.get("answer") or "").strip()
            if not q or not a or normalize(q) in existing:
                skipped += 1
                continue
            qa_id = self.add_qa(q, a, item.get("category", ""), item.get("variants", []))
            if not item.get("enabled", 1):
                self.set_enabled(qa_id, False)
            existing.add(normalize(q))
            added += 1
        have_g = {normalize(g["text"]) for g in self.list_greetings()}
        for g in data.get("greetings", []):
            if g.strip() and normalize(g) not in have_g:
                self.add_greeting(g)
        have_f = {(normalize(f["label"])) for f in self.list_facts()}
        for f in data.get("facts", []):
            if f.get("label") and f.get("value") and normalize(f["label"]) not in have_f:
                self.add_fact(f["label"], f["value"])
        return {"added": added, "skipped": skipped}

    def seed_sample_data(self, office: str, assistant: str) -> None:
        """First-run demo content (category 'sample') so the app works immediately. Delete freely."""
        if self.list_qa() or self.list_facts() or self.list_greetings():
            return
        self.add_greeting(f"Welcome to {office}! I am {assistant}, your virtual receptionist. How may I help you today?")
        self.add_fact("Opening hours (sample)", "Monday to Saturday, 9 AM to 6 PM. Closed on Sundays.")
        self.add_fact("Phone (sample)", "+91 00000 00000")
        samples = [
            ("What are your opening hours?", "We are open Monday to Saturday, from 9 AM to 6 PM.",
             ["when do you open", "what time do you close", "are you open on sunday", "office timings"]),
            ("Where is the office located?", "Please update this sample answer with your office address.",
             ["what is your address", "how do I reach you", "where are you"]),
            ("Where is the restroom?", "The restroom is down the corridor on your left.",
             ["where is the washroom", "toilet please"]),
            ("Is there parking available?", "Yes, visitor parking is available in front of the building.",
             ["where can I park my car", "do you have a parking lot"]),
        ]
        for q, a, v in samples:
            self.add_qa(q, a, "sample", v)
