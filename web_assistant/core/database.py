"""SQLite storage: greetings, Q&A pairs, conversation log."""
import sqlite3
from contextlib import contextmanager

from config import DB_PATH


@contextmanager
def _conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    with _conn() as c:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS greetings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                text TEXT NOT NULL,
                created TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS qa_pairs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                question TEXT NOT NULL,
                answer TEXT NOT NULL,
                hits INTEGER DEFAULT 0,
                created TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS conversations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                question TEXT NOT NULL,
                answer TEXT NOT NULL,
                source TEXT NOT NULL,       -- trained | ollama | fallback
                score REAL DEFAULT 0,
                created TEXT DEFAULT CURRENT_TIMESTAMP
            );
        """)


# ---- greetings ----
def add_greeting(text: str):
    with _conn() as c:
        c.execute("INSERT INTO greetings (text) VALUES (?)", (text,))


def list_greetings():
    with _conn() as c:
        return [dict(r) for r in c.execute("SELECT id, text, created FROM greetings ORDER BY id DESC")]


def delete_greeting(gid: int):
    with _conn() as c:
        c.execute("DELETE FROM greetings WHERE id = ?", (gid,))


# ---- Q&A ----
def add_qa(question: str, answer: str):
    with _conn() as c:
        c.execute("INSERT INTO qa_pairs (question, answer) VALUES (?, ?)", (question, answer))


def list_qa():
    with _conn() as c:
        return [dict(r) for r in c.execute(
            "SELECT id, question, answer, hits, created FROM qa_pairs ORDER BY hits DESC, id DESC")]


def update_qa(qid: int, question: str, answer: str):
    with _conn() as c:
        c.execute("UPDATE qa_pairs SET question = ?, answer = ? WHERE id = ?", (question, answer, qid))


def delete_qa(qid: int):
    with _conn() as c:
        c.execute("DELETE FROM qa_pairs WHERE id = ?", (qid,))


def increment_hit(qid: int):
    with _conn() as c:
        c.execute("UPDATE qa_pairs SET hits = hits + 1 WHERE id = ?", (qid,))


# ---- conversation log ----
def log_conversation(question: str, answer: str, source: str, score: float):
    with _conn() as c:
        c.execute("INSERT INTO conversations (question, answer, source, score) VALUES (?,?,?,?)",
                  (question, answer, source, score))


def recent_conversations(limit: int = 50):
    with _conn() as c:
        return [dict(r) for r in c.execute(
            "SELECT * FROM conversations ORDER BY id DESC LIMIT ?", (limit,))]


def counts():
    with _conn() as c:
        return {
            "qa": c.execute("SELECT COUNT(*) FROM qa_pairs").fetchone()[0],
            "greetings": c.execute("SELECT COUNT(*) FROM greetings").fetchone()[0],
            "conversations": c.execute("SELECT COUNT(*) FROM conversations").fetchone()[0],
        }
