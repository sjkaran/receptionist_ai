"""Fuzzy matching of visitor questions against trained Q&A (same mechanism as the original)."""
from difflib import SequenceMatcher

from core import database as db


def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, a.lower().strip(), b.lower().strip()).ratio()


def find_best_answer(question: str, threshold: float):
    """Return (answer, qa_id, score). answer is None when nothing clears the threshold."""
    best_score, best = 0.0, None
    q_words = set(question.lower().split())

    for row in db.list_qa():
        score = similarity(question, row["question"])
        t_words = set(row["question"].lower().split())
        if t_words:
            score = min(1.0, score + len(q_words & t_words) / len(t_words) * 0.3)
        if score > best_score:
            best_score, best = score, row

    if best and best_score >= threshold:
        return best["answer"], best["id"], best_score
    return None, None, best_score
