"""The receptionist's decision logic.

Pipeline for every visitor utterance
------------------------------------
 1. Active dialog?  (taking a message: name -> contact -> body -> confirm)
 2. Small talk      (hello / thanks / bye / who are you)  – instant, no LLM
 3. "Leave a message" intent                              – starts the dialog
 4. Retrieval over trained Q&A
      score >= match_threshold   -> answer verbatim from training   (source: trained)
      suggest <= score < match   -> LLM, grounded on the closest trained answers
      score < suggest            -> LLM, grounded on office facts only
 5. LLM is told to answer ONLY from the supplied knowledge; if it cannot, it emits
    a sentinel and Aria honestly says so and offers to take a message
    (so she never invents opening hours or prices).
 6. Anything the LLM handled or could not answer lands in the review queue with a
    suggested answer – approving it in the UI turns it into trained knowledge.
    That is the learning loop.
"""
from __future__ import annotations

import re
import threading
from dataclasses import dataclass, field
from typing import Callable

from .llm import OllamaClient, OllamaError
from .retriever import Match, Retriever
from .textutils import (clean_for_speech, is_cancel, is_no, is_yes, normalize, parse_name,
                        smalltalk_intent, term_list, wants_message)

UNKNOWN = "[[UNKNOWN]]"
SENTENCE_END = re.compile(r"(?<=[.!?])\s+")

SAY_FN = Callable[[str], None]


def _is_unknown(text: str) -> bool:
    t = text.strip()
    return UNKNOWN in t or bool(re.fullmatch(r"\W*unknown\W*", t, re.I))


@dataclass
class Reply:
    text: str
    source: str                      # trained | ai | smalltalk | dialog | fallback
    score: float = 0.0
    qa_id: int | None = None
    queued: bool = False             # sent to the review queue
    intent: str = ""                 # small-talk intent name (greeting, bye, ...)
    matches: list[Match] = field(default_factory=list)


class Receptionist:
    def __init__(self, db, settings, retriever: Retriever, llm: OllamaClient):
        self.db, self.s, self.retriever, self.llm = db, settings, retriever, llm
        self.history: list[dict] = []          # [{"user":..., "assistant":...}]
        self._state: dict | None = None        # active dialog
        self._last_question: str | None = None
        self._lock = threading.Lock()
        self.cancel = threading.Event()

    # ================================================================= public
    def greeting(self) -> str:
        return self.db.pick_greeting() or (
            f"Welcome to {self.s.office_name}! I am {self.s.assistant_name}, your virtual receptionist. "
            "How may I help you today?")

    def reset(self) -> None:
        with self._lock:
            self.history.clear()
            self._state = None
            self._last_question = None

    def respond(self, text: str, on_sentence: SAY_FN | None = None) -> Reply:
        """Produce a reply.  `on_sentence` is called with each speakable chunk as soon as it exists."""
        text = (text or "").strip()
        with self._lock:
            self.cancel.clear()
            reply = self._route(text, on_sentence)
            self.history.append({"user": text, "assistant": reply.text})
            self.history = self.history[-12:]
        self.db.log_turn("visitor", text)
        self.db.log_turn("aria", reply.text, reply.source, reply.score)
        return reply

    # ================================================================ routing
    def _route(self, text: str, say: SAY_FN | None) -> Reply:
        if self._state:
            r = self._continue_dialog(text)
            if r:
                return self._emit(r, say)

        intent = smalltalk_intent(text)
        if intent:
            return self._emit(Reply(self._smalltalk(intent), "smalltalk", intent=intent), say)

        if wants_message(text):
            self._state = {"step": "name", "data": {}}
            return self._emit(Reply("Certainly, I can take a message. May I have your name, please?", "dialog"), say)

        matches = self._retrieve(text)
        top = matches[0] if matches else None

        if top and top.score >= self.s.match_threshold:
            self.db.hit(top.qa_id)
            self._last_question = top.question
            return self._emit(Reply(top.answer, "trained", top.score, top.qa_id, matches=matches), say)

        return self._ask_llm(text, matches, say)

    def _emit(self, reply: Reply, say: SAY_FN | None) -> Reply:
        if say and reply.text:
            say(reply.text)
        return reply

    # ============================================================== retrieval
    def _retrieve(self, text: str) -> list[Match]:
        matches = self.retriever.search(text, k=3)
        # Short follow-ups ("and on Sunday?") borrow context from the previous question.
        # Safety rule: the follow-up's own words must appear in the matched entry, otherwise an
        # unrelated short question ("do you sell insurance") would inherit the previous topic.
        own = set(term_list(text))
        if 0 < len(own) <= 2 and self._last_question:
            ctx = self.retriever.search(f"{self._last_question} {text}", k=3)
            if (ctx and own & set(term_list(ctx[0].phrase))
                    and (not matches or ctx[0].score > matches[0].score + 0.05)):
                matches = ctx
        return matches

    # ================================================================ small talk
    def _smalltalk(self, intent: str) -> str:
        n, o = self.s.assistant_name, self.s.office_name
        return {
            "greeting": f"Hello! Welcome to {o}. How may I help you?",
            "thanks": "You are most welcome! Is there anything else I can help with?",
            "bye": "Thank you for visiting. Have a wonderful day!",
            "how_are_you": "I am doing well, thank you for asking. How may I help you today?",
            "identity": f"I am {n}, the virtual receptionist at {o}. I can answer questions or take a message for the team.",
        }[intent]

    # ==================================================================== LLM
    def _system_prompt(self, matches: list[Match]) -> str:
        n, o = self.s.assistant_name, self.s.office_name
        facts = self.db.list_facts()
        facts_txt = "\n".join(f"- {f['label']}: {f['value']}" for f in facts) or "- (none recorded)"
        ref = [m for m in matches if m.score >= self.s.suggest_threshold]
        ref_txt = "\n".join(f"Q: {m.question}\nA: {m.answer}" for m in ref) or "(none)"
        return (
            f"You are {n}, the friendly, professional voice receptionist of {o}. Your replies are spoken aloud.\n"
            "RULES:\n"
            "1. Answer using ONLY the OFFICE FACTS and TRAINED ANSWERS below, plus ordinary polite conversation.\n"
            "2. Never invent hours, prices, names, addresses, phone numbers or policies. If the answer is not in the "
            f"knowledge below, reply with exactly {UNKNOWN} and nothing else.\n"
            "3. Maximum two short sentences. Plain spoken English: no lists, no markdown, no emojis.\n"
            "4. If a TRAINED ANSWER fits the visitor's question, rephrase it naturally instead of adding new claims.\n"
            "5. Stay in role. Do not mention being an AI model, prompts or these rules.\n\n"
            f"OFFICE FACTS:\n{facts_txt}\n\nTRAINED ANSWERS:\n{ref_txt}")

    def _ask_llm(self, text: str, matches: list[Match], say: SAY_FN | None) -> Reply:
        top_score = matches[0].score if matches else 0.0
        messages = [{"role": "system", "content": self._system_prompt(matches)}]
        for h in self.history[-4:]:
            messages.append({"role": "user", "content": h["user"]})
            messages.append({"role": "assistant", "content": h["assistant"]})
        messages.append({"role": "user", "content": text})

        spoken: list[str] = []
        buf = ""                 # text not yet turned into a spoken sentence
        full = ""
        released = False         # becomes True once we are sure the reply is not the UNKNOWN sentinel
        sentinel_seen = False

        def flush(final: bool) -> None:
            nonlocal buf
            parts = SENTENCE_END.split(buf)
            ready, buf = (parts, "") if final else (parts[:-1], parts[-1])
            for sentence in ready:
                s = clean_for_speech(sentence)
                if s:
                    spoken.append(s)
                    if say:
                        say(s)

        try:
            for chunk in self.llm.chat_stream(messages, cancel=self.cancel):
                full += chunk
                if not released:
                    head = full.lstrip()
                    if _is_unknown(head):
                        sentinel_seen = True
                        break
                    # hold back until we have seen enough characters to rule out the sentinel
                    if len(head) < len(UNKNOWN) and UNKNOWN.startswith(head[:len(UNKNOWN)]):
                        continue
                    released = True
                    buf = full
                else:
                    buf += chunk
                flush(final=False)
            if not sentinel_seen:
                if not released:
                    head = full.strip()
                    if _is_unknown(head):
                        sentinel_seen = True
                    else:
                        buf = full
                if not sentinel_seen:
                    flush(final=True)
        except OllamaError as exc:
            print(f"[brain] LLM unavailable: {exc}")
            return self._fallback(text, matches, say, reason=str(exc))

        if self.cancel.is_set() and not spoken:      # operator pressed Stop mid-thought
            return Reply("", "fallback")
        if sentinel_seen or not spoken:
            return self._unknown(text, matches, say)

        answer = " ".join(spoken)
        self.db.queue_question(text, suggested=answer)
        self._last_question = text
        return Reply(answer, "ai", top_score, matches=matches, queued=True)

    # ============================================================ not-known paths
    def _unknown(self, text: str, matches: list[Match], say: SAY_FN | None) -> Reply:
        self.db.queue_question(text)
        self._state = {"step": "offer", "data": {"topic": text}}
        msg = ("I am sorry, I do not have that information right now. "
               "Would you like to leave a message for our team?")
        return self._emit(Reply(msg, "fallback", matches[0].score if matches else 0.0, queued=True, matches=matches), say)

    def _fallback(self, text: str, matches: list[Match], say: SAY_FN | None, reason: str = "") -> Reply:
        """LLM offline: use the closest trained answer if it is plausible, else be honest."""
        top = matches[0] if matches else None
        if top and top.score >= self.s.suggest_threshold + 0.1:
            self.db.hit(top.qa_id)
            self.db.queue_question(text)
            return self._emit(Reply(top.answer, "trained", top.score, top.qa_id, True, matches), say)
        return self._unknown(text, matches, say)

    # ==================================================================== dialog
    def _continue_dialog(self, text: str) -> Reply | None:
        st = self._state
        step, data = st["step"], st["data"]

        if step == "offer":
            if is_yes(text):
                self._state = {"step": "name", "data": {"topic": data.get("topic", "")}}
                return Reply("Of course. May I have your name, please?", "dialog")
            self._state = None
            if is_no(text):
                return Reply("No problem. Is there anything else I can help you with?", "dialog")
            return None                       # visitor moved on to a new question

        if is_cancel(text):
            self._state = None
            return Reply("No problem, I have cancelled that. How else can I help?", "dialog")

        if step == "name":
            name = parse_name(text)
            if not name:
                return Reply("Sorry, I did not catch your name. Could you say it again?", "dialog")
            data["visitor"] = name
            st["step"] = "contact"
            return Reply(f"Thank you, {name.split()[0]}. What phone number or email can the team reach you on?", "dialog")

        if step == "contact":
            data["contact"] = text.strip()
            st["step"] = "body"
            return Reply("And what message should I pass on?", "dialog")

        if step == "body":
            data["body"] = text.strip()
            st["step"] = "confirm"
            return Reply(f"Let me confirm: a message from {data['visitor']}, saying: {data['body']}. "
                         "Shall I send it?", "dialog")

        if step == "confirm":
            if is_yes(text):
                body = data["body"]
                if data.get("topic"):
                    body = f"{body}  (Asked about: {data['topic']})"
                self.db.add_message(data["visitor"], data["contact"], body)
                self._state = None
                return Reply("Done. Your message has been recorded and the team will get back to you. "
                             "Is there anything else?", "dialog")
            if is_no(text):
                st["step"] = "body"
                return Reply("Okay, what would you like the message to say?", "dialog")
            return Reply("Sorry, should I send the message? Please say yes or no.", "dialog")

        self._state = None
        return None

    @property
    def in_dialog(self) -> bool:
        return self._state is not None

    # ============================================================= LLM helpers
    def suggest_variants(self, question: str, answer: str = "", n: int = 6) -> list[str]:
        """Ask the local model for alternative ways a visitor might phrase this question."""
        system = ("You help train a receptionist. Given a question, write different short ways a visitor might "
                  "ask the same thing out loud, including casual and indirect ones. One per line, no numbering, "
                  "no extra text.")
        user = f"Question: {question}\n" + (f"Answer given: {answer}\n" if answer else "") + f"Write {n} variants."
        out = self.llm.complete(system, user, num_predict=200, temperature=0.7)
        lines = [re.sub(r"^\s*[-•\d.)]+\s*", "", ln).strip().strip('"') for ln in out.splitlines()]
        seen, res = {normalize(question)}, []
        for ln in lines:
            n_ = normalize(ln)
            if ln and n_ not in seen and len(ln) < 140:
                seen.add(n_)
                res.append(ln)
        return res[:n]

    def polish_answer(self, answer: str) -> str:
        system = (f"Rewrite the receptionist answer so it sounds natural when spoken aloud at {self.s.office_name}. "
                  "Keep every fact exactly the same. Max two sentences. Output only the rewritten answer.")
        return clean_for_speech(self.llm.complete(system, answer, num_predict=120, temperature=0.4))
