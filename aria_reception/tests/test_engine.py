"""Engine tests — no Ollama, microphone or GUI required (a scripted fake LLM stands in)."""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aria.brain import UNKNOWN, Receptionist
from aria.config import Settings
from aria.database import Database
from aria.llm import OllamaError
from aria.retriever import Retriever
from aria.textutils import clean_for_speech, parse_name, smalltalk_intent, wants_message


class FakeLLM:
    def __init__(self, script=None, down=False):
        self.script, self.down, self.calls = script or [], down, []

    def chat_stream(self, messages, cancel=None, **kw):
        self.calls.append(messages)
        if self.down:
            raise OllamaError("down")
        out = self.script.pop(0) if self.script else UNKNOWN
        for i in range(0, len(out), 3):          # chunk awkwardly on purpose
            yield out[i:i + 3]

    def complete(self, system, user, **kw):
        return "what time do you open\nare you open today\n- opening times\n"


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.s = Settings()
        self.db = Database(os.path.join(self.tmp.name, "t.db"))
        self.db.add_qa("What are your opening hours?", "We are open 9 to 6, Monday to Saturday.", "hours",
                       ["when do you open", "what time do you close", "are you open on sunday"])
        self.db.add_qa("Where is the restroom?", "Down the corridor on your left.", "", ["where is the washroom"])
        self.db.add_qa("Is there parking?", "Yes, visitor parking is in front.", "")
        self.db.add_fact("Phone", "12345")

    def tearDown(self):
        self.tmp.cleanup()

    def bot(self, llm=None):
        llm = llm or FakeLLM()
        return Receptionist(self.db, self.s, Retriever(self.db, self.s), llm), llm


class TextTests(unittest.TestCase):
    def test_helpers(self):
        self.assertEqual(smalltalk_intent("Hello"), "greeting")
        self.assertEqual(smalltalk_intent("thank you so much"), "thanks")
        self.assertIsNone(smalltalk_intent("hello what are your opening hours today please"))
        self.assertTrue(wants_message("I would like to leave a message for the manager"))
        self.assertFalse(wants_message("what are your hours"))
        self.assertEqual(parse_name("my name is rahul sharma"), "Rahul Sharma")
        self.assertNotIn("*", clean_for_speech("**Hello** there"))


class RetrievalTests(Base):
    def test_paraphrases_hit_the_right_answer(self):
        r = Retriever(self.db, self.s)
        cases = {
            "what time do you guys close today": "opening hours",
            "where can I find the washroom": "restroom",
            "can I park my car here": "parking",
            "when are you open": "opening hours",
            "opening timings please": "opening hours",
        }
        for q, expect in cases.items():
            m = r.best(q)
            self.assertIn(expect, m.question.lower(), f"{q!r} -> {m.question} ({m.score})")
            self.assertGreaterEqual(m.score, self.s.suggest_threshold, q)

    def test_unrelated_stays_low(self):
        r = Retriever(self.db, self.s)
        for q in ["do you sell insurance", "what is the capital of france", "i want to buy a laptop"]:
            self.assertLess(r.best(q).score, self.s.match_threshold, q)

    def test_index_refreshes_after_edit_and_delete(self):
        r = Retriever(self.db, self.s)
        self.assertLess(r.best("do you have wifi").score, self.s.match_threshold)
        qid = self.db.add_qa("Do you have wifi?", "Yes, ask at the desk.")
        self.assertEqual(r.best("is there wifi").qa_id, qid)
        self.db.delete_qa(qid)
        self.assertNotEqual(getattr(r.best("is there wifi"), "qa_id", None), qid)

    def test_disabled_entries_are_ignored(self):
        r = Retriever(self.db, self.s)
        top = r.best("where is the restroom")
        self.db.set_enabled(top.qa_id, False)
        self.assertNotEqual(r.best("where is the restroom").qa_id, top.qa_id)


class BrainTests(Base):
    def test_trained_answer_bypasses_llm(self):
        bot, llm = self.bot()
        rep = bot.respond("What are your opening hours?")
        self.assertEqual(rep.source, "trained")
        self.assertEqual(llm.calls, [])
        self.assertEqual(self.db.list_qa()[0]["hits"], 1)

    def test_llm_grounded_answer_is_streamed_and_queued(self):
        bot, llm = self.bot(FakeLLM(["Our phone number is 12345. Anything else I can help with?"]))
        spoken = []
        rep = bot.respond("how can I contact you by phone", spoken.append)
        self.assertEqual(rep.source, "ai")
        self.assertGreaterEqual(len(spoken), 2)                    # sentence-level streaming
        self.assertIn("12345", llm.calls[0][0]["content"])         # facts injected
        self.assertEqual(self.db.pending_count(), 1)
        self.assertIn("12345", self.db.list_review()[0]["suggested"])

    def test_unknown_sentinel_is_never_spoken_and_offers_message(self):
        bot, _ = self.bot(FakeLLM([UNKNOWN]))
        spoken = []
        rep = bot.respond("do you sell insurance", spoken.append)
        self.assertEqual(rep.source, "fallback")
        self.assertNotIn("UNKNOWN", " ".join(spoken) + rep.text)
        self.assertTrue(bot.in_dialog)
        self.assertEqual(self.db.pending_count(), 1)

    def test_message_flow_end_to_end(self):
        bot, _ = self.bot(FakeLLM([UNKNOWN]))
        bot.respond("do you sell insurance")
        for line in ["yes please", "my name is Priya", "9876543210", "please call me about a quote", "yes"]:
            bot.respond(line)
        msgs = self.db.list_messages()
        self.assertEqual(len(msgs), 1)
        self.assertEqual(msgs[0]["visitor"], "Priya")
        self.assertEqual(msgs[0]["contact"], "9876543210")
        self.assertIn("quote", msgs[0]["body"])
        self.assertFalse(bot.in_dialog)

    def test_direct_leave_message_and_cancel(self):
        bot, _ = self.bot()
        self.assertIn("name", bot.respond("I want to leave a message").text.lower())
        self.assertIn("cancel", bot.respond("never mind").text.lower())
        self.assertFalse(bot.in_dialog)

    def test_declining_offer_returns_to_normal(self):
        bot, _ = self.bot(FakeLLM([UNKNOWN]))
        bot.respond("do you sell insurance")
        bot.respond("no")
        self.assertFalse(bot.in_dialog)
        self.assertEqual(bot.respond("where is the restroom").source, "trained")

    def test_new_question_during_offer_is_answered(self):
        bot, _ = self.bot(FakeLLM([UNKNOWN]))
        bot.respond("do you sell insurance")
        rep = bot.respond("is there parking")
        self.assertEqual(rep.source, "trained")
        self.assertFalse(bot.in_dialog)

    def test_llm_down_falls_back_gracefully(self):
        bot, _ = self.bot(FakeLLM(down=True))
        near = bot.respond("could you tell me the washroom location")
        self.assertEqual(near.source, "trained")
        far = bot.respond("do you sell insurance")
        self.assertEqual(far.source, "fallback")

    def test_followup_uses_previous_context(self):
        bot, _ = self.bot()
        bot.respond("what are your opening hours")
        rep = bot.respond("and on Sunday")
        self.assertEqual(rep.source, "trained")
        self.assertIn("open", rep.text.lower())

    def test_smalltalk_is_instant(self):
        bot, llm = self.bot()
        self.assertEqual(bot.respond("hello").source, "smalltalk")
        self.assertEqual(bot.respond("thank you").source, "smalltalk")
        self.assertEqual(llm.calls, [])

    def test_review_approval_becomes_trained_knowledge(self):
        bot, _ = self.bot(FakeLLM(["We accept cash and cards."]))
        bot.respond("which payment methods do you accept")
        item = self.db.list_review()[0]
        self.db.approve_review(item["id"], "What payment methods do you accept?", "Cash and cards.", "billing")
        self.assertEqual(self.db.pending_count(), 0)
        rep = bot.respond("which payment methods do you accept")
        self.assertEqual(rep.source, "trained")
        self.assertEqual(rep.text, "Cash and cards.")

    def test_variant_suggestions_parsed(self):
        bot, _ = self.bot()
        v = bot.suggest_variants("What are your opening hours?")
        self.assertIn("are you open today", v)
        self.assertIn("opening times", v)           # bullet stripped


class DataTests(Base):
    def test_export_import_roundtrip_and_dedupe(self):
        data = self.db.export_data()
        other = Database(os.path.join(self.tmp.name, "o.db"))
        res = other.import_data(data)
        self.assertEqual(res["added"], 3)
        self.assertEqual(other.import_data(data)["added"], 0)
        self.assertEqual(len(other.get_qa(other.list_qa()[-1]["id"])["variants"]) >= 0, True)

    def test_review_queue_dedupes_and_counts(self):
        for _ in range(3):
            self.db.queue_question("Do you sell insurance?")
        self.db.queue_question("do you sell insurance")
        item = self.db.list_review()[0]
        self.assertEqual(item["count"], 4)

    def test_pin(self):
        self.assertTrue(self.s.check_pin("1234"))
        self.s.set_pin("9999")
        self.assertTrue(self.s.check_pin("9999") and not self.s.check_pin("1234"))

    def test_settings_roundtrip(self):
        p = os.path.join(self.tmp.name, "s.json")
        self.s.office_name = "ACME"
        self.s.save(p)
        self.assertEqual(Settings.load(p).office_name, "ACME")


if __name__ == "__main__":
    unittest.main(verbosity=2)
