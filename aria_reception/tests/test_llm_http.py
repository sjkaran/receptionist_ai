"""Real OllamaClient + Receptionist against the mock HTTP server (no GUI needed)."""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import mock_ollama
from aria.config import Settings
from aria.context import Context
from aria.llm import OllamaClient, OllamaError


class HttpTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.srv, cls.url = mock_ollama.start()

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown()

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.s = Settings()
        self.s.ollama_url = self.url
        self.s.db_path = os.path.join(self.tmp.name, "t.db")
        self.s.speak_replies = False
        mock_ollama.Handler.log.clear()

    def tearDown(self):
        self.tmp.cleanup()

    def test_auto_picks_mistral_and_streams(self):
        c = OllamaClient(self.s)
        self.assertEqual(c.resolve_model(), "mistral:7b-instruct")
        chunks = list(c.chat_stream([{"role": "user", "content": "what is your phone"}]))
        self.assertGreater(len(chunks), 3)
        self.assertIn("12345", "".join(chunks))

    def test_explicit_model_and_missing_model_error(self):
        self.s.ollama_model = "llama3"
        self.assertEqual(OllamaClient(self.s).resolve_model(), "llama3:8b")
        self.s.ollama_model = "nonexistent"
        with self.assertRaises(OllamaError) as cm:
            OllamaClient(self.s).resolve_model()
        self.assertIn("Installed", str(cm.exception))

    def test_unreachable_server_gives_helpful_error(self):
        self.s.ollama_url = "http://127.0.0.1:9"
        with self.assertRaises(OllamaError) as cm:
            OllamaClient(self.s).models()
        self.assertIn("ollama serve", str(cm.exception))

    def test_end_to_end_through_context(self):
        ctx = Context.create(self.s)
        r = ctx.brain.respond("how can I contact you by phone")
        self.assertEqual(r.source, "ai")
        self.assertIn("12345", r.text)
        chat = [req for path, req in mock_ollama.Handler.log if path == "/api/chat"][0]
        self.assertEqual(chat["model"], "mistral:7b-instruct")
        self.assertIn("OFFICE FACTS", chat["messages"][0]["content"])
        self.assertEqual(ctx.brain.respond("do you sell insurance").source, "fallback")
        self.assertEqual(ctx.brain.respond("what are your opening hours").source, "trained")

    def test_semantic_embeddings_are_cached_and_used(self):
        self.s.embed_model = "nomic-embed-text"
        ctx = Context.create(self.s)
        m = ctx.retriever.best("when do you open")
        self.assertTrue(ctx.retriever.semantic_active)
        self.assertIsNotNone(m.semantic)
        n_first = sum(1 for p, _ in mock_ollama.Handler.log if p == "/api/embed")
        ctx.retriever.best("when do you open")                      # query vector now cached
        ctx.retriever.best("tell me the closing time please")
        n_after = sum(1 for p, _ in mock_ollama.Handler.log if p == "/api/embed")
        self.assertEqual(n_after - n_first, 1)                      # only the new query was embedded

    def test_semantic_failure_degrades_to_lexical(self):
        self.s.embed_model = "does-not-exist"
        ctx = Context.create(self.s)
        m = ctx.retriever.best("where is the washroom")
        self.assertFalse(ctx.retriever.semantic_active)
        self.assertIn("restroom", m.question.lower())


if __name__ == "__main__":
    unittest.main(verbosity=2)
