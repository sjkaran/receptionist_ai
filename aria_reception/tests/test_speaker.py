"""Speaker queue / stop semantics, using a fake backend (no audio hardware needed)."""
import os
import sys
import threading
import time
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from aria.config import Settings
from aria.speech import Speaker


class FakeSpeaker(Speaker):
    spoken: list
    started: threading.Event

    @staticmethod
    def _detect():
        return "fake"

    def _speak(self, text, gen):
        self.started.set()
        for _ in range(20):                      # ~0.2 s, interruptible exactly like the real backends
            if gen != self._gen:
                return
            time.sleep(0.01)
        self.spoken.append(text)


def make():
    s = Settings()
    s.speak_replies = True
    FakeSpeaker.spoken, FakeSpeaker.started = [], threading.Event()
    return FakeSpeaker(s)


class SpeakerTests(unittest.TestCase):
    def test_queue_speaks_in_order_and_goes_idle(self):
        sp = make()
        for t in ("one", "two", "three"):
            sp.say(t)
        self.assertTrue(sp.wait_idle(5))
        self.assertEqual(sp.spoken, ["one", "two", "three"])

    def test_stop_interrupts_and_drops_queue(self):
        sp = make()
        for t in ("a", "b", "c"):
            sp.say(t)
        sp.started.wait(2)
        sp.stop()
        self.assertTrue(sp.wait_idle(5))
        self.assertEqual(sp.spoken, [])

    def test_next_utterance_after_stop_while_idle_is_not_swallowed(self):
        sp = make()
        sp.stop()                                # nothing playing – must not poison the next say()
        sp.say("hello")
        self.assertTrue(sp.wait_idle(5))
        self.assertEqual(sp.spoken, ["hello"])

    def test_works_again_after_interrupting_speech(self):
        sp = make()
        sp.say("long one")
        sp.started.wait(2)
        sp.stop()
        sp.say("fresh")
        self.assertTrue(sp.wait_idle(5))
        self.assertEqual(sp.spoken, ["fresh"])

    def test_disabled_speaking_is_silent_and_idle(self):
        sp = make()
        sp.s.speak_replies = False
        sp.say("quiet")
        self.assertTrue(sp.wait_idle(1))
        self.assertEqual(sp.spoken, [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
