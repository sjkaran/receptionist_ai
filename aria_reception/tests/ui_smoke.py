"""Drives the real Tk UI against the mock Ollama server.  Needs a display (xvfb-run on Linux).

    xvfb-run -a python tests/ui_smoke.py [screenshot_dir]
"""
import os
import sys
import tempfile
import time
from tkinter import messagebox

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tests"))

import mock_ollama
from aria.config import Settings
from aria.context import Context
from aria.ui.app import App

SHOTS = sys.argv[1] if len(sys.argv) > 1 else None
dialogs = []
for name in ("showinfo", "showwarning", "showerror"):
    setattr(messagebox, name, lambda *a, _n=name, **k: dialogs.append((_n, a)))
messagebox.askyesno = lambda *a, **k: (dialogs.append(("askyesno", a)), True)[1]

fails = []


def check(cond, label):
    print(("  ok   " if cond else "  FAIL ") + label)
    if not cond:
        fails.append(label)


def settle(app, secs=1.2):
    end = time.time() + secs
    while time.time() < end:
        app.update()
        time.sleep(0.02)


def shot(app, name):
    if not SHOTS:
        return
    from PIL import ImageGrab
    settle(app, 0.3)
    app.update_idletasks()
    x, y, w, h = app.winfo_rootx(), app.winfo_rooty(), app.winfo_width(), app.winfo_height()
    ImageGrab.grab(bbox=(x, y, x + w, y + h), xdisplay=os.environ.get("DISPLAY")).save(os.path.join(SHOTS, f"{name}.png"))


def transcript(app):
    return app.reception.log.get("1.0", "end")


def say(app, text, wait=1.0):
    app.reception.entry.insert(0, text)
    app.reception.send()
    settle(app, wait)


srv, url = mock_ollama.start()
tmp = tempfile.mkdtemp()
s = Settings()
s.ollama_url, s.db_path, s.speak_replies = url, os.path.join(tmp, "ui.db"), False
s.save(os.path.join(tmp, "settings.json"))
s = Settings.load(os.path.join(tmp, "settings.json"))
ctx = Context.create(s)
app = App(ctx)
settle(app, 0.8)

print("Reception")
check("Ollama" in app.status_var.get() and "mistral" in app.status_var.get(), f"status shows model: {app.status_var.get()!r}")
say(app, "where is the restroom")
check("trained answer" in transcript(app) and "corridor" in transcript(app).lower() or "sample" in transcript(app).lower() or "restroom" in transcript(app).lower(), "trained answer shown")
say(app, "how can I contact you by phone", 1.5)
check("12345" in transcript(app) and "local AI" in transcript(app), "LLM answer streamed into transcript")
check("sent to Review" in transcript(app), "AI answer queued for review")
say(app, "do you sell insurance", 1.5)
check("not known" in transcript(app) and "UNKNOWN" not in transcript(app), "sentinel hidden, honest fallback")
for line in ("yes", "Priya Nair", "9876543210", "please call me about a quote", "yes"):
    say(app, line, 0.5)
settle(app, 0.5)
msgs = ctx.db.list_messages()
check(len(msgs) == 1 and msgs[0]["visitor"] == "Priya Nair", "message captured through the UI")
shot(app, "1_reception")

print("PIN gate")
from tkinter import simpledialog           # the real prompt is modal; emulate typed PINs
simpledialog.askstring = lambda *a, **k: "0000"
app.nb.select(app.training)
settle(app, 0.2)
check(app.nb.select() == str(app.reception), "wrong PIN keeps trainer tabs locked")
simpledialog.askstring = lambda *a, **k: "1234"
app.nb.select(app.training)
settle(app, 0.3)
check(app.unlocked and app.nb.select() == str(app.training), "correct PIN unlocks")

print("Training")
t = app.training
n0 = len(t.tree.get_children())
t.q_entry.insert(0, "Do you offer guided tours?")
from aria.ui import widgets as w
w.set_text(t.a_text, "Yes, tours run every day at 11 AM.")
t.cat_entry.insert(0, "visits")
t.suggest()
settle(app, 1.2)
check("when do you open" in w.get_text(t.v_text), "AI variant suggestions inserted")
t.save()
settle(app, 0.3)
check(len(t.tree.get_children()) == n0 + 1 and t.save_btn.cget("text") == "Update", "Q&A saved, form switches to Update")
t.t_entry.insert(0, "do you give tours")
t.test()
settle(app, 0.8)
out = t.t_out.get("1.0", "end")
check("guided tours" in out.lower() and "Result:" in out, "match tester shows ranked result + verdict")
shot(app, "2_training")
say(app, "do you give tours", 0.8)            # new knowledge immediately usable (typed from Reception)
app.nb.select(app.reception)
settle(app, 0.3)

print("Review")
app.nb.select(app.review)
settle(app, 0.4)
rv = app.review
check(len(rv.tree.get_children()) >= 2, f"review queue populated ({len(rv.tree.get_children())})")
shot(app, "3_review")
pending0 = ctx.db.pending_count()
w.set_text(rv.a_text, "You can call us on 12345.")
rv.q_entry.delete(0, "end")
rv.q_entry.insert(0, "What is your phone number?")
rv.approve()
settle(app, 0.3)
check(ctx.db.pending_count() == pending0 - 1, "approve trains + clears queue item")
check(ctx.retriever.best("how can I contact you by phone").score >= s.match_threshold or
      ctx.retriever.best("what is your phone number").score >= s.match_threshold, "approved item now answers directly")

print("Knowledge / Messages / Settings")
app.nb.select(app.knowledge)
k = app.knowledge
k.f_label.insert(0, "Wi-Fi")
k.f_value.insert(0, "Network GUEST, ask the desk for the password")
k.add_fact()
w.set_text(k.g_text, "Good day! You have reached our front desk.")
k.add_greeting()
check(len(ctx.db.list_facts()) == 3 and len(ctx.db.list_greetings()) == 2, "facts + greetings added")
shot(app, "4_knowledge")
app.nb.select(app.messages)
settle(app, 0.3)
check(len(app.messages.tree.get_children()) == 1, "message listed")
app.messages.tree.selection_set(app.messages.tree.get_children()[0])
app.messages.mark(True)
check(ctx.db.unhandled_count() == 0, "message marked handled")
shot(app, "5_messages")
app.nb.select(app.settings)
settle(app, 0.5)
st = app.settings
st.v["office_name"].set("ACME Corp")
st.v["match_threshold"].set("0.75")
st.new_pin.set("4321")
st.new_pin2.set("4321")
st.save()
settle(app, 0.3)
check(ctx.settings.office_name == "ACME Corp" and ctx.settings.check_pin("4321"), "settings applied + PIN changed")
check(Settings.load(os.path.join(tmp, "settings.json")).office_name == "ACME Corp", "settings persisted to disk")
check("ACME" in app.title(), "window title refreshed")
st.test()
settle(app, 1.2)
check(st.test_var.get().startswith("OK"), f"connection test: {st.test_var.get()!r}")
shot(app, "6_settings")

app.lock()
check(not app.unlocked and app.nb.select() == str(app.reception), "Lock returns to Reception")
bad = [d for d in dialogs if d[0] in ("showerror", "showwarning") and d[1][0] != "Access denied"]
check(not bad, f"no error/warning dialogs ({bad})")
app._close()
srv.shutdown()
print("\nFAILED: " + ", ".join(fails) if fails else "\nUI SMOKE TEST PASSED")
sys.exit(1 if fails else 0)
