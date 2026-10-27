"""FastAPI backend: JSON API + static UI."""
from pathlib import Path
from typing import Optional

from fastapi import Depends, FastAPI, Header, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from web_version.config import DEFAULTS, settings
from web_version.core import database as db
from web_version.core.assistant import assistant
from web_version.core.stt import STTError
from web_version.core.tts import HAVE_SAPI

STATIC = Path(__file__).parent / "static"
app = FastAPI(title="Reception Assistant")


def require_pin(x_pin: str = Header("")):
    if x_pin != str(settings["trainer_pin"]):
        raise HTTPException(401, "Incorrect trainer PIN")


class Pin(BaseModel):
    pin: str

class Text(BaseModel):
    text: str

class QA(BaseModel):
    question: str
    answer: str

class Ask(BaseModel):
    text: str
    speak: bool = False


# ---------- public ----------
@app.get("/")
def index():
    return FileResponse(STATIC / "index.html")


@app.get("/api/status")
def status():
    return {
        "state": assistant.state,
        "running": assistant.running,
        "ollama": assistant.llm.status(),
        "tts": HAVE_SAPI,
        "counts": db.counts(),
        "office": settings["office_name"],
        "assistant": settings["assistant_name"],
    }


@app.get("/api/events")
def events(after: int = 0):
    items, last = assistant.events_after(after)
    return {"events": items, "last": last}


@app.post("/api/reception/start")
def reception_start():
    assistant.start()
    return {"running": True}


@app.post("/api/reception/stop")
def reception_stop():
    assistant.stop()
    return {"running": False}


@app.post("/api/ask")
def ask(body: Ask):
    """Type a question to test the assistant without the microphone."""
    text = body.text.strip()
    if not text:
        raise HTTPException(400, "Type a question first")
    result = assistant.respond(text)
    assistant.emit("visitor", text, typed=True)
    assistant.emit("assistant", result["answer"], source=result["source"], score=result["score"])
    if body.speak:
        assistant.say(result["answer"])
    return result


@app.post("/api/auth")
def auth(body: Pin):
    if body.pin != str(settings["trainer_pin"]):
        raise HTTPException(401, "Incorrect trainer PIN")
    return {"ok": True}


@app.get("/api/settings")
def get_settings():
    return settings.public()


# ---------- trainer only ----------
@app.put("/api/settings", dependencies=[Depends(require_pin)])
def put_settings(patch: dict):
    new_pin = str(patch.get("trainer_pin", "")).strip()
    if not new_pin:
        patch.pop("trainer_pin", None)
    settings.update({k: v for k, v in patch.items() if k in DEFAULTS})
    return settings.public()


@app.get("/api/greetings", dependencies=[Depends(require_pin)])
def greetings():
    return db.list_greetings()


@app.post("/api/greetings", dependencies=[Depends(require_pin)])
def add_greeting(body: Text):
    if not body.text.strip():
        raise HTTPException(400, "Greeting is empty")
    db.add_greeting(body.text.strip())
    return {"ok": True}


@app.delete("/api/greetings/{gid}", dependencies=[Depends(require_pin)])
def del_greeting(gid: int):
    db.delete_greeting(gid)
    return {"ok": True}


@app.get("/api/qa", dependencies=[Depends(require_pin)])
def qa_list():
    return db.list_qa()


@app.post("/api/qa", dependencies=[Depends(require_pin)])
def qa_add(body: QA):
    if not body.question.strip() or not body.answer.strip():
        raise HTTPException(400, "Both question and answer are required")
    db.add_qa(body.question.strip(), body.answer.strip())
    return {"ok": True}


@app.put("/api/qa/{qid}", dependencies=[Depends(require_pin)])
def qa_update(qid: int, body: QA):
    db.update_qa(qid, body.question.strip(), body.answer.strip())
    return {"ok": True}


@app.delete("/api/qa/{qid}", dependencies=[Depends(require_pin)])
def qa_delete(qid: int):
    db.delete_qa(qid)
    return {"ok": True}


@app.post("/api/listen", dependencies=[Depends(require_pin)])
def dictate():
    """Capture one phrase from the microphone (for filling training forms by voice)."""
    try:
        return {"text": assistant.capture_once()}
    except RuntimeError as e:
        raise HTTPException(409 if "reception" in str(e) or "already" in str(e) else 502, str(e))
    except STTError as e:
        raise HTTPException(502, str(e))


@app.post("/api/speak", dependencies=[Depends(require_pin)])
def speak(body: Text):
    assistant.say(body.text)
    return {"ok": True}


@app.get("/api/conversations", dependencies=[Depends(require_pin)])
def conversations(limit: int = 50):
    return db.recent_conversations(limit)


app.mount("/static", StaticFiles(directory=STATIC), name="static")
