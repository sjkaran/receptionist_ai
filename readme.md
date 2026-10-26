# Reception Assistant

A trainable voice receptionist. Visitors speak, the assistant answers from your trained
Q&A (fuzzy matching). Anything it hasn't been taught goes to your local Ollama **mistral** model.

## Setup (Windows)
1. Install Python 3.10+ and [Ollama](https://ollama.com).
2. `ollama pull mistral` (Ollama must be running).
3. Double-click `run.bat` (creates a venv, installs requirements, opens the UI at http://127.0.0.1:8000).

Speech recognition (Google) needs internet. Everything else runs locally.

## Structure
```
app.py              entry point (starts server, opens browser)
config.py           settings stored in data/settings.json (editable in the UI)
core/
  database.py       SQLite: greetings, qa_pairs, conversations log
  matcher.py        fuzzy match (SequenceMatcher + keyword boost)
  tts.py            Windows SAPI voice (fresh COM thread per utterance)
  stt.py            sounddevice recording + Google recognition
  llm.py            Ollama client (/api/chat) + health check
  assistant.py      reception loop, state machine, event feed
web/
  server.py         FastAPI JSON API
  static/           index.html, style.css, app.js
data/               reception.db, settings.json (created on first run)
```

## Flow
visitor speaks -> stt -> matcher -> trained answer? yes: speak it / no: Ollama -> speak -> log

## Using it
- **Reception desk**: start/stop, live conversation feed, and a typed test box.
- **Training** (PIN, default 1234): greeting, Q&A add/edit/delete (typed or dictated), and a list of
  questions the AI answered so you can promote them to trained answers.
- **Settings**: names, "About the office" facts for the AI, Ollama address/model, match strictness,
  mic sensitivity, voice, PIN.

## Ideas for the next upgrade
Semantic matching (embeddings via Ollama), offline speech-to-text (Whisper/Vosk), barge-in/stop speaking,
visitor intake (name, who they're meeting) and staff notifications, multi-language, mic device picker.
