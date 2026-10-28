# Aria — trainable local AI receptionist

A voice + text receptionist that **learns your office**, runs **100 % locally** on your Ollama model
(Mistral), and never makes things up. It replaces the old Gemini / DialoGPT fallback.

```
visitor says something
        │
        ▼
 ┌──────────────┐  active dialog? ──► take-a-message flow (name → contact → message → confirm)
 │   Receptionist│
 └──────┬───────┘
        │ small talk (hi / thanks / bye) ─────────────────────────────► instant canned reply
        ▼
 hybrid retrieval over your trained Q&A  (keywords + synonyms + typo-tolerant + optional embeddings)
        │
        ├─ score ≥ 0.70 ─► answer verbatim from training            (fast, 100 % yours)
        ├─ 0.40 – 0.70 ─► Mistral, grounded on the closest answers  (natural rephrase)
        └─ < 0.40 ──────► Mistral, grounded on your Office Facts only
                              │
                              ├─ knows it  ─► speaks it, and queues it in REVIEW with its answer
                              └─ doesn't   ─► "I don't have that — leave a message?"  + queued in REVIEW
```

The **Review** tab is the learning loop: every question Aria could not answer from training lands there,
sorted by how often visitors ask it. Fix the wording, approve, and she knows it from then on.

## Quick start (Windows)

```bat
:: 1. Python 3.10+ (tkinter is included with the python.org installer)
pip install -r requirements.txt

:: 2. Ollama running with your Mistral model
ollama serve
ollama list                      :: note the exact model name, e.g. mistral:latest
ollama pull nomic-embed-text     :: optional – adds paraphrase-level matching

:: 3. Launch
python main.py                   :: desktop app
python main.py --cli             :: or a terminal chat for quick tests
```

Open **Settings** → pick your chat model (leave blank to auto-select any model whose name contains
"mistral"), optionally type `nomic-embed-text` as the embedding model, press **Test connection**, **Save**.
Default trainer PIN is **1234** – change it in Settings.

If Ollama is not running, Aria keeps working from her trained answers and says so honestly for anything else.

## Using it

| Tab | What it is for |
|---|---|
| **Reception** | Live conversation. Type, or press *Start voice reception* for hands-free mode. Each reply shows where it came from (trained / local AI / small talk) and its match score. |
| **Training** *(PIN)* | Add Q&A. Add alternate phrasings (or press *Suggest with AI*), *Polish* an answer for speaking, *Dictate* by voice. **Test matching** shows exactly what Aria would do with any phrase. |
| **Review** *(PIN)* | Unanswered / AI-answered visitor questions → edit → *Approve and train*. |
| **Knowledge** *(PIN)* | **Office Facts** (the only ground truth the AI may use), greetings, JSON import / export. |
| **Messages** *(PIN)* | Messages visitors left (name, contact, text). Mark handled. |
| **Settings** *(PIN)* | Model pickers, thresholds, voice, language, PIN. |

Starting voice reception locks the trainer tabs (kiosk mode).

**Tips for great training**
- Put short, true statements in *Office Facts* (hours, address, phone, policies). Mistral can then answer
  questions you never trained, without inventing anything.
- Give every Q&A 3–6 alternate phrasings. Speech-to-text hears things differently than you type.
- Check **Review** weekly; the most-asked questions are at the top.
- Voice answers are spoken: keep them to one or two sentences.

## Project layout

```
main.py                 entry point (GUI by default, --cli for terminal)
aria/
  config.py             persistent settings (data/settings.json), PIN hashing
  database.py           SQLite: Q&A, variants, facts, greetings, review queue, messages, log, embedding cache
  textutils.py          normalising, stemming, synonym folding, small-talk + yes/no intents
  retriever.py          hybrid TF-IDF / coverage / trigram / embedding scorer, auto-refreshing index
  llm.py                stdlib-only Ollama client: streaming chat, embeddings, model discovery
  brain.py              the decision pipeline, grounded prompt, message dialog, learning loop
  speech.py             queued, interruptible TTS (Windows SAPI) + adaptive-noise microphone STT
  session.py            hands-free reception loop (greet → listen → think → speak)
  context.py            wires everything together
  ui/                   minimal Tk interface, one file per tab
tests/                  engine, HTTP-client (mock Ollama), speaker and GUI smoke tests
data/                   created at first run: reception.db, settings.json
```

## What is smarter than the original

- **Local Mistral instead of Gemini/DialoGPT** – private, free, offline; streamed so Aria starts speaking
  after the first sentence instead of waiting for the full answer.
- **Grounded, not guessing** – the model may only use your Facts + trained answers; otherwise it emits an
  internal sentinel and Aria honestly offers to take a message. Hours or prices are never invented.
- **Matching that understands people** – synonyms (*washroom = restroom*), stemming, IDF weighting, typo
  tolerance and optional embeddings, over a main question *plus* alternate phrasings. Short follow-ups
  ("and on Sunday?") inherit the previous topic – but only when their own words support it.
- **Self-improving** – the Review queue turns real visitor questions into training data.
- **Takes messages** – a four-step voice dialog stored for the team.
- **Hands-free robustness** – microphone noise floor is calibrated every turn, the mic stays closed while
  Aria speaks (no self-transcription), Stop interrupts speech instantly, idle visitors are reset.
- **Duplicate guard** when training, enable/disable per entry, usage counters, backup/restore.

## Tests

```bash
python tests/test_engine.py      # retrieval, dialog, learning loop, database
python tests/test_llm_http.py    # real HTTP client vs. a mock Ollama server
python tests/test_speaker.py     # TTS queue / interruption semantics
xvfb-run -a python tests/ui_smoke.py [screenshot_dir]   # drives every tab (Linux: needs xvfb)
```

## Troubleshooting

| Symptom | Fix |
|---|---|
| Header says *Ollama offline* | Run `ollama serve`; check *Server URL* in Settings. |
| "Model 'x' not found" | `ollama list`, then choose the exact name in Settings → Chat model. |
| First answer is slow | The model is loading; Aria warms it up at startup and keeps it loaded for 30 min. |
| No voice input | `pip install sounddevice SpeechRecognition numpy`; check the microphone is the Windows default. Typing always works. |
| Wrong accent recognised | Settings → Speech language (e.g. `en-US`, `en-GB`, `hi-IN`). |
| No spoken replies | Windows only needs `pywin32`; tick *Speak replies*; adjust *Voice (name hints)*. |
