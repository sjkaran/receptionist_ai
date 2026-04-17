"""
Office Reception Voice Assistant
==================================
A trainable voice assistant for office reception using:
- win32com + pythoncom for TTS (no silence issues)
- sounddevice + soundfile for audio capture (no PyAudio needed)
- SpeechRecognition (AudioData) for Google STT
- Google Gemini API for AI fallback
- SQLite for fast, efficient training data storage
"""

import sqlite3
import speech_recognition as sr
import sounddevice as sd
import numpy as np
import pythoncom
import win32com.client
import threading
from google import genai
import json
import os
import re
import time
from datetime import datetime
from difflib import SequenceMatcher

# ─────────────────────────────────────────────
# CONFIG — Edit these before running
# ─────────────────────────────────────────────
if os.path.exists("APIKEY.txt"):
    with open("APIKEY.txt","r") as src:
        GEMINI_API_KEY = src.read().strip()
else:
    GEMINI_API_KEY = os.gentenv("GEMINI_API_KEY","")



OFFICE_NAME    = "Karan's Office"              # Your office/company name
ASSISTANT_NAME = "Aria"                        # Assistant's name
DB_PATH        = "reception_data.db"           # SQLite DB file path
SIMILARITY_THRESHOLD = 0.55                    # How close a question must match (0.0–1.0)

# ─────────────────────────────────────────────
# TTS ENGINE (win32com — no silent-after-first-use bug)
# ─────────────────────────────────────────────
class VoiceEngine:
    def __init__(self):
        self._lock = threading.Lock()

    def speak(self, text: str):
        """Speak text in a fresh COM-initialized thread to avoid silence issues."""
        def _speak_thread(msg):
            pythoncom.CoInitialize()
            try:
                speaker = win32com.client.Dispatch("SAPI.SpVoice")
                # Prefer a female voice if available
                voices = speaker.GetVoices()
                for i in range(voices.Count):
                    voice = voices.Item(i)
                    desc = voice.GetDescription().lower()
                    if "zira" in desc or "female" in desc or "hazel" in desc:
                        speaker.Voice = voice
                        break
                speaker.Rate = 1   # -10 (slow) to 10 (fast)
                speaker.Volume = 100
                speaker.Speak(msg)
            finally:
                pythoncom.CoUninitialize()

        with self._lock:
            t = threading.Thread(target=_speak_thread, args=(text,), daemon=True)
            t.start()
            t.join()

    def speak_async(self, text: str):
        """Non-blocking speak."""
        t = threading.Thread(target=self.speak, args=(text,), daemon=True)
        t.start()

voice = VoiceEngine()


# ─────────────────────────────────────────────
# SPEECH RECOGNITION  (sounddevice — no PyAudio)
# ─────────────────────────────────────────────
SAMPLE_RATE  = 16000   # Hz — good balance of quality vs size for STT
CHANNELS     = 1       # Mono

recognizer = sr.Recognizer()
recognizer.pause_threshold  = 1.2   # seconds of silence before ending phrase
recognizer.energy_threshold = 300   # mic sensitivity (auto-adjusted below)

def _record_until_silence(timeout: int = 8, phrase_limit: int = 15) -> np.ndarray | None:
    """
    Record audio with sounddevice.
    Uses a simple energy-based VAD:
      - Waits up to `timeout` seconds for speech to start
      - Stops after `phrase_limit` seconds OR after 1.2 s of silence post-speech
    Returns numpy int16 array or None on timeout.
    """
    CHUNK       = 1024                         # frames per buffer
    SILENCE_RMS = 500                          # RMS below this = silence
    frames      = []
    speech_started  = False
    silence_chunks  = 0
    timeout_chunks  = int(SAMPLE_RATE / CHUNK * timeout)
    phrase_chunks   = int(SAMPLE_RATE / CHUNK * phrase_limit)
    silence_limit   = int(SAMPLE_RATE / CHUNK * 1.2)  # 1.2 s silence = done
    total_chunks    = 0

    def callback(indata, frame_count, time_info, status):
        nonlocal speech_started, silence_chunks, total_chunks
        frames.append(indata.copy())
        rms = np.sqrt(np.mean(indata.astype(np.float32) ** 2))
        if rms > SILENCE_RMS:
            speech_started = True
            silence_chunks = 0
        elif speech_started:
            silence_chunks += 1
        total_chunks += 1

    with sd.InputStream(samplerate=SAMPLE_RATE, channels=CHANNELS,
                        dtype="int16", blocksize=CHUNK, callback=callback):
        while True:
            time.sleep(0.01)
            if not speech_started and total_chunks > timeout_chunks:
                return None   # timed out waiting for speech
            if speech_started and silence_chunks >= silence_limit:
                break         # natural end of utterance
            if total_chunks > phrase_chunks:
                break         # hard limit reached

    if not speech_started or not frames:
        return None

    return np.concatenate(frames, axis=0)


def listen(prompt: str = "", timeout: int = 8, phrase_limit: int = 15) -> str:
    """
    Capture speech with sounddevice and transcribe via Google STT.
    Returns recognized text or empty string on failure.
    """
    if prompt:
        print(f"\n🎤  {prompt}")
        voice.speak(prompt)

    print("   [Listening...]")
    audio_np = _record_until_silence(timeout=timeout, phrase_limit=phrase_limit)

    if audio_np is None:
        print("   [Timeout — no speech detected]")
        return ""

    # Convert numpy int16 array → sr.AudioData (what Google STT expects)
    raw_bytes  = audio_np.tobytes()
    audio_data = sr.AudioData(raw_bytes, SAMPLE_RATE, 2)  # 2 bytes per int16 sample

    try:
        text = recognizer.recognize_google(audio_data)
        print(f"   [Heard]: {text}")
        return text.strip()
    except sr.UnknownValueError:
        print("   [Could not understand speech]")
        return ""
    except sr.RequestError as e:
        print(f"   [Speech API error: {e}]")
        return ""


# ─────────────────────────────────────────────
# SQLite DATABASE
# ─────────────────────────────────────────────
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # Greetings table
    c.execute("""
        CREATE TABLE IF NOT EXISTS greetings (
            id      INTEGER PRIMARY KEY AUTOINCREMENT,
            text    TEXT NOT NULL,
            created TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Q&A pairs table
    c.execute("""
        CREATE TABLE IF NOT EXISTS qa_pairs (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            question TEXT NOT NULL,
            answer   TEXT NOT NULL,
            hits     INTEGER DEFAULT 0,
            created  TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Create FTS (full-text search) virtual table for fast fuzzy matching
    c.execute("""
        CREATE VIRTUAL TABLE IF NOT EXISTS qa_fts
        USING fts5(question, answer, content='qa_pairs', content_rowid='id')
    """)

    # Trigger to keep FTS in sync
    c.execute("""
        CREATE TRIGGER IF NOT EXISTS qa_pairs_ai AFTER INSERT ON qa_pairs BEGIN
            INSERT INTO qa_fts(rowid, question, answer) VALUES (new.id, new.question, new.answer);
        END
    """)

    conn.commit()
    conn.close()


def db_connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def save_greeting(text: str):
    with db_connect() as conn:
        conn.execute("INSERT INTO greetings (text) VALUES (?)", (text,))


def get_greetings():
    with db_connect() as conn:
        rows = conn.execute("SELECT text FROM greetings ORDER BY id DESC").fetchall()
    return [r["text"] for r in rows]


def save_qa(question: str, answer: str):
    with db_connect() as conn:
        conn.execute("INSERT INTO qa_pairs (question, answer) VALUES (?, ?)", (question, answer))


def get_all_qa():
    with db_connect() as conn:
        rows = conn.execute("SELECT id, question, answer, hits FROM qa_pairs ORDER BY hits DESC").fetchall()
    return rows


def delete_qa(qa_id: int):
    with db_connect() as conn:
        conn.execute("DELETE FROM qa_pairs WHERE id = ?", (qa_id,))
        conn.execute("DELETE FROM qa_fts WHERE rowid = ?", (qa_id,))


def increment_hit(qa_id: int):
    with db_connect() as conn:
        conn.execute("UPDATE qa_pairs SET hits = hits + 1 WHERE id = ?", (qa_id,))


# ─────────────────────────────────────────────
# FUZZY MATCHING ENGINE
# ─────────────────────────────────────────────
def similarity(a: str, b: str) -> float:
    """Return similarity ratio between two strings (0.0–1.0)."""
    a = a.lower().strip()
    b = b.lower().strip()
    return SequenceMatcher(None, a, b).ratio()


def find_best_answer(user_question: str):
    """
    Search trained Q&A for best match.
    Returns (answer, qa_id, score) or (None, None, 0) if no match.
    Strategy:
    1. FTS keyword search for candidates
    2. Fuzzy similarity re-rank
    3. Return best if above threshold
    """
    qa_all = get_all_qa()
    if not qa_all:
        return None, None, 0.0

    best_score = 0.0
    best_answer = None
    best_id = None

    for row in qa_all:
        score = similarity(user_question, row["question"])
        # Boost score for keyword overlap
        uq_words = set(user_question.lower().split())
        tq_words = set(row["question"].lower().split())
        common = uq_words & tq_words
        if tq_words:
            keyword_boost = len(common) / len(tq_words) * 0.3
            score = min(1.0, score + keyword_boost)

        if score > best_score:
            best_score = score
            best_answer = row["answer"]
            best_id = row["id"]

    if best_score >= SIMILARITY_THRESHOLD:
        return best_answer, best_id, best_score
    return None, None, best_score


# ─────────────────────────────────────────────
# GEMINI AI FALLBACK
# ─────────────────────────────────────────────
_client = None

def get_gemini():
    global _client
    if _client is None:
        if not GEMINI_API_KEY:
            print("❌ Error: APIKEY.txt is missing or empty.")
            return None
        
        # Explicitly set vertexai=False to avoid 401 errors
        _client = genai.Client(api_key=GEMINI_API_KEY, vertexai=False)
    return _client


def ask_gemini(user_question: str, conversation_history: list) -> str:
    try:
        client = get_gemini()
        if not client: return "System error: AI not configured."

        system_context = f"""You are {ASSISTANT_NAME}, a professional and friendly voice receptionist at {OFFICE_NAME}.
Your job is to assist visitors and callers politely and helpfully.
Keep responses concise (2–3 sentences max) as they will be spoken aloud.
Do not use bullet points, markdown, or special characters.
Be warm, professional, and helpful."""
        
        # Format history for the new SDK
        history_text = ""
        for entry in conversation_history[-6:]:
            history_text += f"Visitor: {entry['user']}\nAria: {entry['assistant']}\n"

        full_prompt = f"{system_context}\n\nRecent context:\n{history_text}\nVisitor: {user_question}\n{ASSISTANT_NAME}:"
        
        # Use a 2026 stable model
        response = client.models.generate_content(
            model="gemini-2.0-flash", # Use 2.0 or 1.5-flash
            contents=full_prompt
        )
        return response.text.strip()
    except Exception as e:
        # This will now show the actual error if it still fails
        print(f"   [Gemini SDK error: {e}]")
        return "I'm having trouble connecting to the network. Please wait a moment."
# ─────────────────────────────────────────────
# TRAINING MODE
# ─────────────────────────────────────────────
def run_training_mode():
    print("\n" + "="*60)
    print("          TRAINING MODE")
    print("="*60)
    voice.speak("Welcome to training mode. I will guide you through the setup.")

    while True:
        print("\n--- Training Menu ---")
        print("  1. Train greeting message")
        print("  2. Add Q&A pair (question + answer)")
        print("  3. View all trained data")
        print("  4. Delete a Q&A pair")
        print("  5. Exit training mode")
        voice.speak("Please choose an option: 1 for greeting, 2 for question and answer pair, 3 to view data, 4 to delete, or 5 to exit.")

        choice = input("\nEnter choice (1-5): ").strip()

        if choice == "1":
            train_greeting()
        elif choice == "2":
            train_qa_pair()
        elif choice == "3":
            view_training_data()
        elif choice == "4":
            delete_training_entry()
        elif choice == "5":
            voice.speak("Exiting training mode. Goodbye trainer!")
            break
        else:
            print("Invalid choice. Please enter 1–5.")


def train_greeting():
    print("\n📝  TRAIN GREETING")
    print("You can provide the greeting by voice or typing.")
    method = input("Enter greeting by (v)oice or (t)ype? ").strip().lower()

    if method == "v":
        greeting = listen("Please say the greeting message for customers.")
    else:
        voice.speak("Please type the greeting message.")
        greeting = input("Type greeting: ").strip()

    if greeting:
        save_greeting(greeting)
        print(f"✅  Greeting saved: \"{greeting}\"")
        voice.speak(f"Greeting saved. It will say: {greeting}")
    else:
        print("❌  No greeting captured.")
        voice.speak("No greeting was captured. Please try again.")


def train_qa_pair():
    print("\n📝  TRAIN Q&A PAIR")
    voice.speak("Let's add a new question and answer pair.")
    method = input("Enter by (v)oice or (t)ype? ").strip().lower()

    if method == "v":
        question = listen("Please say the question a customer might ask.")
        if not question:
            voice.speak("No question captured. Returning to menu.")
            return
        answer = listen("Now say the answer I should give to that question.")
        if not answer:
            voice.speak("No answer captured. Returning to menu.")
            return
    else:
        voice.speak("Please type the question.")
        question = input("Question: ").strip()
        voice.speak("Please type the answer.")
        answer = input("Answer: ").strip()

    if question and answer:
        save_qa(question, answer)
        print(f"✅  Saved — Q: \"{question}\" → A: \"{answer}\"")
        voice.speak(f"Saved. If a customer asks: {question}. I will answer: {answer}")
    else:
        print("❌  Incomplete Q&A pair.")


def view_training_data():
    print("\n📋  TRAINED DATA")
    greetings = get_greetings()
    print(f"\n  Greetings ({len(greetings)}):")
    for i, g in enumerate(greetings, 1):
        print(f"    {i}. {g}")

    qa_list = get_all_qa()
    print(f"\n  Q&A Pairs ({len(qa_list)}):")
    for row in qa_list:
        print(f"    [{row['id']}] Q: {row['question']}")
        print(f"         A: {row['answer']}  (used {row['hits']} times)")

    voice.speak(f"You have {len(greetings)} greetings and {len(qa_list)} question-answer pairs trained.")


def delete_training_entry():
    qa_list = get_all_qa()
    if not qa_list:
        print("No Q&A pairs to delete.")
        voice.speak("There are no question and answer pairs to delete.")
        return

    view_training_data()
    try:
        qa_id = int(input("\nEnter the ID number to delete: ").strip())
        delete_qa(qa_id)
        print(f"✅  Deleted Q&A entry #{qa_id}")
        voice.speak(f"Entry {qa_id} has been deleted.")
    except ValueError:
        print("Invalid ID.")


# ─────────────────────────────────────────────
# RECEPTION MODE
# ─────────────────────────────────────────────
def run_reception_mode():
    print("\n" + "="*60)
    print("          RECEPTION MODE  —  Active")
    print("="*60)
    print("Press Ctrl+C to stop.\n")

    greetings = get_greetings()
    if greetings:
        greeting_msg = greetings[0]  # Most recently trained greeting
    else:
        greeting_msg = f"Welcome to {OFFICE_NAME}! I am {ASSISTANT_NAME}, your virtual receptionist. How may I help you today?"

    conversation_history = []

    voice.speak(greeting_msg)
    print(f"\n🤖  {ASSISTANT_NAME}: {greeting_msg}")

    no_input_count = 0
    MAX_SILENCE = 3  # Silently re-prompt after 3 misses

    while True:
        user_input = listen(timeout=10, phrase_limit=20)

        if not user_input:
            no_input_count += 1
            if no_input_count >= MAX_SILENCE:
                nudge = "I am here to help you. Please feel free to ask me anything."
                voice.speak(nudge)
                print(f"\n🤖  {ASSISTANT_NAME}: {nudge}")
                no_input_count = 0
            continue

        no_input_count = 0
        print(f"\n👤  Visitor: {user_input}")

        # Check for exit keywords (for demo/testing)
        if any(w in user_input.lower() for w in ["exit reception", "quit assistant", "stop assistant"]):
            farewell = "Thank you for visiting. Have a wonderful day!"
            voice.speak(farewell)
            print(f"\n🤖  {ASSISTANT_NAME}: {farewell}")
            break

        # Step 1: Search trained Q&A
        answer, qa_id, score = find_best_answer(user_input)

        if answer:
            print(f"   [Matched trained Q&A — score: {score:.2f}]")
            increment_hit(qa_id)
            response = answer
        else:
            # Step 2: Fallback to Gemini
            print(f"   [No match (best: {score:.2f}) — using Gemini AI]")
            response = ask_gemini(user_input, conversation_history)

        print(f"\n🤖  {ASSISTANT_NAME}: {response}")
        voice.speak(response)

        # Save to conversation history
        conversation_history.append({"user": user_input, "assistant": response})
        if len(conversation_history) > 20:
            conversation_history = conversation_history[-20:]


# ─────────────────────────────────────────────
# MAIN ENTRY POINT
# ─────────────────────────────────────────────
def main():
    init_db()

    print("\n" + "="*60)
    print(f"   🏢  {OFFICE_NAME}  |  Reception Assistant: {ASSISTANT_NAME}")
    print("="*60)
    print("\nSelect mode:")
    print("  1. Reception Mode  (for customers)")
    print("  2. Training Mode   (for trainer/admin)")
    print("  3. Exit")

    choice = input("\nEnter choice (1/2/3): ").strip()

    if choice == "1":
        run_reception_mode()
    elif choice == "2":
        # Simple PIN protection for training mode
        pin = input("Enter trainer PIN (default: 1234): ").strip()
        if pin == "1234":
            run_training_mode()
        else:
            print("❌  Incorrect PIN. Access denied.")
            voice.speak("Incorrect PIN. Access denied.")
    elif choice == "3":
        print("Goodbye!")
    else:
        print("Invalid choice.")


if __name__ == "__main__":
    main()
