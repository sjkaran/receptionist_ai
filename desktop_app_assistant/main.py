"""Aria — trainable AI receptionist (local Ollama / Mistral).

    python main.py          # desktop app (recommended)
    python main.py --cli    # terminal chat, handy for quick testing
"""
from __future__ import annotations

import argparse
import sys
import threading

from aria.context import Context


def run_cli(ctx: Context) -> None:
    s = ctx.settings
    print(f"\n{s.assistant_name} @ {s.office_name}  (Ctrl+C to quit)")
    try:
        print(f"LLM: {ctx.llm.resolve_model()}")
    except Exception as exc:
        print(f"LLM: offline ({exc}) — trained answers only")
    greeting = ctx.brain.greeting()
    print(f"\n{s.assistant_name}: {greeting}")
    ctx.speaker.say(greeting)
    while True:
        try:
            text = input("\nYou: ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            break
        if not text:
            continue
        print(f"{s.assistant_name}: ", end="", flush=True)

        def say(sentence: str) -> None:
            print(sentence + " ", end="", flush=True)
            ctx.speaker.say(sentence)

        reply = ctx.brain.respond(text, on_sentence=say)
        print(f"\n   [{reply.source} · {reply.score:.2f}{' · queued for review' if reply.queued else ''}]")
    ctx.speaker.stop()


def run_gui(ctx: Context) -> None:
    from aria.ui.app import App
    try:                                         # sharp text on Windows high-DPI displays
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = App(ctx)
    threading.Thread(target=ctx.llm.warmup, daemon=True).start()
    app.mainloop()


def main() -> int:
    ap = argparse.ArgumentParser(description="Aria AI receptionist")
    ap.add_argument("--cli", action="store_true", help="terminal chat instead of the desktop UI")
    args = ap.parse_args()
    ctx = Context.create()
    (run_cli if args.cli else run_gui)(ctx)
    return 0


if __name__ == "__main__":
    sys.exit(main())
