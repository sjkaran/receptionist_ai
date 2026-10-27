"""Tiny fake Ollama server (stdlib) so the real HTTP client and the UI can be tested offline."""
from __future__ import annotations

import hashlib
import json
import math
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

MODELS = ["llama3:8b", "mistral:7b-instruct", "nomic-embed-text:latest"]


def fake_embed(text: str, dims: int = 64) -> list[float]:
    v = [0.0] * dims
    for w in text.lower().split():
        v[int(hashlib.md5(w.encode()).hexdigest(), 16) % dims] += 1.0
    n = math.sqrt(sum(x * x for x in v)) or 1.0
    return [x / n for x in v]


def scripted_reply(messages: list[dict]) -> str:
    user = messages[-1]["content"].lower()
    if "insurance" in user:
        return "[[UNKNOWN]]"
    if "phone" in user or "contact" in user:
        return "Our phone number is 12345. I am happy to help with anything else."
    if "variants" in user:
        return "when do you open\nwhat time do you close\n- are you open today"
    return "Certainly, how can I help further?"


class Handler(BaseHTTPRequestHandler):
    log = []

    def log_message(self, *a):  # silence
        pass

    def _json(self, obj, code=200):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/api/tags":
            self._json({"models": [{"name": m} for m in MODELS]})
        else:
            self._json({"error": "not found"}, 404)

    def do_POST(self):
        n = int(self.headers.get("Content-Length", 0))
        req = json.loads(self.rfile.read(n) or b"{}")
        Handler.log.append((self.path, req))
        if self.path == "/api/chat":
            if req.get("model") not in MODELS:
                return self._json({"error": f"model '{req.get('model')}' not found"}, 404)
            text = scripted_reply(req["messages"])
            self.send_response(200)
            self.send_header("Content-Type", "application/x-ndjson")
            self.end_headers()
            for i in range(0, len(text), 4):
                line = json.dumps({"message": {"role": "assistant", "content": text[i:i + 4]}, "done": False})
                self.wfile.write((line + "\n").encode())
                self.wfile.flush()
            self.wfile.write((json.dumps({"message": {"content": ""}, "done": True}) + "\n").encode())
        elif self.path == "/api/embed":
            if req.get("model") not in MODELS and f"{req.get('model')}:latest" not in MODELS:
                return self._json({"error": f"model \"{req.get('model')}\" not found, try pulling it first"}, 404)
            inputs = req["input"] if isinstance(req["input"], list) else [req["input"]]
            self._json({"embeddings": [fake_embed(t) for t in inputs]})
        else:
            self._json({"error": "not found"}, 404)


def start(port: int = 0):
    srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv, f"http://127.0.0.1:{srv.server_address[1]}"
