"""Zero-dependency Ollama client (stdlib only): chat streaming, embeddings, model discovery."""
from __future__ import annotations

import json
import threading
import urllib.error
import urllib.request
from typing import Iterator


class OllamaError(RuntimeError):
    pass


class ModelNotFound(OllamaError):
    pass


class OllamaClient:
    def __init__(self, settings):
        self.s = settings
        self._resolved: str | None = None

    # ------------------------------------------------------------- transport
    def _request(self, path: str, payload: dict | None = None, timeout: float | None = None):
        url = self.s.ollama_url.rstrip("/") + path
        data = json.dumps(payload).encode() if payload is not None else None
        req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
        try:
            return urllib.request.urlopen(req, timeout=timeout or self.s.llm_timeout)
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "ignore")
            try:
                body = json.loads(body).get("error", body)
            except ValueError:
                pass
            if e.code == 404 and "model" in body.lower():
                raise ModelNotFound(f"{body}. Pull it with:  ollama pull <model>") from None
            raise OllamaError(f"Ollama HTTP {e.code}: {body}") from None
        except (urllib.error.URLError, OSError) as e:
            raise OllamaError(
                f"Cannot reach Ollama at {self.s.ollama_url} ({getattr(e, 'reason', e)}). "
                "Start it with 'ollama serve'.") from None

    # -------------------------------------------------------------- discovery
    def models(self) -> list[str]:
        with self._request("/api/tags", timeout=4) as r:
            data = json.load(r)
        return [m["name"] for m in data.get("models", [])]

    def is_up(self) -> bool:
        try:
            self.models()
            return True
        except OllamaError:
            return False

    def resolve_model(self, force: bool = False) -> str:
        if self._resolved and not force:
            return self._resolved
        available = self.models()
        if not available:
            raise OllamaError("Ollama is running but has no models. Run:  ollama pull mistral")
        want = (self.s.ollama_model or "").strip()
        chosen = None
        if want:
            chosen = next((m for m in available if m == want), None) \
                or next((m for m in available if m.split(":")[0] == want.split(":")[0]), None)
            if not chosen:
                raise OllamaError(f"Model '{want}' not found. Installed: {', '.join(available)}")
        else:
            chosen = next((m for m in available if "mistral" in m.lower()), None) or available[0]
        self._resolved = chosen
        return chosen

    def reset(self) -> None:
        self._resolved = None

    # ------------------------------------------------------------------- chat
    def chat_stream(self, messages: list[dict], cancel: threading.Event | None = None,
                    num_predict: int | None = None, temperature: float | None = None) -> Iterator[str]:
        payload = {
            "model": self.resolve_model(),
            "messages": messages,
            "stream": True,
            "keep_alive": "30m",
            "options": {
                "temperature": self.s.temperature if temperature is None else temperature,
                "num_predict": num_predict or self.s.num_predict,
                "top_p": 0.9,
                "repeat_penalty": 1.1,
            },
        }
        resp = self._request("/api/chat", payload)
        try:
            for raw in resp:
                if cancel is not None and cancel.is_set():
                    break
                raw = raw.strip()
                if not raw:
                    continue
                try:
                    obj = json.loads(raw)
                except ValueError:
                    continue
                if obj.get("error"):
                    raise OllamaError(obj["error"])
                chunk = (obj.get("message") or {}).get("content", "")
                if chunk:
                    yield chunk
                if obj.get("done"):
                    break
        finally:
            resp.close()

    def chat(self, messages: list[dict], **kw) -> str:
        return "".join(self.chat_stream(messages, **kw)).strip()

    def complete(self, system: str, user: str, **kw) -> str:
        return self.chat([{"role": "system", "content": system}, {"role": "user", "content": user}], **kw)

    def warmup(self) -> None:
        """Load the model into memory so the first visitor does not wait."""
        try:
            self.chat([{"role": "user", "content": "Hi"}], num_predict=1)
        except OllamaError as e:
            print(f"[llm] warm-up skipped: {e}")

    # ------------------------------------------------------------- embeddings
    def embed(self, texts: list[str]) -> list[list[float]]:
        model = (self.s.embed_model or "").strip()
        if not model:
            raise OllamaError("No embedding model configured.")
        try:
            with self._request("/api/embed", {"model": model, "input": texts}) as r:
                data = json.load(r)
            vecs = data.get("embeddings")
            if vecs and len(vecs) == len(texts):
                return vecs
        except ModelNotFound:
            raise
        except OllamaError as first:                # endpoint missing on older Ollama -> legacy API below
            if "404" not in str(first):
                raise
        out = []                                    # older Ollama: one call per text
        for t in texts:
            with self._request("/api/embeddings", {"model": model, "prompt": t}) as r:
                out.append(json.load(r)["embedding"])
        return out
