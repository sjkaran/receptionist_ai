"""Local LLM fallback through Ollama's HTTP API (default model: mistral)."""
import re
import time
import requests


class LLMError(RuntimeError):
    pass


class OllamaClient:
    def __init__(self, settings):
        self.s = settings
        self._cache = (0.0, None)

    @property
    def base(self) -> str:
        return str(self.s["ollama_url"]).rstrip("/")

    def status(self) -> dict:
        """Cached for 5s so UI polling doesn't hammer Ollama."""
        ts, val = self._cache
        if val is None or time.time() - ts > 5:
            val = self._check()
            self._cache = (time.time(), val)
        return val

    def _check(self) -> dict:
        model = str(self.s["ollama_model"])
        try:
            r = requests.get(f"{self.base}/api/tags", timeout=3)
            r.raise_for_status()
            names = [m.get("name", "") for m in r.json().get("models", [])]
        except requests.RequestException:
            return {"online": False, "model": model,
                    "message": "Ollama is not running. Start it, then refresh."}
        if not any(n == model or n.split(":")[0] == model.split(":")[0] for n in names):
            return {"online": False, "model": model,
                    "message": f"Model '{model}' not installed. Run: ollama pull {model}"}
        return {"online": True, "model": model, "message": "Ready"}

    def _system_prompt(self) -> str:
        info = str(self.s["office_info"]).strip()
        prompt = (
            f"You are {self.s['assistant_name']}, the voice receptionist at {self.s['office_name']}. "
            "Help visitors politely and warmly. Your reply is spoken aloud, so answer in one to three "
            "short sentences, with no lists, markdown, or emojis. "
            "If you do not know something specific about the office, say so and suggest speaking to a staff member. "
            "Never invent opening hours, prices, names, or locations."
        )
        if info:
            prompt += f"\n\nFacts about the office you may use:\n{info}"
        return prompt

    def ask(self, question: str, history: list) -> str:
        messages = [{"role": "system", "content": self._system_prompt()}]
        for turn in history[-6:]:
            messages.append({"role": "user", "content": turn["user"]})
            messages.append({"role": "assistant", "content": turn["assistant"]})
        messages.append({"role": "user", "content": question})

        try:
            r = requests.post(
                f"{self.base}/api/chat",
                json={"model": self.s["ollama_model"], "messages": messages, "stream": False,
                      "options": {"temperature": 0.4, "num_predict": 120}},
                timeout=int(self.s["ollama_timeout"]),
            )
            r.raise_for_status()
            text = r.json()["message"]["content"]
        except requests.RequestException as e:
            raise LLMError(f"Ollama request failed: {e}")
        except (KeyError, ValueError):
            raise LLMError("Unexpected response from Ollama")

        text = re.sub(r"[*#_`>]+", "", text)
        text = re.sub(r"\s+", " ", text).strip()
        if not text:
            raise LLMError("Empty response from Ollama")
        return text
