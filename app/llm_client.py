"""LLM brain. BYO key. Three provider modes (spec D6):
  anthropic -> https://api.anthropic.com/v1/messages
  openai    -> https://api.openai.com/v1/chat/completions
  generic   -> {base_url}/chat/completions  (OpenAI schema)
Output is sanitized to a clean {"actions":[...]} JSON string."""
import json
import re

import httpx

from app.config import Settings

_SYSTEM = (
    "You are Decko's action generator. Given a PowerPoint deck snapshot "
    "and a user request, return ONLY a JSON object {\"actions\":[...]}. "
    "No prose, no markdown fences."
)


def _build_user_prompt(snapshot: str, user_request: str) -> str:
    return (
        "DECK SNAPSHOT:\n" + snapshot + "\n\n"
        "USER REQUEST:\n" + user_request + "\n\n"
        'Return only {"actions":[...]}.'
    )


def sanitize_actions(raw: str) -> str:
    """Strip fences/prose; return the outermost {...} containing "actions"."""
    text = raw.strip()
    text = re.sub(r"```[a-zA-Z]*\n?", "", text)
    text = text.replace("```", "").strip()
    start = text.find("{")
    if start == -1:
        raise ValueError("no JSON object in LLM output")
    try:
        obj, _ = json.JSONDecoder().raw_decode(text[start:])
    except json.JSONDecodeError as e:
        raise ValueError(f"unparseable JSON in LLM output: {e}") from e
    if "actions" not in obj or not isinstance(obj["actions"], list):
        raise ValueError('LLM output missing "actions" array')
    return json.dumps(obj)


class LLMClient:
    def __init__(self, settings: Settings, api_key: str,
                 _transport: httpx.BaseTransport | None = None):
        settings.validate()
        self.s = settings
        self.api_key = api_key
        self._transport = _transport

    def _http(self) -> httpx.Client:
        return httpx.Client(timeout=120.0, transport=self._transport)

    def call(self, snapshot: str, user_request: str) -> str:
        user = _build_user_prompt(snapshot, user_request)
        if self.s.provider == "anthropic":
            url = "https://api.anthropic.com/v1/messages"
            headers = {"x-api-key": self.api_key,
                       "anthropic-version": "2023-06-01"}
            body = {"model": self.s.model, "max_tokens": 8000,
                    "system": _SYSTEM,
                    "messages": [{"role": "user", "content": user}]}
            with self._http() as h:
                r = h.post(url, json=body, headers=headers)
                r.raise_for_status()
                raw = r.json()["content"][0]["text"]
        else:
            base = ("https://api.openai.com/v1" if self.s.provider == "openai"
                    else self.s.base_url.rstrip("/"))
            url = base + "/chat/completions"
            headers = {"authorization": f"Bearer {self.api_key}"}
            body = {"model": self.s.model,
                    "messages": [{"role": "system", "content": _SYSTEM},
                                 {"role": "user", "content": user}]}
            with self._http() as h:
                r = h.post(url, json=body, headers=headers)
                r.raise_for_status()
                raw = r.json()["choices"][0]["message"]["content"]
        return sanitize_actions(raw)
