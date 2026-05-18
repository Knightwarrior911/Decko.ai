import json

import httpx

from app.config import Settings
from app.llm_client import LLMClient


class _Capture:
    def __init__(self, payload_text):
        self.payload_text = payload_text
        self.seen = {}

    def handler(self, request: httpx.Request) -> httpx.Response:
        self.seen["url"] = str(request.url)
        self.seen["headers"] = dict(request.headers)
        self.seen["body"] = json.loads(request.content)
        return httpx.Response(200, json=self._reply())

    def _reply(self):
        raise NotImplementedError


class _Anthropic(_Capture):
    def _reply(self):
        return {"content": [{"type": "text", "text": self.payload_text}]}


class _OpenAI(_Capture):
    def _reply(self):
        return {"choices": [{"message": {"content": self.payload_text}}]}


ACTIONS = '{"actions":[{"type":"apply_template","template":"title",' \
          '"content":{"title":"Hi","subtitle":"There"}}]}'


def _client(settings, cap):
    transport = httpx.MockTransport(cap.handler)
    return LLMClient(settings, api_key="sk-x", _transport=transport)


def test_anthropic_request_shape_and_parse():
    cap = _Anthropic("```json\n" + ACTIONS + "\n```")
    c = _client(Settings(provider="anthropic", model="claude-opus-4-7"), cap)
    out = c.call(snapshot="{SNAP}", user_request="make a title slide")
    assert cap.seen["url"] == "https://api.anthropic.com/v1/messages"
    assert cap.seen["headers"]["x-api-key"] == "sk-x"
    assert cap.seen["body"]["model"] == "claude-opus-4-7"
    sent = json.dumps(cap.seen["body"])
    assert "{SNAP}" in sent and "make a title slide" in sent
    assert json.loads(out)["actions"][0]["type"] == "apply_template"


def test_openai_request_shape_and_parse():
    cap = _OpenAI(ACTIONS)
    c = _client(Settings(provider="openai", model="gpt-4o"), cap)
    out = c.call(snapshot="{S}", user_request="x")
    assert cap.seen["url"] == "https://api.openai.com/v1/chat/completions"
    assert cap.seen["headers"]["authorization"] == "Bearer sk-x"
    assert cap.seen["body"]["model"] == "gpt-4o"
    assert json.loads(out)["actions"][0]["template"] == "title"


def test_generic_uses_base_url_openai_schema():
    cap = _OpenAI(ACTIONS)
    s = Settings(provider="generic", model="deepseek-chat",
                 base_url="https://api.deepseek.com/v1")
    c = _client(s, cap)
    c.call(snapshot="{S}", user_request="x")
    assert cap.seen["url"] == "https://api.deepseek.com/v1/chat/completions"
    assert cap.seen["headers"]["authorization"] == "Bearer sk-x"


import pytest


def test_sanitize_strips_leading_prose_and_fence():
    from app.llm_client import sanitize_actions
    raw = 'Sure! Here is the batch:\n```json\n' + ACTIONS + '\n```\nDone.'
    assert json.loads(sanitize_actions(raw))["actions"][0]["type"] == "apply_template"


def test_sanitize_rejects_no_object():
    from app.llm_client import sanitize_actions
    with pytest.raises(ValueError):
        sanitize_actions("no json here")


def test_sanitize_rejects_missing_actions():
    from app.llm_client import sanitize_actions
    with pytest.raises(ValueError):
        sanitize_actions('{"foo": 1}')


def test_http_error_raises():
    cap = _OpenAI(ACTIONS)
    def boom(request):
        return httpx.Response(500, json={"error": "bad"})
    s = Settings(provider="openai", model="gpt-4o")
    c = LLMClient(s, api_key="sk-x", _transport=httpx.MockTransport(boom))
    with pytest.raises(httpx.HTTPStatusError):
        c.call(snapshot="{}", user_request="x")


def test_build_fill_prompt_names_slots_and_brief():
    from app.llm_client import build_fill_prompt
    p = build_fill_prompt(["title", "subtitle"], "Q3 board review")
    assert "title" in p and "subtitle" in p
    assert "Q3 board review" in p
    assert "JSON" in p


def test_fill_with_ai_returns_slot_dict(monkeypatch):
    import app.main as M
    class FakeLLM:
        def __init__(self, *a, **k): pass
        def raw(self, prompt):
            return '{"title":"Q3","subtitle":"Review"}'
    monkeypatch.setattr(M, "LLMClient", FakeLLM)
    api = M.Api()
    from app.config import Settings
    api.settings = Settings(provider="openai", model="gpt-4o")
    monkeypatch.setattr(M.secrets, "get_api_key", lambda: "sk-x")
    monkeypatch.setattr(M.secrets, "has_api_key", lambda: True)
    r = api.fill_with_ai("title", "Q3 board review")
    assert r["content"] == {"title": "Q3", "subtitle": "Review"}
