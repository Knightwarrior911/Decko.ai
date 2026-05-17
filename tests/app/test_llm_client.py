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
