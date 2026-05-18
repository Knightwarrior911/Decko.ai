"""Deterministic LLM stub. No network. Returns a fixed valid batch so
the core loop is gated without a live model (spec §8)."""

FIXED_ACTIONS = (
    '{"actions":[{"type":"apply_template","template":"title",'
    '"content":{"title":"Stub Title","subtitle":"Stub Sub"}}]}'
)


class StubLLM:
    def call(self, snapshot: str, user_request: str,
             prompt_template: str | None = None) -> str:
        return FIXED_ACTIONS
