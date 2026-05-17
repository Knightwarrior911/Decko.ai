"""One chat turn: text -> snapshot -> LLM -> actions -> engine ->
summary -> persist. `deck` and `llm` are duck-typed so this is unit
testable without COM (spec §6)."""
import re


class ChatOrchestrator:
    def __init__(self, deck, llm, store):
        self.deck = deck
        self.llm = llm
        self.store = store

    @staticmethod
    def _warn_count(summary: str) -> int:
        m = re.search(r"(\d+)\s+warning", summary)
        return int(m.group(1)) if m else 0

    def run(self, text: str) -> dict:
        snapshot = self.deck.get_snapshot()
        actions = self.llm.call(snapshot=snapshot, user_request=text)
        summary = self.deck.run_actions(actions)
        warnings = self._warn_count(summary)
        tid = self.store.add_turn(request=text, actions_json=actions,
                                  result_summary=summary, warnings=warnings)
        return {"id": tid, "summary": summary, "warnings": warnings,
                "actions_json": actions}
