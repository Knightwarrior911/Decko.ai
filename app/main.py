"""pywebview window + js_api bridge wiring the chat + side panel to the
Python core. UI is NOT in the deterministic gate (spec §8). `--selfcheck`
is retained for packaging_smoke (proves the bundled carrier resolves)."""
import sys

import webview

from app import secrets
from app.carrier import _bundled_carrier, ensure_carrier
from app.config import DB_PATH, Settings, ensure_app_dirs
from app.deck_controller import (DeckController, EmptyDeckError,
                                 NoOpenDeckError, NoPowerPointError)
from app.llm_client import LLMClient
from app.orchestrator import ChatOrchestrator
from app.store import Store


def _selfcheck() -> int:
    ensure_app_dirs()
    src = _bundled_carrier()
    if not src.exists():
        print(f"SELFCHECK FAIL: bundled carrier missing at {src}")
        return 1
    dest = ensure_carrier()
    if not dest.exists():
        print(f"SELFCHECK FAIL: carrier not installed at {dest}")
        return 1
    print(f"SELFCHECK OK: bundled={src} installed={dest}")
    return 0


class Api:
    def __init__(self):
        self.settings = Settings()
        self.store = Store(DB_PATH)
        self.dc = None
        self.orch = None

    def boot(self):
        ensure_app_dirs()
        self.store.init()
        return {"has_key": secrets.has_api_key(),
                "history": self.store.list_turns()}

    def save_settings(self, provider, model, base_url, api_key):
        self.settings = Settings(provider=provider, model=model,
                                 base_url=base_url or "")
        self.settings.validate()
        if api_key:
            secrets.set_api_key(api_key)
        return {"ok": True}

    def open_session(self, mode, file_path=""):
        self.dc = DeckController()
        try:
            self.dc.start()
            if mode == "attach":
                self.dc.attach_open_deck()
            else:
                self.dc.open_file(file_path)
        except NoPowerPointError:
            return {"error": "Microsoft PowerPoint is required and was "
                             "not found. Install PowerPoint and retry."}
        except NoOpenDeckError:
            return {"error": "No deck open in PowerPoint. Open one, or "
                             "choose 'Open file' instead."}
        if not secrets.has_api_key():
            self.orch = None
            return {"error": "Save your LLM API key in the side panel "
                             "first, then start a session."}
        llm = LLMClient(self.settings, secrets.get_api_key() or "")
        self.orch = ChatOrchestrator(self.dc, llm, self.store)
        return {"ok": True}

    def send(self, text):
        if self.orch is None:
            return {"error": "Start a session first."}
        try:
            return self.orch.run(text)
        except EmptyDeckError as e:
            return {"error": str(e)}
        except Exception as e:  # noqa: BLE001
            return {"error": str(e)}

    def shutdown(self):
        if self.dc is not None:
            self.dc.close(save_deck=False)
        return {"ok": True}


def main() -> int:
    if "--selfcheck" in sys.argv[1:]:
        return _selfcheck()
    api = Api()
    webview.create_window("Decko", "app/web/index.html",
                          js_api=api, width=1100, height=720)
    webview.start()
    api.shutdown()
    return 0


if __name__ == "__main__":
    sys.exit(main())
