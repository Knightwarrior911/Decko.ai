"""pywebview window + js_api bridge wiring the chat + side panel to the
Python core. UI is NOT in the deterministic gate (spec §8). `--selfcheck`
is retained for packaging_smoke (proves the bundled carrier resolves)."""
import concurrent.futures
import os
import sys
from datetime import datetime

import webview

from app import secrets
from app.carrier import _bundled_carrier, ensure_carrier
from app.config import (DB_PATH, Settings, ensure_app_dirs, load_persisted,
                        save_persisted, settings_from_persisted)
from app.deck_controller import (DeckController, EmptyDeckError,
                                 NoOpenDeckError, NoPowerPointError)
from app.llm_client import LLMClient
from app.orchestrator import ChatOrchestrator
from app.store import Store


def parse_captured_registry(path: str) -> list:
    import json
    import os
    if not path or not os.path.exists(path):
        return []
    try:
        data = json.loads(open(path, "r", encoding="utf-8").read())
    except Exception:
        return []
    tpls = (data or {}).get("templates", {})
    return [{"name": n, "slots": sorted(list(v.keys()))}
            for n, v in tpls.items()] if isinstance(tpls, dict) else []


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
        self.settings = settings_from_persisted()
        self.store = Store(DB_PATH)
        self.dc = None
        self.orch = None
        self.session_id = None
        # pywebview dispatches each js_api call on a different thread.
        # PowerPoint COM objects are apartment-bound — create AND use
        # them on ONE thread (single STA worker).
        self._com_pool = concurrent.futures.ThreadPoolExecutor(
            max_workers=1, thread_name_prefix="decko-com")

    def _com(self, fn):
        return self._com_pool.submit(fn).result()

    def boot(self):
        ensure_app_dirs()
        self.store.init()
        p = load_persisted()
        return {
            "has_key": secrets.has_api_key(),
            "settings": {"provider": self.settings.provider,
                         "model": self.settings.model,
                         "base_url": self.settings.base_url},
            "last_deck_path": p.get("last_deck_path", ""),
            "last_mode": p.get("last_mode", "attach"),
            "sessions": self.store.list_sessions(),
        }

    def list_builtin_templates(self):
        from app.template_slots import BUILTIN_SLOTS
        return {"templates": [{"name": n, "slots": s}
                              for n, s in BUILTIN_SLOTS.items()]}

    def list_captured_templates(self):
        if self.dc is None:
            return {"templates": []}
        try:
            path = self._com(self.dc.captured_registry_path)
        except Exception:
            return {"templates": []}
        return {"templates": parse_captured_registry(path)}

    def save_settings(self, provider, model, base_url, api_key):
        self.settings = Settings(provider=provider, model=model,
                                 base_url=base_url or "")
        self.settings.validate()
        if api_key:
            secrets.set_api_key(api_key)
        pr = load_persisted()
        save_persisted(self.settings,
                       pr.get("last_deck_path", ""),
                       pr.get("last_mode", "attach"))
        return {"ok": True}

    def new_session(self):
        # Drop the current chat context; the next Start session opens a
        # fresh session. Does not close PowerPoint.
        self.orch = None
        self.session_id = None
        return {"ok": True}

    def open_session(self, mode, file_path=""):
        def _start():
            dc = DeckController()
            dc.start()
            if mode == "attach":
                dc.attach_open_deck()
            else:
                dc.open_file(file_path)
            return dc
        try:
            self.dc = self._com(_start)
        except NoPowerPointError:
            return {"error": "Microsoft PowerPoint is required and was "
                             "not found. Install PowerPoint and retry."}
        except NoOpenDeckError:
            return {"error": "No deck open in PowerPoint. Open one, or "
                             "choose 'Open file' instead."}
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            self.orch = None
            return {"error": f"Could not start session: "
                             f"{type(e).__name__}: {e}"}
        if not secrets.has_api_key():
            self.orch = None
            return {"error": "Save your LLM API key in the side panel "
                             "first, then start a session."}
        # Remember the deck for next launch.
        save_persisted(self.settings, file_path, mode)
        # New chat session.
        label = (os.path.basename(file_path) if mode == "file" and file_path
                 else "Open deck")
        title = f"{label} — {datetime.now().strftime('%b %d %H:%M')}"
        self.session_id = self.store.create_session(title)
        llm = LLMClient(self.settings, secrets.get_api_key() or "")
        self.orch = ChatOrchestrator(self.dc, llm, self.store,
                                     session_id=self.session_id)
        return {"ok": True, "session_id": self.session_id, "title": title}

    def send(self, text):
        if self.orch is None:
            return {"error": "Start a session first."}
        if not secrets.has_api_key():
            return {"error": "Save your LLM API key in the side panel "
                             "first."}
        # Rebuild the LLM from CURRENT settings every turn so saving
        # settings after Start session takes effect (order-independent).
        self.orch.llm = LLMClient(self.settings,
                                  secrets.get_api_key() or "")
        try:
            return self._com(lambda: self.orch.run(text))
        except EmptyDeckError as e:
            return {"error": str(e)}
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            try:
                import httpx
                if isinstance(e, httpx.HTTPStatusError):
                    body = ""
                    try:
                        body = e.response.text[:800]
                    except Exception:
                        pass
                    return {"error": f"LLM API error "
                                     f"{e.response.status_code}: {body}"}
            except Exception:
                pass
            return {"error": f"{type(e).__name__}: {e}"}

    def _log_turn(self, request: str, summary: str, actions: dict):
        import json
        warnings = 0
        try:
            warnings = self.orch._warn_count(summary) if self.orch else 0
        except Exception:
            warnings = 0
        if self.session_id is not None:
            self.store.add_turn(request=request,
                                actions_json=json.dumps(actions),
                                result_summary=summary or "",
                                warnings=warnings,
                                session_id=self.session_id)

    def _require_session(self):
        if self.orch is None or self.dc is None:
            return {"error": "Start a session first."}
        return None

    def apply_template(self, template, content, target):
        g = self._require_session()
        if g:
            return g
        act = {"type": "apply_template", "template": template,
               "content": content}
        tgt = "append"
        if target and target.get("mode") == "replace":
            slide_raw = target.get("slide")
            if slide_raw is None or str(slide_raw).strip() == "":
                return {"error": "Replace mode requires a slide number."}
            try:
                act["slide"] = int(slide_raw)
            except (TypeError, ValueError):
                return {"error": f"Invalid slide number: {slide_raw!r}"}
            tgt = f"replace slide {act['slide']}"
        try:
            summary = self._com(lambda: self.dc.run_action(act))
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        self._log_turn(f"Apply template: {template} ({tgt})",
                       summary, {"actions": [act]})
        return {"ok": True, "summary": summary}

    def capture_template(self, name):
        g = self._require_session()
        if g:
            return g
        if not str(name).strip():
            return {"error": "Enter a template name."}
        act = {"type": "capture_template", "name": str(name).strip()}
        try:
            summary = self._com(lambda: self.dc.run_action(act))
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        self._log_turn(f"Capture template: {name}", summary,
                       {"actions": [act]})
        return {"ok": True, "summary": summary}

    def rename_template(self, from_name, to_name):
        g = self._require_session()
        if g:
            return g
        if not str(from_name).strip() or not str(to_name).strip():
            return {"error": "Both names are required."}
        act = {"type": "rename_template",
               "from": str(from_name).strip(),
               "to": str(to_name).strip()}
        try:
            summary = self._com(lambda: self.dc.run_action(act))
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        self._log_turn(f"Rename template: {from_name} -> {to_name}",
                       summary, {"actions": [act]})
        return {"ok": True, "summary": summary}

    def delete_template(self, name):
        g = self._require_session()
        if g:
            return g
        act = {"type": "delete_template", "name": str(name).strip()}
        try:
            summary = self._com(lambda: self.dc.run_action(act))
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        self._log_turn(f"Delete template: {name}", summary,
                       {"actions": [act]})
        return {"ok": True, "summary": summary}

    def list_sessions(self):
        return {"sessions": self.store.list_sessions()}

    def load_session(self, session_id):
        return {"turns": self.store.turns_for_session(session_id)}

    def save_powerpoint(self):
        if self.dc is None:
            return {"error": "No session — nothing to save."}
        try:
            self._com(self.dc.save_deck_now)
            return {"ok": True}
        except Exception as e:  # noqa: BLE001
            return {"error": f"Save failed: {type(e).__name__}: {e}"}

    def shutdown(self):
        try:
            if self.dc is not None:
                self._com(lambda: self.dc.close(save_deck=False))
        finally:
            self._com_pool.shutdown(wait=False)
        return {"ok": True}


def _web_index() -> str:
    # Absolute path so pywebview resolves it in dev AND under PyInstaller
    # (decko.spec bundles app/web -> sys._MEIPASS/app/web).
    from pathlib import Path
    base = Path(getattr(sys, "_MEIPASS",
                        Path(__file__).resolve().parent.parent))
    return str(base / "app" / "web" / "index.html")


def main() -> int:
    if "--selfcheck" in sys.argv[1:]:
        return _selfcheck()
    api = Api()
    webview.create_window("Decko", _web_index(), js_api=api,
                          width=1180, height=760, text_select=True)
    webview.start()
    api.shutdown()
    return 0


if __name__ == "__main__":
    sys.exit(main())
