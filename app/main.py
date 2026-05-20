"""pywebview window + js_api bridge wiring the chat + side panel to the
Python core. UI is NOT in the deterministic gate (spec §8). `--selfcheck`
is retained for packaging_smoke (proves the bundled carrier resolves)."""
import concurrent.futures
import os
import sys
import time
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
                         "base_url": self.settings.base_url,
                         "dock_mode": self.settings.dock_mode,
                         "decko_on_top": self.settings.decko_on_top,
                         "resize_ppt_for_dock":
                             self.settings.resize_ppt_for_dock},
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

    def save_settings(self, provider, model, base_url, api_key,
                       decko_on_top=None, resize_ppt_for_dock=None):
        # SP7: optional kwargs let the Settings dialog persist the two
        # new dock toggles in the same round-trip as the legacy fields.
        # Existing front-end callers pass 4 positional args — both new
        # kwargs default to "keep current".
        prev_on_top = self.settings.decko_on_top
        self.settings = Settings(
            provider=provider, model=model, base_url=base_url or "",
            dock_mode=self.settings.dock_mode,
            decko_on_top=(prev_on_top if decko_on_top is None
                          else bool(decko_on_top)),
            resize_ppt_for_dock=(self.settings.resize_ppt_for_dock
                                 if resize_ppt_for_dock is None
                                 else bool(resize_ppt_for_dock)))
        self.settings.validate()
        if api_key:
            secrets.set_api_key(api_key)
        pr = load_persisted()
        save_persisted(self.settings,
                       pr.get("last_deck_path", ""),
                       pr.get("last_mode", "attach"))
        # Apply on-top live.
        try:
            _apply_on_top(_decko_hwnd(), bool(self.settings.decko_on_top))
        except Exception:  # noqa: BLE001
            pass
        # If reflow toggled off, restore any cached PPT rect immediately.
        try:
            if resize_ppt_for_dock is False:
                from app import dock as _dock
                for h in list(_dock._ORIGINAL_PPT_RECTS.keys()):
                    _dock.restore_ppt_window(h)
        except Exception:  # noqa: BLE001
            pass
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

    def generate_variants(self, payload):
        # payload: {"template": name, "n": int, "content": {...}} OR
        #          {"templates": [names], "content": {...}}
        g = self._require_session()
        if g:
            return g
        act = {"type": "generate_variants"}
        act.update(payload or {})
        try:
            summary = self._com(lambda: self.dc.run_action(act))
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        self._log_turn("Generate variants", summary, {"actions": [act]})
        return {"ok": True, "summary": summary}

    def build_deck_from_spec(self, spec, clear_existing=False):
        g = self._require_session()
        if g:
            return g
        import json
        if isinstance(spec, str):
            try:
                spec = json.loads(spec)
            except Exception as e:  # noqa: BLE001
                return {"error": f"Spec is not valid JSON: {e}"}
        # Frozen engine contract: spec MUST be {"deck":[{template,
        # content},...]} (ExtractDeckSpecJson emits the same shape).
        # Accept a bare list for convenience and wrap it.
        if isinstance(spec, list):
            spec = {"deck": spec}
        if not isinstance(spec, dict) or "deck" not in spec:
            return {"error": 'spec must be a list of slides or '
                             '{"deck":[{"template":...,"content":...}]}'}
        act = {"type": "build_deck_from_spec", "spec": spec}
        if clear_existing:
            act["clear_existing"] = True
        try:
            summary = self._com(lambda: self.dc.run_action(act))
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        self._log_turn("Build deck from spec", summary,
                       {"actions": [act]})
        return {"ok": True, "summary": summary}

    def extract_spec(self):
        g = self._require_session()
        if g:
            return g
        try:
            js = self._com(self.dc.get_deck_spec)
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        return {"ok": True, "spec": js}

    def fill_with_ai(self, template, brief):
        if not secrets.has_api_key():
            return {"error": "Save your LLM API key first."}
        from app.llm_client import build_fill_prompt
        from app.template_slots import BUILTIN_SLOTS
        slots = BUILTIN_SLOTS.get(template)
        if slots is None:
            cap = self.list_captured_templates()["templates"]
            m = next((c for c in cap if c["name"] == template), None)
            slots = m["slots"] if m else []
        if not slots:
            return {"error": f"Unknown template '{template}'."}
        llm = LLMClient(self.settings, secrets.get_api_key() or "")
        try:
            raw = llm.raw(build_fill_prompt(slots, brief))
        except Exception as e:  # noqa: BLE001
            import httpx
            if isinstance(e, httpx.HTTPStatusError):
                body = ""
                try:
                    body = e.response.text[:800]
                except Exception:
                    pass
                return {"error": f"LLM API error "
                                 f"{e.response.status_code}: {body}"}
            return {"error": f"{type(e).__name__}: {e}"}
        import json
        import re
        s = re.sub(r"```[a-zA-Z]*\n?", "", raw).replace("```", "").strip()
        i, j = s.find("{"), s.rfind("}")
        if i == -1 or j == -1:
            return {"error": "AI did not return JSON."}
        try:
            content = json.loads(s[i:j + 1])
        except Exception as e:  # noqa: BLE001
            return {"error": f"AI JSON parse failed: {e}"}
        return {"ok": True, "content": content}

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

    def set_window_title(self, title):
        # Narrow cosmetic-only addition (SP5). Sets the OS window title bar
        # so the consumer UI can show the current deck name + dirty marker.
        try:
            if webview.windows:
                webview.windows[0].title = str(title or "Decko")
            return {"ok": True}
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}

    def pick_pptx_path(self):
        # Native file picker for consumer flow. Returns the selected path or
        # empty string on cancel. Stays on the UI thread (no COM).
        try:
            if not webview.windows:
                return ""
            res = webview.windows[0].create_file_dialog(
                webview.OPEN_DIALOG, allow_multiple=False,
                file_types=("PowerPoint files (*.pptx;*.pptm)", "All files (*.*)"))
            if not res:
                return ""
            return res[0] if isinstance(res, (list, tuple)) else str(res)
        except Exception:  # noqa: BLE001
            return ""

    def window_minimize(self):
        # SP7: bypass pywebview's Window.minimize and call ShowWindow
        # directly. After SP7's WS_EX_APPWINDOW fix the taskbar restore
        # works natively even on frameless windows.
        try:
            if sys.platform.startswith("win"):
                import ctypes
                SW_MINIMIZE = 6
                hwnd = _decko_hwnd()
                if hwnd:
                    ctypes.windll.user32.ShowWindow(int(hwnd), SW_MINIMIZE)
                    return {"ok": True}
            # Non-Windows fallback to pywebview's helper.
            if webview.windows:
                w = webview.windows[0]
                if hasattr(w, "minimize"):
                    w.minimize()
            return {"ok": True}
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}

    def window_close(self):
        # SP6: custom title-bar control for frameless dock mode.
        try:
            if webview.windows:
                webview.windows[0].destroy()
            return {"ok": True}
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}

    def set_dock_mode(self, enabled):
        # SP6: live toggle between dock (frameless, snapped to PPT) and
        # detached (framed, free-floating SP5 layout). Persists. Does NOT
        # recreate the window — mutates style + geometry in place so chat
        # state survives.
        from app import dock as _dock
        enabled = bool(enabled)
        try:
            # Persist.
            self.settings.dock_mode = enabled
            pr = load_persisted()
            save_persisted(self.settings,
                           pr.get("last_deck_path", ""),
                           pr.get("last_mode", "attach"))
            # Apply at runtime.
            if not webview.windows:
                return {"ok": True, "dock_mode": enabled}
            w = webview.windows[0]
            if enabled:
                # Switch to docked layout.
                _stop_existing_dock_loop()
                rect = _dock.compute_dock_rect(_dock.find_ppt_window())
                _apply_frameless_style(True)
                try:
                    w.move(rect[0], rect[1])
                    w.resize(rect[2], rect[3])
                except Exception:  # noqa: BLE001
                    pass
                # Force taskbar entry on frameless dock window (SP7).
                _force_app_window_taskbar(_decko_hwnd())
                _start_dock_loop_for_window(w)
            else:
                # Switch to detached SP5 layout (1180x760, framed, no dock).
                # SP7: restore any PowerPoint window we shrank.
                try:
                    for h in list(_dock._ORIGINAL_PPT_RECTS.keys()):
                        _dock.restore_ppt_window(h)
                except Exception:  # noqa: BLE001
                    pass
                _stop_existing_dock_loop()
                _apply_frameless_style(False)
                try:
                    mleft, mtop, mright, mbottom = _dock._monitor_rect_for_hwnd(0)
                    cx = mleft + ((mright - mleft) - 1180) // 2
                    cy = mtop + ((mbottom - mtop) - 760) // 2
                    w.move(cx, cy)
                    w.resize(1180, 760)
                except Exception:  # noqa: BLE001
                    pass
            return {"ok": True, "dock_mode": enabled}
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}

    def shutdown(self):
        # SP7: restore any PowerPoint window we shrank for dock mode.
        try:
            from app import dock as _dock
            for h in list(_dock._ORIGINAL_PPT_RECTS.keys()):
                _dock.restore_ppt_window(h)
        except Exception:  # noqa: BLE001
            pass
        try:
            _stop_existing_dock_loop()
        except Exception:  # noqa: BLE001
            pass
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


# ----------------------------------------------------------------------
# SP6 dock-loop helpers (module-level so set_dock_mode + main can share).

_DOCK_LOOP = None


def _stop_existing_dock_loop():
    global _DOCK_LOOP
    if _DOCK_LOOP is None:
        return
    try:
        from app import dock
        dock.stop_dock_loop(_DOCK_LOOP)
    except Exception:  # noqa: BLE001
        pass
    _DOCK_LOOP = None


def _decko_hwnd():
    """Return the OS hwnd of the active Decko window, or 0."""
    try:
        if not webview.windows:
            return 0
        w = webview.windows[0]
        hwnd = getattr(getattr(w, "native", None), "Handle", None)
        if hwnd:
            return int(hwnd)
    except Exception:  # noqa: BLE001
        pass
    # Fallback: enumerate top-level windows and match title "Decko".
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        hits = []

        def _cb(hwnd, _):
            if not user32.IsWindowVisible(hwnd):
                return True
            n = user32.GetWindowTextLengthW(hwnd)
            buf = ctypes.create_unicode_buffer(n + 1)
            user32.GetWindowTextW(hwnd, buf, n + 1)
            if buf.value and buf.value.startswith("Decko"):
                hits.append(int(hwnd))
            return True

        EnumWindowsProc = ctypes.WINFUNCTYPE(
            wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        user32.EnumWindows(EnumWindowsProc(_cb), 0)
        return hits[0] if hits else 0
    except Exception:  # noqa: BLE001
        return 0


def _apply_frameless_style(frameless: bool):
    """Mutate the live Decko window's GWL_STYLE to add/remove WS_CAPTION +
    WS_THICKFRAME. Lets us toggle dock ↔ detached without recreating the
    window (which would destroy chat state)."""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        user32 = ctypes.windll.user32
        GWL_STYLE = -16
        WS_CAPTION = 0x00C00000
        WS_THICKFRAME = 0x00040000
        WS_SYSMENU = 0x00080000
        WS_MINIMIZEBOX = 0x00020000
        WS_MAXIMIZEBOX = 0x00010000
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        SWP_NOZORDER = 0x0004
        SWP_FRAMECHANGED = 0x0020
        hwnd = _decko_hwnd()
        if not hwnd:
            return
        # GetWindowLongW returns LONG; on 64-bit you'd want GetWindowLongPtrW,
        # but ctypes exposes the W variant which is sufficient for style.
        cur = user32.GetWindowLongW(hwnd, GWL_STYLE)
        mask = WS_CAPTION | WS_THICKFRAME | WS_SYSMENU | WS_MINIMIZEBOX | WS_MAXIMIZEBOX
        if frameless:
            new = cur & ~mask
        else:
            new = cur | mask
        user32.SetWindowLongW(hwnd, GWL_STYLE, new)
        user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER
                            | SWP_FRAMECHANGED)
    except Exception:  # noqa: BLE001
        pass


def _dock_event_handler(name: str, payload: dict):
    """Translate dock-loop events into pywebview window mutations.
    Runs on the dock thread — pywebview's move/resize/hide/show/minimize
    are thread-safe shims that marshal back to the UI thread internally.

    SP7 changes:
      - Decko no longer hides on PPT minimize (we don't even receive
        that event anymore — see app.dock).
      - Slideshow lowers Decko's z-order instead of hiding it.
      - On move_resize / restore: optionally shrink PPT to free Decko's
        width when Settings.resize_ppt_for_dock=True.
    """
    if not webview.windows:
        return
    w = webview.windows[0]
    try:
        if name in ("move_resize", "restore", "slideshow_exit", "ppt_gone"):
            rect = payload.get("rect")
            ppt_hwnd = payload.get("ppt_hwnd") or 0
            if rect:
                x, y, ww, hh = rect
                try:
                    w.show()
                except Exception:  # noqa: BLE001
                    pass
                try:
                    w.move(x, y)
                    w.resize(ww, hh)
                except Exception:  # noqa: BLE001
                    pass
                # SP7 reflow PPT to free Decko's gutter when enabled.
                if ppt_hwnd and name in ("move_resize", "restore"):
                    try:
                        from app.config import settings_from_persisted
                        s = settings_from_persisted()
                        if s.resize_ppt_for_dock:
                            from app import dock as _dock
                            _dock.reflow_ppt_window(ppt_hwnd, x, ww)
                    except Exception:  # noqa: BLE001
                        pass
        elif name == "slideshow_lower":
            slideshow_hwnd = payload.get("slideshow_hwnd") or 0
            try:
                _lower_decko_zorder_behind(slideshow_hwnd)
            except Exception:  # noqa: BLE001
                pass
    except Exception:  # noqa: BLE001
        pass


def _lower_decko_zorder_behind(other_hwnd: int):
    """Drop Decko's z-order below `other_hwnd` without hiding the window.
    User can still alt-tab back to Decko anytime."""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        decko_hwnd = _decko_hwnd()
        if not decko_hwnd or not other_hwnd:
            return
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        SWP_NOACTIVATE = 0x0010
        ctypes.windll.user32.SetWindowPos(
            int(decko_hwnd), int(other_hwnd), 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE)
    except Exception:  # noqa: BLE001
        pass


def _force_app_window_taskbar(hwnd: int):
    """Force WS_EX_APPWINDOW (taskbar entry) and clear WS_EX_TOOLWINDOW.
    Frameless pywebview windows sometimes default to tool-window style
    which hides them from the taskbar — fatal for a 'dock' UX since
    minimized Decko can't be restored without a taskbar button.

    Cycle Hide→Show with SWP_NOACTIVATE so Windows re-registers the
    taskbar entry without stealing focus from PowerPoint."""
    if not sys.platform.startswith("win") or not hwnd:
        return
    try:
        import ctypes
        user32 = ctypes.windll.user32
        GWL_EXSTYLE = -20
        WS_EX_APPWINDOW = 0x00040000
        WS_EX_TOOLWINDOW = 0x00000080
        SW_HIDE = 0
        SW_SHOWNA = 8  # show without activating
        # Prefer the Ptr variant if available (64-bit).
        get_ex = getattr(user32, "GetWindowLongPtrW", None) \
            or user32.GetWindowLongW
        set_ex = getattr(user32, "SetWindowLongPtrW", None) \
            or user32.SetWindowLongW
        cur = get_ex(int(hwnd), GWL_EXSTYLE)
        new = (cur | WS_EX_APPWINDOW) & ~WS_EX_TOOLWINDOW
        set_ex(int(hwnd), GWL_EXSTYLE, new)
        # Cycle visibility to force taskbar re-registration.
        user32.ShowWindow(int(hwnd), SW_HIDE)
        user32.ShowWindow(int(hwnd), SW_SHOWNA)
    except Exception:  # noqa: BLE001
        pass


def _apply_on_top(hwnd: int, on_top: bool):
    """SetWindowPos with HWND_TOPMOST / HWND_NOTOPMOST. Lets the user
    toggle 'Keep Decko on top' live without restarting."""
    if not sys.platform.startswith("win") or not hwnd:
        return
    try:
        import ctypes
        user32 = ctypes.windll.user32
        HWND_TOPMOST = -1
        HWND_NOTOPMOST = -2
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        SWP_NOACTIVATE = 0x0010
        user32.SetWindowPos(int(hwnd),
                            HWND_TOPMOST if on_top else HWND_NOTOPMOST,
                            0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE)
    except Exception:  # noqa: BLE001
        pass


def _start_dock_loop_for_window(_w):
    """Find the OS hwnd of the Decko window and install the dock loop."""
    global _DOCK_LOOP
    try:
        from app import dock
        hwnd = _decko_hwnd()
        _DOCK_LOOP = dock.start_dock_loop(hwnd, on_dock_event=_dock_event_handler)
    except Exception:  # noqa: BLE001
        _DOCK_LOOP = None


def _on_window_ready():
    """Pywebview start() callback — runs after the window is created."""
    from app.config import settings_from_persisted
    s = settings_from_persisted()
    # SP7: force taskbar entry on every launch (frameless or not) so the
    # user can always restore Decko from the taskbar after minimizing.
    try:
        time.sleep(0.2)  # give pywebview time to register OS window
    except Exception:  # noqa: BLE001
        pass
    decko_hwnd = _decko_hwnd()
    _force_app_window_taskbar(decko_hwnd)
    _apply_on_top(decko_hwnd, bool(s.decko_on_top))
    if s.dock_mode:
        _start_dock_loop_for_window(webview.windows[0] if webview.windows else None)


def main() -> int:
    if "--selfcheck" in sys.argv[1:]:
        return _selfcheck()
    api = Api()
    from app.config import settings_from_persisted
    s = settings_from_persisted()
    if s.dock_mode:
        # Compute initial dock geometry against the active PowerPoint window.
        try:
            from app import dock as _dock
            x, y, w, h = _dock.compute_dock_rect(_dock.find_ppt_window())
        except Exception:  # noqa: BLE001
            x, y, w, h = 100, 100, 380, 720
        # SP7: drop on_top=True default; user controls via Settings.
        webview.create_window("Decko", _web_index(), js_api=api,
                              width=w, height=h, x=x, y=y,
                              frameless=True, easy_drag=True,
                              text_select=True)
    else:
        webview.create_window("Decko", _web_index(), js_api=api,
                              width=1180, height=760, text_select=True)
    webview.start(_on_window_ready)
    api.shutdown()
    return 0


if __name__ == "__main__":
    sys.exit(main())
