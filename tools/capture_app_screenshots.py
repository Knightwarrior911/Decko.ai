"""Capture annotated screenshots of the Decko Desktop frontend.

Strategy:
- Boot pywebview pointing at app/web/index.html with a STUB Api so the page
  loads without PowerPoint or a real LLM.
- Use evaluate_js to drive the UI into each of the 5 documented states.
- Grab the window region via PIL.ImageGrab (Windows).

Usage: python tools/capture_app_screenshots.py <prefix>
  <prefix> = "before" or "after"
Saves to: docs/screenshots/consumer-polish/<prefix>_NN_<name>.png

NOTE: the stub Api below mirrors `app.main:Api` method names exactly so
app.js works against it unchanged. It does NOT call PowerPoint or any LLM.
"""
from __future__ import annotations

import ctypes
import sys
import time
from pathlib import Path

import webview
import win32gui
from PIL import ImageGrab

REPO = Path(__file__).resolve().parent.parent
WEB_INDEX = REPO / "app" / "web" / "index.html"
OUT_DIR = REPO / "docs" / "screenshots" / "consumer-polish"
OUT_DIR.mkdir(parents=True, exist_ok=True)


class StubApi:
    """Drop-in stand-in for app.main.Api — canned responses only."""

    def __init__(self):
        self._has_key = False
        self._sessions = []
        self._provider = "anthropic"
        self._model = "claude-opus-4-7"
        self._base_url = ""
        self._title = "Decko"

    def boot(self):
        return {
            "has_key": self._has_key,
            "settings": {
                "provider": self._provider,
                "model": self._model,
                "base_url": self._base_url,
            },
            "last_deck_path": "",
            "last_mode": "attach",
            "sessions": self._sessions,
        }

    def save_settings(self, provider, model, base_url, api_key):
        self._provider = provider or "anthropic"
        self._model = model or self._model
        self._base_url = base_url or ""
        if api_key:
            self._has_key = True
        return {"ok": True}

    def new_session(self):
        return {"ok": True}

    def open_session(self, mode, file_path=""):
        title = "Acme Pitch.pptx — May 20 16:00"
        self._sessions = [{"id": 1, "title": title, "turn_count": 0}]
        return {"ok": True, "session_id": 1, "title": title}

    def send(self, text):
        return {"ok": True, "summary": "Updated slide 3 title and bullets.",
                "warnings": 0}

    def list_sessions(self):
        return {"sessions": self._sessions}

    def load_session(self, session_id):
        return {"turns": []}

    def save_powerpoint(self):
        return {"ok": True}

    def set_window_title(self, title):
        if webview.windows:
            webview.windows[0].title = title
        return {"ok": True}

    def pick_pptx_path(self):
        return ""

    def list_builtin_templates(self):
        return {"templates": [
            {"name": "title",     "slots": ["title", "subtitle"]},
            {"name": "section",   "slots": ["section_number", "section_title"]},
            {"name": "bullets",   "slots": ["heading", "bullets"]},
            {"name": "two_col",   "slots": ["heading", "left_body", "right_body"]},
            {"name": "comparison","slots": ["heading", "left_label", "left_body",
                                            "right_label", "right_body"]},
            {"name": "kpi_dashboard","slots": ["heading", "tiles"]},
            {"name": "quote",     "slots": ["quote_text", "attribution"]},
        ]}

    def list_captured_templates(self):
        return {"templates": [
            {"name": "Brand cover",   "slots": ["title", "subtitle"]},
            {"name": "KPI grid",      "slots": ["heading", "tiles"]},
        ]}

    def apply_template(self, *a, **kw): return {"ok": True, "summary": "Applied."}
    def capture_template(self, name):   return {"ok": True, "summary": "Captured."}
    def delete_template(self, name):    return {"ok": True}
    def rename_template(self, *a, **kw):return {"ok": True}
    def generate_variants(self, p):     return {"ok": True, "summary": "3 variants appended."}
    def build_deck_from_spec(self, *a, **kw): return {"ok": True, "summary": "Built 5 slides."}
    def extract_spec(self):             return {"ok": True, "spec": '{"deck":[]}'}
    def fill_with_ai(self, *a, **kw):   return {"ok": True, "content": {}}


def _find_decko_window():
    target = []

    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            t = win32gui.GetWindowText(hwnd)
            if t and t.startswith("Decko"):
                target.append((hwnd, t))
        return True

    win32gui.EnumWindows(cb, None)
    return target[0] if target else None


def _raise_window(hwnd):
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    try:
        ctypes.windll.user32.ShowWindow(hwnd, 9)
        cur = ctypes.windll.kernel32.GetCurrentThreadId()
        tgt = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, 0)
        ctypes.windll.user32.AttachThreadInput(tgt, cur, True)
        ctypes.windll.user32.BringWindowToTop(hwnd)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        ctypes.windll.user32.AttachThreadInput(tgt, cur, False)
    except Exception:
        pass


def _grab(out_path: Path):
    hit = _find_decko_window()
    if not hit:
        print(f"  WARN: no Decko window found for {out_path.name}")
        return
    hwnd, _title = hit
    _raise_window(hwnd)
    time.sleep(0.6)
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    img = ImageGrab.grab(bbox=(left, top, right, bottom), all_screens=True)
    img.save(str(out_path))
    print(f"  saved {out_path.name} ({img.size})")


def _capture_sequence(prefix: str):
    w = webview.windows[0]
    time.sleep(1.4)

    # 1. First launch — wizard step 1 (after polish) OR raw sidebar (before).
    #    The current HTML auto-opens wizard when has_key=false. The OLD HTML
    #    has no wizard, so the same screenshot shows the dense sidebar.
    _grab(OUT_DIR / f"{prefix}_01_first_launch.png")

    # 2. Idle empty state. Hide wizard if present, show hero/main pane.
    w.evaluate_js("""
        var wz = document.getElementById('wizard');
        if (wz) wz.classList.add('hidden');
        var hero = document.getElementById('hero');
        if (hero) hero.classList.remove('hidden');
        var cw = document.getElementById('composerWrap');
        if (cw) cw.classList.add('hidden');
    """)
    time.sleep(0.5)
    _grab(OUT_DIR / f"{prefix}_02_idle_empty.png")

    # 3. Composer state. Force a session illusion + show composer + chips.
    w.evaluate_js("""
        var hero = document.getElementById('hero');
        if (hero) hero.classList.add('hidden');
        var cw = document.getElementById('composerWrap');
        if (cw) cw.classList.remove('hidden');
        var th = document.getElementById('threadTitle');
        if (th) th.textContent = 'Acme Pitch.pptx — May 20 16:00';
        var ds = document.getElementById('deckStatus');
        if (ds) ds.innerHTML = '<b>Acme Pitch.pptx</b>';
        var t = document.getElementById('thread');
        if (t) {
            t.innerHTML = '';
            ['user','app'].forEach(function(k,i){
                var d = document.createElement('div');
                d.className = 'bubble ' + k;
                d.textContent = (i===0)
                    ? 'Rewrite slide 3 for Acme.'
                    : 'Updated slide 3 title and bullets to reflect Acme.';
                t.appendChild(d);
            });
        }
        var m = document.getElementById('msg');
        if (m) m.value = 'Apply our brand colors';
    """)
    time.sleep(0.5)
    _grab(OUT_DIR / f"{prefix}_03_composer.png")

    # 4. Templates panel open.
    w.evaluate_js("""
        var tp = document.getElementById('tplPanel');
        if (tp) tp.classList.remove('tpl-hidden');
        if (window.tplInit) try { window.tplInit(); } catch(e) {}
    """)
    time.sleep(0.7)
    _grab(OUT_DIR / f"{prefix}_04_templates_panel.png")

    # 5. Settings dialog / sidebar LLM section.
    w.evaluate_js("""
        var tp = document.getElementById('tplPanel');
        if (tp) tp.classList.add('tpl-hidden');
        var sd = document.getElementById('settingsDialog');
        if (sd) sd.classList.remove('hidden');
    """)
    time.sleep(0.5)
    _grab(OUT_DIR / f"{prefix}_05_settings_inline.png")

    time.sleep(0.4)
    w.destroy()


def main(argv):
    if len(argv) < 2:
        print("Usage: capture_app_screenshots.py <before|after>")
        return 2
    prefix = argv[1]
    api = StubApi()
    webview.create_window("Decko", str(WEB_INDEX), js_api=api,
                          width=1180, height=760, text_select=True)
    webview.start(_capture_sequence, args=(prefix,))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
