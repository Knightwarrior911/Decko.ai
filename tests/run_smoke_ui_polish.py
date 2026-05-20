"""UI polish gate (SP5 consumer-readiness).

Verifies the Decko Desktop frontend (`app/web/`) has shed engineer-test
jargon AND grown the consumer-grade affordances spelled out in
`docs/superpowers/specs/2026-05-20-desktop-consumer-polish.md`.

Two-part check:

1. JARGON BLOCKLIST — every string in BLOCKLIST must be ABSENT from
   `app/web/index.html` AND `app/web/app.js`.
2. SELECTOR ALLOWLIST — every selector in REQUIRED_SELECTORS must be
   PRESENT somewhere across `index.html` + `app.css` + `app.js`.

Substring matching (not full regex) keeps the gate cheap and obvious.
"""
from __future__ import annotations

import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
WEB = REPO / "app" / "web"
HTML = WEB / "index.html"
CSS = WEB / "app.css"
JS = WEB / "app.js"

BLOCKLIST = [
    "Deck DNA",
    "Decks-as-code",
    "Build deck from spec",
    "Extract spec from deck",
    "base URL (generic only)",
    "Attach to open deck",
]

# Each item: (label, must_appear_in_any_of[paths], substring)
REQUIRED_SELECTORS = [
    ("wizard root",        [HTML],        'id="wizard"'),
    ("wizard step 1",      [HTML],        'id="wizardStep1"'),
    ("wizard step 2",      [HTML],        'id="wizardStep2"'),
    ("wizard step 3",      [HTML],        'id="wizardStep3"'),
    ("settings dialog",    [HTML],        'id="settingsDialog"'),
    ("gear button",        [HTML],        'id="gearBtn"'),
    ("curated model pick", [HTML],        'id="modelPick"'),
    ("hero block",         [HTML],        'id="hero"'),
    ("hero connect btn",   [HTML],        'id="heroConnect"'),
    ("hero open-file btn", [HTML],        'id="heroOpenFile"'),
    ("composer chips",     [HTML],        'id="chips"'),
    ("chip class",         [HTML, CSS],   ".chip"),
    ("toast container",    [HTML],        'id="toast"'),
    ("theme attribute",    [CSS],         'data-theme'),
    ("tooltip attribute",  [HTML],        'title="'),
    ("Decko wordmark",     [HTML],        'class="wordmark"'),
    ("My saved layouts",   [HTML],        "My saved layouts"),
    ("Power user tools",   [HTML],        "Power user tools"),
    ("Connect to open deck (hero)", [HTML], "Connect to open deck"),
    ("friendly-error helper",        [JS], "friendlyError"),
    ("showToast helper",             [JS], "showToast"),
    ("set_window_title call",        [JS], "set_window_title"),
    ("MODELS_BY_PROVIDER constant",  [JS], "MODELS_BY_PROVIDER"),
    ("composer placeholder rewrite", [HTML], "What would you like to change?"),
]


def _read(p: Path) -> str:
    if not p.exists():
        return ""
    return p.read_text(encoding="utf-8")


def check_blocklist() -> list[str]:
    fails: list[str] = []
    html = _read(HTML)
    js = _read(JS)
    for needle in BLOCKLIST:
        if needle in html:
            fails.append(f"jargon '{needle}' still present in index.html")
        if needle in js:
            fails.append(f"jargon '{needle}' still present in app.js")
    return fails


def check_selectors() -> list[str]:
    fails: list[str] = []
    cache: dict[Path, str] = {p: _read(p) for p in (HTML, CSS, JS)}
    for label, paths, needle in REQUIRED_SELECTORS:
        if not any(needle in cache.get(p, "") for p in paths):
            where = ", ".join(p.name for p in paths)
            fails.append(f"missing {label!r}: '{needle}' not found in {where}")
    return fails


def main() -> int:
    if not HTML.exists() or not CSS.exists() or not JS.exists():
        print(f"FAIL: app/web sources missing under {WEB}")
        return 1
    fails = check_blocklist() + check_selectors()
    if fails:
        print(f"run_smoke_ui_polish: FAIL ({len(fails)} issue(s))")
        for f in fails:
            print(f"  - {f}")
        return 1
    print("run_smoke_ui_polish: PASS — jargon clean, selectors present.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
