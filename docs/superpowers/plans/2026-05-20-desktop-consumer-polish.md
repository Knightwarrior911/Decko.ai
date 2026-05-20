# Decko Desktop — Consumer Polish (SP5) — Implementation Plan

Source spec: `docs/superpowers/specs/2026-05-20-desktop-consumer-polish.md`
Branch: `feat/desktop-consumer-polish` off `main` (HEAD `fba2dea`)
Autonomy: pre-approved under `/goal` lock — execute spec→plan→build→verify→PR.

## Task graph

T1 ── T2 ── T3 ── T4 ── T5 ── T6 ── T7 ── T8
                       (gates)        (push)

### T1. Write plan (this doc)

DONE on creation.

### T2. Branch + smoke scaffold

- `git checkout -b feat/desktop-consumer-polish`
- Create `tests/run_smoke_ui_polish.py`:
  - Reads `app/web/index.html` and `app/web/app.js` as text.
  - Asserts every jargon string in the blocklist is absent (spec §4.1).
  - Asserts required DOM selector ids/classes exist (spec §4.2-§4.9):
    `#wizard`, `#wizardStep1`, `#wizardStep2`, `#wizardStep3`,
    `#settingsDialog`, `#gearBtn`, `#modelPick`, `#hero`, `#heroConnect`,
    `#heroOpenFile`, `#chips`, `.chip`, `#toast`, `[data-theme]`, `title=`.
  - Exit 0 on pass, 1 on first failure with a diff-style message.
- Register in `tests/run_smoke_app.py` GATES list as `ui_polish`.
- Run baseline: `run_smoke_app.py` should now FAIL on `ui_polish` and PASS the
  other gates. That's the desired "red" state before T4.

### T3. BEFORE screenshots

- New helper `tools/capture_app_screenshots.py`:
  - Boot pywebview in a child process (no real LLM, no real PowerPoint —
    stub via env var `DECKO_NO_LAUNCH=1` if already supported, else just open
    the static HTML in `pywebview` and screenshot the chrome).
  - Use `webview.windows[0].evaluate_js` to set DOM state, then
    `webview.windows[0].create_confirmation_dialog` is N/A — use
    `pyautogui` or pywebview's native screenshot if available; fallback to
    Windows `gdi32`-based PIL ImageGrab of the window region.
- Capture 5 PNGs to `docs/screenshots/consumer-polish/before_*.png`:
  1. `before_01_first_launch.png` — initial sidebar.
  2. `before_02_idle_empty.png` — `#chat` empty state.
  3. `before_03_composer.png` — composer area.
  4. `before_04_templates_panel.png` — `Templates ▸` open.
  5. `before_05_settings_inline.png` — sidebar LLM section.

If headless capture proves flaky, fall back to a documented manual capture:
write a `BEFORE_SCREENSHOTS.md` next to the PNGs describing the manual steps
the maintainer ran, and stage the PNGs as committed artifacts. (Honest scope —
matches spec §6.)

### T4. Frontend rewrite (atomic, single commit)

Order of work inside the rewrite:

1. **`app/web/decko_icon.svg`** — 32×32 stylized D in a slide rectangle,
   accent #3a6ea5.
2. **`app/web/app.css`** — full rewrite:
   - CSS variables under `:root` and `[data-theme="light"]` (spec §4.12).
   - Layout grid: header bar (48px) + body (flex: sidebar 300px + chat).
   - Wizard overlay styles.
   - Settings dialog styles (modal, max-width 480px).
   - Hero styles (centered card, 2 large buttons).
   - Chip styles (rounded pills above composer).
   - Toast container styles (bottom-right stack).
   - Templates panel styles (kept right-side slide-over, refreshed colors).
   - Light theme overrides via attribute selector.
3. **`app/web/index.html`** — full rewrite:
   - `<header>` with icon + Decko wordmark + gear button.
   - `<aside id="side">` with Deck status + sessions list + Templates button.
     (LLM section moved into Settings dialog.)
   - `<main id="chat">` containing `#hero` (shown when no session) and
     `#thread` + chips + composer (shown when session live).
   - `<div id="wizard">` overlay with 3 step panels.
   - `<dialog id="settingsDialog">` with provider/model/key/theme.
   - `<aside id="tplPanel">` rewritten w/ rename pass (My saved layouts,
     Power user tools).
   - `<div id="toast"></div>` container.
4. **`app/web/app.js`** — full rewrite:
   - `MODELS_BY_PROVIDER` constant.
   - `friendlyError(raw)` helper (spec §4.11 mapping).
   - `showToast(text, kind)`.
   - Wizard state machine (`#decko_wizard_done` localStorage flag).
   - Settings dialog open/close handlers calling existing
     `api.save_settings`.
   - Hero button handlers calling `api.open_session("attach")` and
     `api.open_session("file", path)` — `path` obtained via a small
     pywebview `create_file_dialog` flow exposed through a new
     `api.pick_pptx()` IF needed (check existing flow — current path is
     manual text input; for consumer parity, replace text input with native
     file picker via `webview.windows[0].create_file_dialog`).
   - Composer chips (click prefills `#msg`).
   - Dirty tracking + `#savePptBtn.dirty` class + `api.set_window_title(...)`
     calls on session-open, send-success, save-success.
   - Theme toggle: read/write `localStorage.decko_theme`, apply
     `document.documentElement.setAttribute("data-theme", v)`.

Note on file picker: the current backend `open_session(mode, file_path)`
takes a string path; if no Api method exposes a file dialog, add
`Api.pick_pptx_path()` that calls `webview.windows[0].create_file_dialog(
webview.OPEN_DIALOG, file_types=("PowerPoint Files (*.pptx;*.pptm)",))` and
returns the chosen path or `""`. This stays within the "no new Api methods
unless rename forces it" rule because the legacy text-input path is replaced
in the consumer wizard, and the engine is untouched.

### T5. Backend touch — `Api.set_window_title` (+ optional `pick_pptx_path`)

- Add `def set_window_title(self, title): ...` to `app/main.py:Api`. Guard
  `if webview.windows: webview.windows[0].title = title`.
- Add `def pick_pptx_path(self): ...` returning the chosen path string
  (or empty string on cancel).

### T6. Run all gates

Sequence:

1. `python tools/build_carrier.py` (idempotent — needed if carrier missing).
2. `python update_macros.py` — sync src → carrier (engine unchanged but
   keeps carrier consistent).
3. `python tools/sync_actions_guidance.py` — keep auto-appendix consistent
   (engine unchanged so should be a no-op).
4. `python tests/run_smoke_ui_polish.py` — primary gate.
5. `python tests/run_smoke_app.py` — full SP1 gate (now includes
   `ui_polish`).
6. `python tests/run_smoke.py` — engine regression guard.
7. `python tests/test_guidance_coverage.py`.
8. `python tests/test_guidance_doc_sync.py`.

If any step fails: fix forward — no scope creep. If a fix requires touching
the engine, STOP and revisit the plan with the user.

### T7. AFTER screenshots

Capture 5 paired PNGs to `docs/screenshots/consumer-polish/after_*.png`
matching the before set. Layout in a folder so reviewer can diff side-by-side.

### T8. Commit + push + PR

Conventional commit per logical chunk:

1. `docs: spec + plan for desktop consumer polish (SP5)`
2. `test(ui): add run_smoke_ui_polish.py — jargon blocklist + selector gate`
3. `feat(app): consumer-grade frontend rewrite (wizard, settings, hero, chips, theme)`
4. `feat(app): Api.set_window_title + Api.pick_pptx_path` (only if touched)
5. `docs(screenshots): before/after consumer-polish walkthrough`

`git push -u origin feat/desktop-consumer-polish` then `gh pr create` with a
PR body listing the 12 gaps + check marks + screenshot diffs.

## Stop condition

All four smoke gates green, 5 screenshot pairs committed, branch pushed, PR
open. Goal auto-clears.
