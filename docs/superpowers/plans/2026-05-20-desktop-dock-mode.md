# Decko Desktop — Dock Mode (SP6) — Implementation Plan

Source spec: `docs/superpowers/specs/2026-05-20-desktop-dock-mode.md`
Branch: `feat/desktop-dock-mode` off `feat/desktop-consumer-polish` (HEAD `5f64bd3`)
Autonomy: pre-approved under `/goal` lock.

## Tasks

### T1. Spec + plan (this doc + spec)
DONE on creation.

### T2. Branch + smoke gate scaffold

- `git checkout -b feat/desktop-dock-mode feat/desktop-consumer-polish`
- Create `tests/run_smoke_ui_dock.py`:
  - Imports `app.dock`; asserts `find_ppt_window`, `compute_dock_rect`,
    `start_dock_loop`, `stop_dock_loop` are callables.
  - Asserts `app.config.Settings.dock_mode` exists with default `True`.
  - Asserts `app/web/index.html` contains `id="undockBtn"`, `id="winMinBtn"`,
    `id="winCloseBtn"`, `class="dragArea"`.
  - Asserts `app/web/app.js` references `set_dock_mode`, `window_minimize`,
    `window_close`.
  - Asserts `app/web/app.css` has `[data-dock=` and `.winctl`.
  - Calls `compute_dock_rect(0, 380, 600)` with a stub hwnd path that
    returns a fake PPT rect — verifies math without a live PPT.
- Register in `tests/run_smoke_app.py` `GATES` as `ui_dock`.
- Run: should FAIL (red) before T3-T5 land.

### T3. `app/dock.py` snap engine

- ctypes bindings for `EnumWindows`, `GetWindowTextW`, `GetClassNameW`,
  `IsWindowVisible`, `GetWindowRect`, `IsIconic`, `IsWindow`,
  `MonitorFromWindow`, `GetMonitorInfoW`, `GetForegroundWindow`.
- ctypes bindings for `SetWinEventHook`, `UnhookWinEvent`,
  `GetMessageW`/`PeekMessageW` loop, `PostQuitMessage`.
- `find_ppt_window()`:
  - EnumWindows + class + title filtering per spec §4.1.
  - Last-foreground cache updated by hook callback.
- `compute_dock_rect(ppt_hwnd, width=380, min_height=600)`:
  - If `ppt_hwnd` is 0 or invalid → center on primary monitor.
  - Else `GetWindowRect(ppt_hwnd)` + `MonitorFromWindow` + GetMonitorInfo;
    do the math per spec.
- `DockLoop` dataclass.
- `start_dock_loop(decko_hwnd, on_dock_event)`:
  - Spawn daemon MSG-pump thread.
  - `SetWinEventHook(EVENT_OBJECT_LOCATIONCHANGE, EVENT_SYSTEM_FOREGROUND,
    None, _callback, 0, 0, WINEVENT_OUTOFCONTEXT)` and so on.
  - Callback emits structured events (`move_resize`, `minimize`, `restore`,
    `slideshow_enter`, `slideshow_exit`, `ppt_gone`).
  - 250ms polling fallback if hook returns 0.
- `stop_dock_loop(loop)`.

Unit-testable surface: `compute_dock_rect` is pure (no Win32 calls when
ppt_hwnd is None/0). `find_ppt_window` returns None on no PPT.

### T4. Wire `dock_mode` + new Api methods

- `app/config.py`: add `dock_mode: bool = True` to `Settings`; pass through
  `save_persisted`; load via `settings_from_persisted`.
- `app/main.py`:
  - Read dock_mode at startup.
  - Compute initial geometry from `find_ppt_window()` + `compute_dock_rect()`
    when dock_mode=True; SP5 default when False.
  - `webview.create_window(...)` with `frameless=dock_mode`, `easy_drag=
    dock_mode`, `on_top=dock_mode`, `width`, `height`, `x`, `y`.
  - `webview.start(_post_start, ...)` callback installs the dock loop and
    initial snap.
  - New Api methods: `window_minimize`, `window_close`, `set_dock_mode`.
    `set_dock_mode` mutates style via `SetWindowLongW(GWL_STYLE, ...)` +
    `SetWindowPos` to switch frameless ↔ framed without recreating the
    window (preserves chat state).
- Update `app.boot()` payload to include `dock_mode` so the frontend can
  reflect state.

### T5. Frontend updates

- `app/web/index.html`:
  - Add `<div class="dragArea">` inside `<header id="appHeader">` (between
    wordmark and headerSpacer).
  - Add `<button id="undockBtn" class="winctl" title="...">📌</button>`
    (or 📍 when dock_mode=false) before gear.
  - Add `<button id="winMinBtn" class="winctl">━</button>` and
    `<button id="winCloseBtn" class="winctl">✕</button>` after gear.
  - Settings dialog: add "Dock to PowerPoint" checkbox row.
- `app/web/app.js`:
  - Honor `bootSnapshot.settings.dock_mode` → set `document.body.dataset.dock`.
  - Wire undock pin → `api.set_dock_mode(!current)` and update icon +
    body.dataset.
  - Wire min/close buttons → `api.window_minimize/close`.
  - Settings dialog: read/write `dock_mode`.
- `app/web/app.css`:
  - `.winctl { width:28px; height:28px; padding:0; background:transparent;
    border:none; color:var(--muted); border-radius:6px; }`
  - `.winctl:hover { background:var(--surface-2); color:var(--text); }`
  - `.dragArea { flex:1; height:100%; -webkit-app-region: drag; }` (the
    pywebview easy_drag picks up any non-interactive element click but
    we keep the class for visual clarity + smoke detection).
  - `body[data-dock="on"]` rules: sidebar becomes drawer toggle, narrow
    composer, chips wrap.

### T6. Run all gates

`run_smoke_ui_dock` → `run_smoke_ui_polish` → `run_smoke_app` →
`run_smoke` → `test_guidance_coverage` → `test_guidance_doc_sync`.

### T7. 3 screenshots

`tools/capture_app_screenshots.py` already produces app screenshots with
a stub Api. Extend or create `tools/capture_dock_screenshots.py` that:

1. Boots Decko with dock_mode=True and a faked PPT-window stub (Tk window
   acting as a stand-in for PowerPoint) for `docked_with_ppt.png`.
2. Boots Decko with dock_mode=False for `undocked_detached.png`.
3. Boots Decko with dock_mode=True + faked slideshow window for
   `slideshow_hidden.png`.

If the Tk-stub-as-PPT approach proves flaky, fall back to honest manual
capture with a real PowerPoint open (document the steps in a README).

### T8. Commit + push + PR

Conventional commits per chunk:
1. `docs: spec + plan for desktop dock mode (SP6)`
2. `test(ui): add run_smoke_ui_dock.py — dock surface + Settings gate`
3. `feat(app): app/dock.py — snap engine (find/compute/start/stop)`
4. `feat(app): wire dock_mode into Settings + Api (set_dock_mode, window_minimize/close)`
5. `feat(app): frontend window controls + undock pin + dock-aware CSS (SP6)`
6. `docs(screenshots): dock-mode walkthrough (docked/undocked/slideshow)`

Push + `gh pr create`.

## Stop condition

All gates green AND lifecycle cases handled AND screenshots committed AND
PR open. Goal auto-clears.
