# Decko Desktop — Dock Mode (SP6)

**Status:** APPROVED — 2026-05-20
**Owner:** Claude (autonomous execution under /goal lock)
**Branch:** `feat/desktop-dock-mode` off `feat/desktop-consumer-polish` (HEAD `5f64bd3`)

## 1. Problem

After SP5, Decko Desktop is a polished consumer app but still LIVES as a
separate top-level window. Users perceive PowerPoint + Decko as two unrelated
apps to alt-tab between. The intent is one workspace: PowerPoint with Decko
pinned to its right edge as a sidebar.

## 2. Goal

Make Decko Desktop auto-dock to the right edge of the active PowerPoint
window — a floating-pinned sidebar that visually feels like part of the
PowerPoint workspace. Track PPT move/resize/minimize/foreground/slideshow
events and re-snap deterministically. Keep the legacy detached SP5 window
available via a Settings toggle.

## 3. Non-goals

- True Office task-pane integration (VSTO / office-js). Out of scope.
- Mac / Linux / cross-platform.
- New engine actions or Api method renames.
- Cloud sync, licensing, marketing.
- Replacing pywebview.

## 4. Architecture

### 4.1 New module: `app/dock.py`

Pure Windows-only. No COM. Talks `ctypes` + `win32gui` only. Owns the snap
engine and is unit-importable without launching pywebview.

Public surface:

```
find_ppt_window() -> Optional[int]
    Returns the hwnd of the active PowerPoint window to dock against, or
    None. Selection rule (highest priority first):
      1. Most-recently-foreground PPT hwnd (tracked via EVENT_SYSTEM_FOREGROUND).
      2. Any visible top-level window with class "PPTFrameClass".
      3. Any visible top-level window whose title ends with " - PowerPoint".
    Slideshow windows (class "screenClass") are EXCLUDED from this lookup —
    they're treated as the slideshow signal, not the dock target.

compute_dock_rect(ppt_hwnd, decko_width=380, decko_min_height=600)
    -> tuple[int, int, int, int]
    Returns (x, y, w, h) screen pixels for the Decko window.
      - y = PPT top.
      - h = max(PPT height, decko_min_height).
      - w = decko_width.
      - x: if PPT.right + decko_width <= monitor.right, dock flush to PPT
        right edge (x = PPT.right). Otherwise overlap PPT's last decko_width
        pixels (x = PPT.right - decko_width). PowerPoint redraws and the
        ribbon stays visible — content shifts left under the overlay.
    Monitor detected via MonitorFromWindow(ppt_hwnd, MONITOR_DEFAULTTONEAREST)
    + GetMonitorInfo. DPI is per-monitor; ctypes already process-DPI-aware
    after main.py boot.

start_dock_loop(decko_hwnd, on_dock_event=None) -> DockLoop
    Installs SetWinEventHook for:
      - EVENT_OBJECT_LOCATIONCHANGE (move/resize)
      - EVENT_SYSTEM_MINIMIZESTART  (minimize)
      - EVENT_SYSTEM_MINIMIZEEND    (restore)
      - EVENT_SYSTEM_FOREGROUND     (focus changes — also tracks slideshow window)
    Hook callback is dispatched on the UI thread (single MSG-pump thread we
    start internally). On each fired event, recompute target hwnd, recompute
    dock rect, schedule a webview move/resize via on_dock_event callback
    (NEVER reach into webview directly from the hook thread).
    Fallback: if SetWinEventHook returns 0 (failure), spin a daemon Timer
    polling at 250ms intervals.

stop_dock_loop(loop) -> None
    Unhook + quit the MSG pump + join the timer.

DockLoop dataclass (returned from start_dock_loop): hook handle(s),
foreground-pid-tracker, msg-pump-thread, polling-fallback-timer.
```

Slideshow detection: if `GetForegroundWindow` returns a window whose class
matches `screenClass` (PPT slideshow) AND covers ≥ 90% of any monitor → the
dock loop emits `event="slideshow_enter"`. On any subsequent
`EVENT_SYSTEM_FOREGROUND` whose target isn't a screenClass window, emit
`event="slideshow_exit"`.

PPT-closed detection: every dispatch checks if the last-known PPT hwnd is
still `IsWindow()`. When no PPT window is alive → emit `event="ppt_gone"`.

### 4.2 `app/config.py` change

Add `dock_mode: bool = True` to `Settings` dataclass. Persist through
`save_persisted` and `settings_from_persisted`. No migration needed — old
settings.json without the key reads as default True.

### 4.3 `app/main.py` changes

Three new `Api` methods, all cosmetic (no engine touch):

- `window_minimize()` — calls `webview.windows[0].minimize()` (or
  `ShowWindow(hwnd, SW_MINIMIZE)` via ctypes fallback).
- `window_close()` — calls `webview.windows[0].destroy()`.
- `set_dock_mode(enabled: bool)` — persists the new value via
  `save_persisted` AND applies at runtime: starts/stops the dock loop and
  resizes the window. Re-entry safe.

In `main()`:
- Read `dock_mode` from persisted settings.
- If `dock_mode=True`: `webview.create_window("Decko", ..., frameless=True,
  easy_drag=True, on_top=True, width=380, height=720, x=..., y=...)` —
  compute initial x/y from `find_ppt_window()` + `compute_dock_rect()`.
- If `dock_mode=False`: SP5 behavior unchanged (framed 1180×760, no on_top).
- After `webview.start(...)` callback fires, call `start_dock_loop(decko_hwnd,
  on_dock_event=_apply_dock_event)` where `_apply_dock_event` translates
  events to `webview.windows[0].move/resize/hide/show/minimize`.

### 4.4 Frontend changes (`app/web/`)

#### 4.4.1 `index.html`

Add three controls inside `<header id="appHeader">`:

- `<button id="undockBtn" title="Detach Decko from PowerPoint" class="winctl">📌</button>`
  — flips `dock_mode` via `api.set_dock_mode(false)`.
- `<button id="winMinBtn" title="Minimize" class="winctl">━</button>`
  — calls `api.window_minimize()`.
- `<button id="winCloseBtn" title="Close" class="winctl">✕</button>`
  — calls `api.window_close()`.

When in detached mode (SP5), `undockBtn` becomes "Dock to PowerPoint" (📍)
and flips back via `api.set_dock_mode(true)`. The min/close buttons hide
when not frameless (native OS chrome takes over).

Header gains an explicit `<div class="dragArea"></div>` for `easy_drag`
behavior (pywebview's easy_drag picks up clicks on non-interactive elements;
the dragArea is the explicit visual zone).

#### 4.4.2 `app.css`

New CSS:

- `body[data-dock="on"]` — narrow layout: collapse sidebar to a single
  "drawer" toggled by a tab icon (since 380px doesn't fit 280px sidebar +
  chat). Hero text condensed. Composer chips wrap to 2 lines.
- `.winctl` — 28×28 borderless buttons w/ hover state.
- `body[data-dock="off"]` — default SP5 layout, no changes.

#### 4.4.3 `app.js`

- Read `bootSnapshot.settings.dock_mode` (now persisted), set
  `document.body.dataset.dock = "on" | "off"`.
- Wire undockBtn/winMinBtn/winCloseBtn handlers.
- Settings dialog gains a "Dock to PowerPoint" checkbox bound to
  `api.set_dock_mode(...)`. Toggling it without saving the dialog still
  applies (live toggle).

## 5. Lifecycle table

| Event                              | Decko action                                          |
| ---------------------------------- | ----------------------------------------------------- |
| Initial launch + PPT open          | Snap to right edge of active PPT window.              |
| Initial launch + no PPT open       | Show centered on primary monitor (380×720). Hero says "No PowerPoint deck open." |
| PPT moves (drag, snap, Win+arrow)  | Recompute dock rect → move Decko on next frame.       |
| PPT resizes                        | Recompute dock rect → move + resize Decko.            |
| PPT minimized                      | Minimize Decko.                                       |
| PPT restored from minimize         | Restore Decko and re-snap.                            |
| PPT closed                         | Decko stays open, re-centers, shows "No PowerPoint deck open" hero. |
| PPT slideshow (F5) starts          | Hide Decko (visibility off). Track slideshow hwnd.    |
| PPT slideshow ends                 | Show Decko, re-snap.                                  |
| Multiple PPT windows               | Track last-foreground via EVENT_SYSTEM_FOREGROUND; snap to it. |
| PPT on different monitor than Decko| MonitorFromWindow → move Decko to PPT's monitor.      |
| User clicks Undock pin             | Persist `dock_mode=False`, stop hook, resize to 1180×760, switch to framed window. |
| User toggles Dock-to-PowerPoint ON | Persist `dock_mode=True`, start hook, switch to frameless + snap. |
| Decko killed (X)                   | Stop hook, save dock_mode unchanged.                  |

## 6. Files touched

- `app/dock.py` (new, ~250 lines).
- `app/config.py` (+1 field, +1 line in save).
- `app/main.py` (+3 Api methods, frameless toggle, dock loop lifecycle).
- `app/web/index.html` (header window controls + undockBtn).
- `app/web/app.css` (dock-on layout + winctl styles).
- `app/web/app.js` (handlers, body data-dock, settings checkbox).
- `tests/run_smoke_ui_dock.py` (new gate).
- `tests/run_smoke_app.py` (register `ui_dock` row).
- `docs/screenshots/dock-mode/` (3 PNGs + README).

## 7. Metric

Deterministic:
- `python tests/run_smoke_ui_dock.py` → PASS.
- `python tests/run_smoke_ui_polish.py` → PASS (unchanged).
- `python tests/run_smoke_app.py` → PASS (all 7 gates including new ui_dock).
- `python tests/run_smoke.py` → PASS.
- `python tests/test_guidance_coverage.py` + `test_guidance_doc_sync.py` → PASS.

Manual:
- 3 screenshots at `docs/screenshots/dock-mode/`:
  1. `docked_with_ppt.png` — Decko snapped to PowerPoint right edge.
  2. `undocked_detached.png` — SP5 fallback mode active.
  3. `slideshow_hidden.png` — PPT in slideshow, Decko absent.

## 8. Risks

- pywebview `frameless=True` + `on_top=True` combination is OS-quirky;
  verify on Windows 11 before depending on it. Fallback: keep native frame,
  emulate "dockedness" via window-position alone.
- `SetWinEventHook` requires a message pump on the registering thread; we
  start a dedicated thread with `PeekMessageW` loop. Must `PostQuitMessage`
  to shut down cleanly to avoid orphan threads on app exit.
- 250ms polling fallback can race with user drags producing visible jitter;
  prefer hook-based path. Polling exists only when hook registration fails
  (high-integrity processes, group policy lockdown).
- Switching `dock_mode` at runtime requires destroying + recreating the
  pywebview window OR mutating window styles via ctypes (`SetWindowLongW
  GWL_STYLE`). Choose the second path to avoid losing chat state.
- Headless CI: `app/dock.py` import path must succeed even when no PPT
  window exists; functions must return None / no-op rather than raise.

## 9. Out of scope (firm)

VSTO/office-js task pane, Mac/Linux, new engine actions, cloud sync,
licensing, marketing, replacing pywebview.
