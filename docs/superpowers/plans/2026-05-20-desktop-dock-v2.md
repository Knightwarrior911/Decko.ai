# Decko Desktop — Dock Mode v2 (SP7) — Plan

Source: `docs/superpowers/specs/2026-05-20-desktop-dock-v2.md`
Branch: `feat/desktop-dock-v2` off `feat/desktop-dock-mode`

## Tasks

### T1. Spec + plan (this doc)
DONE on creation.

### T2. Branch + update smoke gate

- `git checkout -b feat/desktop-dock-v2 feat/desktop-dock-mode`
- Edit `tests/run_smoke_ui_dock.py`:
  - Assert `Settings.decko_on_top` exists (default False).
  - Assert `Settings.resize_ppt_for_dock` exists (default True).
  - Assert `app.dock.reflow_ppt_window` + `restore_ppt_window` callable.
  - Assert index.html contains `id="setDeckoOnTop"` + `id="setResizePpt"`.

### T3. `app/dock.py` behavior changes

- Add `_ORIGINAL_PPT_RECTS: dict[int, tuple]` module cache.
- Add `reflow_ppt_window(ppt_hwnd, decko_left)` — caches current rect,
  shrinks PPT so `ppt.right == decko_left`. Gated by a
  `_REFLOWING` flag to avoid hook re-entry storms.
- Add `restore_ppt_window(ppt_hwnd)` — re-applies cached rect, then
  pops from cache.
- Drop the minimize handler's `webview.windows[0].minimize()` call in
  `app/main.py:_dock_event_handler` (Decko stays where it was).
- Drop `slideshow_enter` hide path; emit a new event `lower_z_order`
  the handler translates to `SetWindowPos(decko_hwnd, slideshow_hwnd,
  ..., SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE)`.
- `stop_dock_loop` calls `restore_ppt_window` for every cached hwnd.

### T4. `app/config.py` + `app/main.py` wiring

- `app/config.py`:
  - `Settings.decko_on_top: bool = False`
  - `Settings.resize_ppt_for_dock: bool = True`
  - Persist both through save/load.
- `app/main.py`:
  - Remove `on_top=True` from frameless `create_window`.
  - After `webview.start(...)` callback fires, run `_force_app_window(
    decko_hwnd)` — applies `WS_EX_APPWINDOW`, removes
    `WS_EX_TOOLWINDOW`, cycles Hide/Show to force taskbar registration.
  - `_dock_event_handler` updated:
    - On `move_resize`/`restore`: also call `dock.reflow_ppt_window`
      when `Settings.resize_ppt_for_dock=True`.
    - On `slideshow_enter`: drop hide, schedule `SetWindowPos` lowering
      Decko behind slideshow.
    - On `minimize`: no-op (drop SP6's hide).
  - `Api.window_minimize` uses ctypes `ShowWindow(hwnd, SW_MINIMIZE)`.
  - `Api.set_dock_mode(False)` + `Api.shutdown` call
    `dock.restore_ppt_window` for any cached PPT hwnd.
  - `Api.boot` payload includes the two new Settings fields.
  - `Api.save_settings` accepts the two new bools (extend signature
    with kwargs OR derive from the JSON body — but signature is JS-
    facing, so add explicit kwargs with defaults).

### T5. Frontend Settings checkboxes

- `app/web/index.html`: two new checkboxes in the Settings dialog:
  - `<input type="checkbox" id="setDeckoOnTop">` Keep Decko on top
  - `<input type="checkbox" id="setResizePpt" checked>` Resize
    PowerPoint to make room
- `app/web/app.js`:
  - On Settings dialog open: read `bootSnapshot.settings.decko_on_top`
    and `.resize_ppt_for_dock` into the checkboxes.
  - On Settings save: pass both fields to `api.save_settings`.
  - On checkbox change (live toggle): call `api.save_settings` with
    the new value so behavior applies without OK/Cancel.

### T6. Gates + screenshot + PR

- Run `tests/run_smoke_ui_dock.py` (now extended).
- Run `tests/run_smoke_app.py` (all 7 gates).
- Run `tests/run_smoke.py` (engine).
- Run `tests/test_guidance_coverage.py` + `test_guidance_doc_sync.py`.
- Produce 1 composite at `docs/screenshots/dock-v2/reflowed_side_by_side.png`
  showing PPT shrunk to make room for Decko (no overlap). Use a small
  variant of `tools/capture_dock_screenshots.py` — extend with a
  `_build_reflowed()` function that paints PPT narrower so the dock
  sits in dedicated gutter space.
- Conventional commits per chunk. Push. Open PR against
  `feat/desktop-dock-mode`.

## Stop

All gates green, screenshot committed, PR open. Goal auto-clears.
