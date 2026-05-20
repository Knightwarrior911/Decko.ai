# Decko Desktop — Dock Mode v2 (SP7)

**Status:** APPROVED — 2026-05-20
**Branch:** `feat/desktop-dock-v2` off `feat/desktop-dock-mode` (HEAD `14ea0a2`)
**Stacks on:** PR #2 (SP6).

## 1. Problem

SP6 shipped a floating-pinned sidebar that snaps to PowerPoint's right
edge, but three behaviors break the task-pane illusion:

- **Vanish-on-PPT-minimize.** SP6 minimized Decko when PowerPoint
  minimized. The frameless Decko window had no taskbar entry, so users
  had no way back. Cardinal sin for a "task pane" UX.
- **`on_top=True` by default.** Decko hovered above every other window,
  blocking alt-tab to other apps.
- **No PowerPoint reflow.** Decko docked flush to PPT's right edge but
  PPT didn't yield space — when PPT spanned the screen, Decko overlapped
  slides. Real task panes (Claude for PowerPoint, native Office task
  panes) shrink the document area.

## 2. Goal

Make dock mode feel like a real Office task pane:

- Decko always reachable (taskbar entry + own lifecycle).
- Decko stops hovering above unrelated windows.
- PowerPoint yields width to Decko; restores on undock.

## 3. Non-goals

- True Office Add-in (VSTO/office-js manifest) — out of scope. SP7
  stays in the pywebview + win32 stack.
- Engine changes. `src/*.bas` frozen.
- New actions.
- Mac, Linux, cloud sync, licensing.
- Replacing pywebview.

## 4. Scope — exact behavior changes

### 4.1 Stay visible on PPT minimize

Drop the `minimize` branch in `app/main.py:_dock_event_handler`. Dock
loop continues to receive `EVENT_SYSTEM_MINIMIZESTART` but no longer
calls `webview.windows[0].minimize()`. Decko remains in place.

### 4.2 Real taskbar entry

After `webview.start(...)` callback fires, force `WS_EX_APPWINDOW` on
the Decko window (so it appears in the taskbar even when frameless):

```python
GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080
SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED | SWP_NOACTIVATE
```

Apply via `SetWindowLongPtrW` (fall back to `SetWindowLongW` on 32-bit
or when ctypes lacks the Ptr variant). To force Windows to re-register
the taskbar entry, briefly `ShowWindow(SW_HIDE)` then `ShowWindow(
SW_SHOWNORMAL)`.

### 4.3 Drop `on_top=True` default

Remove `on_top=True` from `webview.create_window(...)` in dock mode.
New `Settings.decko_on_top: bool = False` controls it via a "Keep Decko
on top" checkbox in the Settings dialog. When toggled true at runtime,
apply via `SetWindowPos(hwnd, HWND_TOPMOST, ...)`; when false,
`HWND_NOTOPMOST`.

### 4.4 PowerPoint reflow

When the dock loop emits `move_resize`/`restore` AND
`Settings.resize_ppt_for_dock=True`, also shrink PowerPoint's window to
free Decko's width:

```
new_ppt_right = decko_left
SetWindowPos(ppt_hwnd, 0,
             ppt.left, ppt.top,
             new_ppt_right - ppt.left, ppt.height,
             SWP_NOZORDER | SWP_NOACTIVATE)
```

If PPT already fits on the monitor with room for Decko (PPT.right +
decko_width ≤ monitor.right), do NOT shrink — leave PPT alone and dock
Decko in the existing gutter.

Cache the pre-shrink PPT rect in `app.dock` (per-hwnd dict) so we can
restore it on undock/shutdown.

### 4.5 Restore PPT on undock/close

`app/dock.py:restore_ppt_window(ppt_hwnd)` reapplies the cached
original rect via `SetWindowPos`. Called from:

- `Api.set_dock_mode(False)` (user clicks undock pin or Settings
  checkbox).
- `Api.shutdown()` (Decko window closing).
- `app.dock.stop_dock_loop(...)` (defensive).

### 4.6 Frameless controls fix

`Api.window_minimize` now calls `ShowWindow(hwnd, SW_MINIMIZE)` via
ctypes (bypasses pywebview's `Window.minimize` which doesn't reliably
register a taskbar restore on frameless windows). After 4.2's
WS_EX_APPWINDOW fix, taskbar click → restore works natively.

### 4.7 No slideshow auto-hide

Drop the `webview.windows[0].hide()` call on `slideshow_enter`. Instead
lower z-order behind the slideshow window: `SetWindowPos(decko_hwnd,
slideshow_hwnd, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE)`.
User alt-tabs back to Decko whenever they want.

## 5. Lifecycle table (v2 — replaces SP6 §5)

| Event                              | Decko action                                          |
| ---------------------------------- | ----------------------------------------------------- |
| Initial launch + PPT open          | Snap to right edge; reflow PPT to free Decko's width. |
| Initial launch + no PPT open       | Center; "No PowerPoint deck open" hero.               |
| PPT moves                          | Recompute dock rect; reflow PPT if needed.            |
| PPT resizes                        | Re-dock; PPT reflow respects Settings checkbox.       |
| PPT minimized                      | Decko stays put.                                      |
| PPT restored from minimize         | Re-snap; reflow.                                      |
| PPT closed                         | Decko stays; hero state.                              |
| PPT slideshow (F5) starts          | Z-order: Decko goes behind slideshow window.          |
| PPT slideshow ends                 | Re-snap; restore z-order.                             |
| User clicks Decko min button       | Minimize via SW_MINIMIZE; taskbar restore works.      |
| User clicks Decko close button     | Restore PPT to original width; destroy window.        |
| User toggles dock_mode OFF         | Restore PPT to original width; switch to framed SP5 layout. |
| User toggles dock_mode ON          | Cache PPT original rect; snap + reflow.               |
| User toggles "Keep Decko on top"   | HWND_TOPMOST ↔ HWND_NOTOPMOST live.                   |
| User toggles "Resize PPT for dock" | If true → reflow now; if false → restore PPT, dock as gutter overlay. |

## 6. Files touched

- `app/dock.py` — add `reflow_ppt_window` + `restore_ppt_window`,
  cached rect dict, drop slideshow/minimize hide paths.
- `app/config.py` — `+decko_on_top: bool = False`,
  `+resize_ppt_for_dock: bool = True`. Persisted.
- `app/main.py` — drop on_top, force WS_EX_APPWINDOW, ctypes-based
  minimize, wire new fields, restore PPT on undock/shutdown.
- `app/web/index.html` — 2 new Settings checkboxes.
- `app/web/app.js` — wire checkboxes through `api.save_settings`.
- `tests/run_smoke_ui_dock.py` — assert new fields + callables.
- `docs/screenshots/dock-v2/` — 1 PIL composite (reflowed side-by-side).

## 7. Metric

- `tests/run_smoke_ui_dock.py` updated + PASS.
- `tests/run_smoke_ui_polish.py` PASS.
- `tests/run_smoke_app.py` PASS (all 7 gates).
- `tests/run_smoke.py` PASS.
- `tests/test_guidance_coverage.py` + `test_guidance_doc_sync.py` PASS.

## 8. Risks

- Shrinking PPT via SetWindowPos triggers PPT's own size-recalc; if
  Decko's hook fires recursively on the resulting LOCATIONCHANGE,
  loops can race. Mitigation: gate the reflow path with a "currently
  reflowing" flag and skip reflow if `event.source == "self_reflow"`.
- Cached original PPT rect goes stale if user manually resizes PPT
  while docked. Mitigation: update the cached rect whenever the user
  resizes PPT (compare LOCATIONCHANGE delta vs our last-reflowed
  target; if it differs significantly, refresh the cache).
- `SetWindowLongPtrW` may not be in `ctypes.windll.user32` on older
  Pythons; fall back to `SetWindowLongW` (still works for EXSTYLE on
  64-bit since EXSTYLE fits in a LONG).
- WS_EX_APPWINDOW + frameless can cause flicker during the
  Hide→Show toggle. Mitigation: use `SWP_NOACTIVATE` so focus stays
  with PowerPoint during the cycle.

## 9. Out of scope (firm)

Office Add-in, Mac, engine, new actions, licensing, cloud sync,
replacing pywebview.
