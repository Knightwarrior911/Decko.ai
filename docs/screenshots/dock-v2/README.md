# Dock mode v2 (SP7) walkthrough

`reflowed_side_by_side.png` — Decko docked to PowerPoint's right edge
WITH PowerPoint shrunk via `SetWindowPos` so Decko occupies dedicated
gutter space. No overlap. Restored on undock or shutdown via cached
original rect (`app/dock.py:_ORIGINAL_PPT_RECTS`).

## Behavior summary vs SP6

| Concern                            | SP6 (broken)                | SP7 (fixed)                                                   |
| ---------------------------------- | --------------------------- | ------------------------------------------------------------- |
| Decko on PPT minimize              | Hidden, no taskbar entry    | Stays visible; user controls own min/restore via taskbar      |
| Always-on-top                      | Forced (`on_top=True`)      | Off by default; Settings toggle                                |
| PPT real estate vs Decko           | Decko overlapped PPT slides | PPT shrinks via `SetWindowPos`; restored on undock/shutdown   |
| Slideshow (F5)                     | Decko hidden                | Decko z-order lowered behind slideshow; alt-tab back anytime  |
| Frameless taskbar entry            | Missing (tool-window style) | Forced `WS_EX_APPWINDOW` via `SetWindowLongPtrW`              |

## Regenerating

```
python tools/capture_dock_v2_screenshot.py
```
