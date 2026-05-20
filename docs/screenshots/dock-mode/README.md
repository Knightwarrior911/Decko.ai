# Dock-mode walkthrough (SP6)

Three illustrative shots covering the lifecycle states defined in
`docs/superpowers/specs/2026-05-20-desktop-dock-mode.md` §5.

| File                       | Lifecycle state                                                        |
| -------------------------- | ---------------------------------------------------------------------- |
| `docked_with_ppt.png`      | Dock mode ON — Decko snapped flush to PowerPoint's right edge          |
| `undocked_detached.png`    | Dock mode OFF — SP5 framed layout (free-floating window)               |
| `slideshow_hidden.png`     | PowerPoint slideshow (F5) — dock loop hides Decko via `window.hide()` |

## How these are made

`tools/capture_dock_screenshots.py` composes the docked-vs-PPT picture
from a PIL-drawn PowerPoint stand-in plus a cropped slice of the proven
SP5 hero render (`docs/screenshots/consumer-polish/after_02_idle_empty.png`).
We render the visual relationship offline rather than spawning two live
windows because pywebview's main loop + a Tk fake-PPT mainloop race on
the same thread on Windows.

The dock loop itself (`app/dock.py`) is exercised live whenever Decko
runs against a real PowerPoint window — its behavior is verified
deterministically by `tests/run_smoke_ui_dock.py` (callable surface +
math), and observed in person against a real PPT install.
