"""Dock mode gate (SP6).

Deterministic checks that the SP6 surface exists:

1. `app.dock` module imports cleanly and exposes the four callables.
2. `Settings.dock_mode` field exists and defaults to True.
3. `compute_dock_rect` math is sane for the null-hwnd path (centers a
   380×600 default rect on a positive-coordinate monitor).
4. `app/web/index.html` exposes the undock pin + window control buttons +
   drag area.
5. `app/web/app.js` calls the new Api methods (`set_dock_mode`,
   `window_minimize`, `window_close`).
6. `app/web/app.css` carries dock-aware rules (`[data-dock=` selector +
   `.winctl` class).

Does NOT spawn pywebview, does NOT touch COM. Safe in CI.
"""
from __future__ import annotations

import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
WEB = REPO / "app" / "web"
HTML = WEB / "index.html"
CSS = WEB / "app.css"
JS = WEB / "app.js"


def _read(p: Path) -> str:
    return p.read_text(encoding="utf-8") if p.exists() else ""


def check_module() -> list[str]:
    fails: list[str] = []
    try:
        sys.path.insert(0, str(REPO))
        from app import dock  # noqa: F401  (import-time check)
    except Exception as e:  # noqa: BLE001
        return [f"`from app import dock` failed: {type(e).__name__}: {e}"]
    for name in ("find_ppt_window", "compute_dock_rect",
                 "start_dock_loop", "stop_dock_loop"):
        if not callable(getattr(dock, name, None)):
            fails.append(f"app.dock.{name} missing or not callable")
    # compute_dock_rect math for the null-hwnd path.
    try:
        rect = dock.compute_dock_rect(0, 380, 600)
    except Exception as e:  # noqa: BLE001
        fails.append(f"compute_dock_rect(0, 380, 600) raised "
                     f"{type(e).__name__}: {e}")
        return fails
    if not (isinstance(rect, tuple) and len(rect) == 4 and
            all(isinstance(v, int) for v in rect)):
        fails.append(f"compute_dock_rect returned {rect!r}, expected "
                     "tuple of 4 ints")
        return fails
    x, y, w, h = rect
    if w != 380 or h < 600:
        fails.append(f"compute_dock_rect width/height wrong: w={w}, h={h}")
    return fails


def check_settings() -> list[str]:
    fails: list[str] = []
    try:
        sys.path.insert(0, str(REPO))
        from app.config import Settings  # type: ignore
    except Exception as e:  # noqa: BLE001
        return [f"`from app.config import Settings` failed: "
                f"{type(e).__name__}: {e}"]
    inst = Settings()
    if not hasattr(inst, "dock_mode"):
        fails.append("Settings.dock_mode field missing")
    elif inst.dock_mode is not True:
        fails.append(f"Settings.dock_mode default is {inst.dock_mode!r}, "
                     "expected True")
    return fails


def check_frontend() -> list[str]:
    fails: list[str] = []
    html = _read(HTML)
    js = _read(JS)
    css = _read(CSS)

    for needle in ('id="undockBtn"', 'id="winMinBtn"',
                   'id="winCloseBtn"', 'class="dragArea"'):
        if needle not in html:
            fails.append(f"index.html missing {needle!r}")

    for needle in ("set_dock_mode", "window_minimize", "window_close"):
        if needle not in js:
            fails.append(f"app.js missing reference to api.{needle}")

    for needle in ("[data-dock=", ".winctl"):
        if needle not in css:
            fails.append(f"app.css missing {needle!r}")

    return fails


def main() -> int:
    fails = check_module() + check_settings() + check_frontend()
    if fails:
        print(f"run_smoke_ui_dock: FAIL ({len(fails)} issue(s))")
        for f in fails:
            print(f"  - {f}")
        return 1
    print("run_smoke_ui_dock: PASS — dock surface + Settings + frontend ok.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
