"""SP9 — verify partial resize_shape/move_shape applies cleanly.

Creates a blank deck via COM, adds an autoshape with a known initial
rect, then runs four ExecuteFromString batches:

1. resize_shape with width only → height should stay.
2. resize_shape with height only → width should stay (from step 1).
3. move_shape with left only → top should stay.
4. move_shape with top only → left should stay (from step 3).

After each batch, reads the shape's actual rect and asserts the
preserved dimension didn't change.
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
CARRIER = REPO / "PPT_AI_Editor.pptm"

import win32com.client


def main() -> int:
    if not CARRIER.exists():
        print(f"FAIL: carrier missing at {CARRIER}")
        return 1
    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.Visible = True
    pres = app.Presentations.Open(str(CARRIER), WithWindow=True)
    try:
        # Clean deck: one blank slide we add a rect onto.
        if pres.Slides.Count == 0:
            pres.Slides.Add(1, 12)  # ppLayoutBlank = 12
        sl = pres.Slides(1)
        sh = sl.Shapes.AddShape(1, 100, 100, 200, 150)  # msoShapeRectangle = 1
        sid = sh.Id
        sh.Name = "sp9_target"

        def fire(batch):
            return app.Run("'" + CARRIER.name + "'!ExecuteFromString",
                           json.dumps({"actions": batch, "verify_after": False}))

        # 1. resize width only.
        fire([{"type": "resize_shape", "slide": 1, "shape_id": sid, "width": 280}])
        if abs(sh.Width - 280) > 0.5:
            print(f"FAIL width-only: expected 280, got {sh.Width}")
            return 1
        if abs(sh.Height - 150) > 0.5:
            print(f"FAIL width-only: height changed to {sh.Height}")
            return 1
        print(f"  width-only ok: ({sh.Width:.0f}, {sh.Height:.0f})")

        # 2. resize height only.
        fire([{"type": "resize_shape", "slide": 1, "shape_id": sid, "height": 220}])
        if abs(sh.Width - 280) > 0.5 or abs(sh.Height - 220) > 0.5:
            print(f"FAIL height-only: got ({sh.Width}, {sh.Height})")
            return 1
        print(f"  height-only ok: ({sh.Width:.0f}, {sh.Height:.0f})")

        # 3. move left only.
        fire([{"type": "move_shape", "slide": 1, "shape_id": sid, "left": 36}])
        if abs(sh.Left - 36) > 0.5 or abs(sh.Top - 100) > 0.5:
            print(f"FAIL left-only: got L={sh.Left}, T={sh.Top}")
            return 1
        print(f"  left-only ok: L={sh.Left:.0f}, T={sh.Top:.0f}")

        # 4. move top only.
        fire([{"type": "move_shape", "slide": 1, "shape_id": sid, "top": 60}])
        if abs(sh.Left - 36) > 0.5 or abs(sh.Top - 60) > 0.5:
            print(f"FAIL top-only: got L={sh.Left}, T={sh.Top}")
            return 1
        print(f"  top-only ok: L={sh.Left:.0f}, T={sh.Top:.0f}")

        # 5. Negative-path: both missing should be rejected.
        try:
            r = fire([{"type": "resize_shape", "slide": 1, "shape_id": sid}])
            if "width or height" not in str(r):
                print(f"FAIL: empty resize should error, got: {r}")
                return 1
            print("  empty-resize rejected ok")
        except Exception as e:  # noqa: BLE001
            print(f"  empty-resize raised ok: {type(e).__name__}")

        print("\nrun_smoke_partial_geometry: PASS")
        return 0
    finally:
        try:
            pres.Saved = True
            pres.Close()
        except Exception:  # noqa: BLE001
            pass
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass


if __name__ == "__main__":
    sys.exit(main())
