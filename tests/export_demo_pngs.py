"""Export each slide of the test verification deck as a PNG so the user
can visually see what each planted quality issue looks like."""

import os
from pathlib import Path
import win32com.client
import pythoncom

REPO_ROOT = Path(__file__).resolve().parent.parent
TEST_DECK = REPO_ROOT / "tests" / "test_verify_deck.pptx"
OUT_DIR = REPO_ROOT / "tests" / "demo_pngs"


def main() -> int:
    if not TEST_DECK.exists():
        print(f"ERROR: {TEST_DECK} not found. Run test_verify_loop.py first.")
        return 1

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    # Wipe old PNGs
    for f in OUT_DIR.glob("*.png"):
        f.unlink()

    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True
    while app.Presentations.Count > 0:
        try:
            app.Presentations(1).Close()
        except Exception:
            break

    pres = app.Presentations.Open(str(TEST_DECK))
    print(f"Exporting {pres.Slides.Count} slides as PNG...")

    for i, sl in enumerate(pres.Slides, start=1):
        out_path = OUT_DIR / f"slide_{i:02d}.png"
        sl.Export(str(out_path), "PNG", 1280, 720)
        print(f"  -> {out_path}")

    pres.Close()
    print(f"\nAll PNGs written to: {OUT_DIR}")
    return 0


if __name__ == "__main__":
    main()
