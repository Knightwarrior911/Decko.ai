"""End-to-end smoke tests driven via PowerPoint COM.

Run: python tests/run_smoke.py

Each test opens a fresh copy of a test deck plus the carrier, calls VBA
functions/subs via Application.Run, asserts on returned values, and tears
down. One failure prints the diff and exits non-zero.
"""
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
DECKS_DIR = REPO_ROOT / "test_decks"


def open_app():
    import win32com.client
    return win32com.client.DispatchEx("PowerPoint.Application")


def open_pair(app, deck_name: str):
    """Open a copy of the test deck + the carrier. Returns (deck, carrier)."""
    src = DECKS_DIR / deck_name
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_"))
    deck_copy = tmpdir / src.name
    shutil.copy2(src, deck_copy)

    app.Visible = True
    deck = app.Presentations.Open(str(deck_copy), WithWindow=True)
    carrier = app.Presentations.Open(str(CARRIER), WithWindow=False)
    deck.Windows(1).Activate()
    return deck, carrier, tmpdir


def teardown(app, *presentations, tmpdir=None):
    for p in presentations:
        try:
            p.Close()
        except Exception:
            pass
    try:
        app.Quit()
    except Exception:
        pass
    time.sleep(0.5)
    if tmpdir and tmpdir.exists():
        shutil.rmtree(tmpdir, ignore_errors=True)


def assert_eq(actual, expected, label):
    if actual != expected:
        print(f"FAIL [{label}]")
        print(f"  expected: {expected!r}")
        print(f"  actual:   {actual!r}")
        sys.exit(1)
    print(f"  ok  [{label}]")


def test_snapshot_smoke_3slide():
    print("test_snapshot_smoke_3slide")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "smoke_3slide.pptx")
    try:
        json_text = app.Run("BuildSnapshotJson")
        snap = json.loads(json_text)
        assert_eq(len(snap["slides"]), 3, "slide count")
        assert_eq(snap["slides"][0]["slide_number"], 1, "slide 1 number")
        # First slide title should match what we set
        first_slide_texts = [s.get("text", "") for s in snap["slides"][0]["shapes"]]
        assert "Q3 Results" in first_slide_texts, f"title text not found in {first_slide_texts}"
        print("  ok  [title text present]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def main() -> int:
    test_snapshot_smoke_3slide()
    print("\nall tests passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
