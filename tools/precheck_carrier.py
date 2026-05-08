"""Verify the carrier's VBProject compiles after a sync.

Run: python tools/precheck_carrier.py

Exits 0 if `BuildSnapshotJson` is callable on a fresh test deck; exits
non-zero with a diagnostic message if the carrier has a compile error
or the macro is missing.

Usage in agent workflow: always run this AFTER `python update_macros.py`
and BEFORE `python tests/run_smoke.py`. If it fails, read the .bas file
you most recently edited and fix the compile error. Do not proceed to
smoke.
"""
import os
import shutil
import sys
import tempfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
SAMPLE_DECK = REPO_ROOT / "test_decks" / "smoke_3slide.pptx"


def main() -> int:
    if not CARRIER.exists():
        print(f"ERROR: carrier not found at {CARRIER}")
        return 2
    if not SAMPLE_DECK.exists():
        print(f"ERROR: sample deck not found at {SAMPLE_DECK}")
        return 2

    try:
        import win32com.client
    except ImportError:
        print("ERROR: pywin32 not installed")
        return 2

    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_precheck_"))
    deck_copy = tmpdir / SAMPLE_DECK.name
    shutil.copy2(SAMPLE_DECK, deck_copy)

    app = win32com.client.DispatchEx("PowerPoint.Application")
    deck = None
    carrier = None
    try:
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
        deck = app.Presentations.Open(str(deck_copy), WithWindow=True)
        deck.Windows(1).Activate()

        try:
            json_text = app.Run("PPT_AI_Editor!BuildSnapshotJson")
        except Exception as e:
            print(f"FAIL: BuildSnapshotJson call raised: {e}")
            print("  This usually means the carrier has a compile error.")
            print("  Read the .bas you most recently edited and fix the error.")
            return 1

        if not isinstance(json_text, str) or not json_text.strip().startswith("{"):
            print(f"FAIL: BuildSnapshotJson returned non-JSON: {json_text!r}")
            return 1

        print("OK: carrier compiles and BuildSnapshotJson returns valid JSON")
        return 0
    finally:
        for p in (deck, carrier):
            if p is not None:
                try:
                    p.Saved = True
                except Exception:
                    pass
                try:
                    p.Close()
                except Exception:
                    pass
        try:
            app.Quit()
        except Exception:
            pass
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
