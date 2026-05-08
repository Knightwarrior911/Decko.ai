"""One-time bootstrap: create an empty PPT_AI_Editor.pptm file.

Run after cloning the repo. Idempotent: skips if file already exists.
"""
import os
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"


def main() -> int:
    if CARRIER.exists():
        print(f"[skip] {CARRIER} already exists")
        return 0

    try:
        import win32com.client
    except ImportError:
        print("ERROR: pywin32 not installed. Run: pip install -r requirements.txt")
        return 1

    print(f"[build] Creating {CARRIER}")
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Add(WithWindow=False)
        # ppSaveAsOpenXMLPresentationMacroEnabled = 25
        pres.SaveAs(str(CARRIER), 25)
        pres.Close()
    finally:
        app.Quit()
    print("[done]")
    return 0


if __name__ == "__main__":
    sys.exit(main())
