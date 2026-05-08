"""Sync src/*.bas and src/*.frm files into PPT_AI_Editor.pptm.

Usage: python update_macros.py

Behavior:
  1. Open carrier headlessly.
  2. For each .bas / .frm in src/, remove existing module of the same
     name (if present) and import fresh.
  3. Save and close.

Requires: pywin32, and "Trust access to the VBA project object model"
enabled in PowerPoint Trust Center.
"""
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent
SRC_DIR = REPO_ROOT / "src"
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"

# vbext_ct_StdModule = 1, vbext_ct_MSForm = 3
MODULE_EXTS = {".bas", ".frm"}


def main() -> int:
    if not CARRIER.exists():
        print(f"ERROR: {CARRIER} not found. Run tools/build_carrier.py first.")
        return 1
    if not SRC_DIR.exists():
        print(f"ERROR: {SRC_DIR} not found.")
        return 1

    sources = sorted(p for p in SRC_DIR.iterdir() if p.suffix.lower() in MODULE_EXTS)
    if not sources:
        print(f"ERROR: no .bas / .frm files in {SRC_DIR}")
        return 1

    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(CARRIER), WithWindow=False)
        try:
            project = pres.VBProject
            components = project.VBComponents

            for src in sources:
                name = src.stem
                # Remove existing component by name (if any)
                for comp in list(components):
                    if comp.Name == name:
                        print(f"  [remove] {name}")
                        components.Remove(comp)
                        break
                print(f"  [import] {src.name}")
                components.Import(str(src))

            pres.Save()
        finally:
            pres.Close()
    finally:
        app.Quit()
        time.sleep(0.5)
    print("[done]")
    return 0


if __name__ == "__main__":
    sys.exit(main())
