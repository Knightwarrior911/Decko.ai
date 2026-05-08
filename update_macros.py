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

            # Ensure Microsoft Scripting Runtime reference is present.
            # modJSON uses early-bound Dictionary (New Dictionary); without this
            # reference the module fails to compile and every call to it raises
            # RPC_E_SERVERFAULT via COM.
            _scripting_guid = "{420B2830-E718-11CF-893D-00A0C9054228}"
            _has_scripting = any(
                getattr(r, "Guid", "") == _scripting_guid
                for r in project.References
            )
            if not _has_scripting:
                try:
                    project.References.AddFromGuid(_scripting_guid, 1, 0)
                    print("  [ref] Added Microsoft Scripting Runtime")
                except Exception as ref_err:
                    print(f"  [warn] Could not add Scripting Runtime reference: {ref_err}")

            components = project.VBComponents

            for src in sources:
                name = src.stem
                # Remove existing component by name (if any).
                # Re-enumerate fresh each time: the COM collection reference can
                # become stale after a prior Import(), causing 0x80070006
                # (ERROR_INVALID_HANDLE) on the next .Name access.
                to_remove = None
                for comp in components:
                    try:
                        if comp.Name == name:
                            to_remove = comp
                            break
                    except Exception:
                        pass
                if to_remove is not None:
                    print(f"  [remove] {name}")
                    components.Remove(to_remove)
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
