"""Export current UserForms from carrier .pptm to src/ (.frm + .frx).

Captures whatever state is currently in PPT_AI_Editor.pptm — useful when
the carrier was edited manually in the VBA IDE and we want to bring src/
back in sync without regenerating from build_forms.py.
"""
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
SRC = REPO_ROOT / "src"

FORMS = ("frmExport", "frmExecute", "frmImportSlides")


def main():
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.Visible = True
    pres = app.Presentations.Open(str(CARRIER), WithWindow=True)
    time.sleep(2)
    try:
        project = pres.VBProject
        for name in FORMS:
            try:
                comp = project.VBComponents(name)
            except Exception as e:
                print(f"[skip] {name}: {e}")
                continue
            out = SRC / f"{name}.frm"
            comp.Export(str(out))
            print(f"  exported {out}")
    finally:
        try:
            pres.Close()
        except Exception:
            pass
        app.Quit()
        time.sleep(1)


if __name__ == "__main__":
    sys.exit(main())
