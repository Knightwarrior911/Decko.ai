"""Inspect a form's controls and report Font.Name, Size, AutoSize, Width, Height."""
import time
import sys
from pathlib import Path
import win32com.client

CARRIER = Path(__file__).resolve().parent.parent / "PPT_AI_Editor.pptm"

app = win32com.client.DispatchEx("PowerPoint.Application")
app.Visible = True
pres = app.Presentations.Open(str(CARRIER), WithWindow=True)
time.sleep(2)
try:
    project = pres.VBProject
    for form_name in ("frmExport", "frmExecute", "frmImportSlides"):
        try:
            comp = project.VBComponents(form_name)
        except Exception as e:
            print(f"[{form_name}] not found: {e}")
            continue
        designer = comp.Designer
        try:
            print(f"\n=== {form_name} ===")
            print(f"  Form BackColor={designer.BackColor}")
            for prop in ("Zoom", "Width", "Height", "InsideWidth", "InsideHeight"):
                try:
                    print(f"  Form {prop}={getattr(designer, prop)}")
                except Exception as e:
                    print(f"  Form {prop}=<err {e}>")
        except Exception as e:
            print(f"  designer access failed: {e}")
        for c in designer.Controls:
            try:
                fname = c.Font.Name
                fsize = c.Font.Size
            except Exception as e:
                fname, fsize = f"<err {e}>", "?"
            try:
                bg = c.BackColor
            except Exception:
                bg = "?"
            try:
                fg = c.ForeColor
            except Exception:
                fg = "?"
            try:
                autosize = getattr(c, "AutoSize", "n/a")
            except Exception:
                autosize = "?"
            print(f"  {c.Name:20} Top={c.Top:>4} Left={c.Left:>4} W={c.Width:>4} H={c.Height:>4} | "
                  f"font={fname!r}/{fsize} bg={bg} fg={fg} auto={autosize}")
finally:
    pres.Close()
    app.Quit()
