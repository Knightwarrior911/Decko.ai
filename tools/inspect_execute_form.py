"""Dump every control on frmExecute with its position so we can see what's there."""
import win32com.client, pythoncom
from pathlib import Path
REPO = Path(__file__).resolve().parent.parent
pythoncom.CoInitialize()
app = win32com.client.Dispatch("PowerPoint.Application")
app.Visible = True
while app.Presentations.Count > 0:
    try: app.Presentations(1).Close()
    except: break
p = app.Presentations.Open(str(REPO / "PPT_AI_Editor.pptm"))
comp = p.VBProject.VBComponents("frmExecute")
des = comp.Designer
prop = comp.Properties
print(f"Form properties:")
for prop_name in ("Width", "Height", "ScrollHeight", "ScrollWidth"):
    try:
        print(f"  {prop_name} = {prop(prop_name).Value}")
    except Exception as e:
        print(f"  {prop_name} = (err: {e})")
print(f"{'Name':<20} {'Caption':<25} {'Left':>8} {'Top':>8} {'W':>6} {'H':>6} {'Visible':>8}")
for c in des.Controls:
    cap = ""
    try: cap = c.Caption
    except: pass
    print(f"{c.Name:<20} {cap:<25} {c.Left:>8} {c.Top:>8} {c.Width:>6} {c.Height:>6} {str(c.Visible):>8}")
p.Close()
