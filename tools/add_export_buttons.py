"""
add_export_buttons.py — idempotent installer for two action-shortcut buttons
on frmExport: "Copy deck spec" and "Scan palette".

Mirrors tools/add_fix_button.py exactly: adds the CommandButtons via the VBE
Designer, injects the Click handlers via the form CodeModule (append/replace
only — the form's PromptTemplate / snapshot code is left untouched), saves
the carrier, then exports frmExport back to src/ so future
`python update_macros.py` syncs preserve the buttons.

These two buttons expose existing actions that previously required the user
to hand-paste JSON:
  - btnCopyDeckSpec  -> modActionsSpec.ExtractDeckSpecJson() -> clipboard
       (extract_spec normally only writes <deck>.spec.json — invisible)
  - btnScanPalette   -> runs {"actions":[{"type":"scan_palette"}]}
       (scan_palette already self-copies role-tagged JSON to the clipboard)

Run: python tools/add_export_buttons.py
"""

import sys
from pathlib import Path

import win32com.client
import pythoncom

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
SRC_DIR = REPO_ROOT / "src"
FORM_NAME = "frmExport"

# Append-only handlers. CopyToClipboard is an existing Private Sub in
# frmExport; ExtractDeckSpecJson / ExecuteFromString are Public in their
# modules. scan_palette is read-only and self-clipboards.
CLICK_HANDLER = r'''
Private Sub btnCopyDeckSpec_Click()
    On Error GoTo fail
    Dim js As String
    js = modActionsSpec.ExtractDeckSpecJson()
    CopyToClipboard js
    lblStatus.Caption = "Deck spec copied to clipboard (" & Len(js) & " chars)."
    Exit Sub
fail:
    lblStatus.Caption = "Extract spec failed: " & Err.Description
End Sub

Private Sub btnScanPalette_Click()
    On Error GoTo fail
    Dim r As String
    r = modExecuteInstructions.ExecuteFromString( _
        "{""actions"":[{""type"":""scan_palette""}]}")
    lblStatus.Caption = "Palette scanned -> clipboard + %TEMP%\decko_palette.json"
    Exit Sub
fail:
    lblStatus.Caption = "Scan palette failed: " & Err.Description
End Sub
'''


def main() -> int:
    if not CARRIER.exists():
        print(f"ERROR: {CARRIER} not found.")
        return 1

    pythoncom.CoInitialize()
    print("Launching PowerPoint...")
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True

    # Close any leftover presentations to avoid lock conflicts
    while app.Presentations.Count > 0:
        try:
            app.Presentations(1).Close()
        except Exception:
            break

    print(f"Opening carrier: {CARRIER}")
    carrier = app.Presentations.Open(str(CARRIER))

    vbproj = carrier.VBProject
    form_comp = vbproj.VBComponents(FORM_NAME)
    designer = form_comp.Designer

    def ensure_button(name, caption, left, top, width):
        existing = None
        for c in designer.Controls:
            if c.Name == name:
                existing = c
                break
        if existing is not None:
            existing.Left = left
            existing.Top = top
            existing.Width = width
            existing.Height = 24
            existing.Caption = caption
            print(f"Button '{name}' repositioned to ({left}, {top}).")
            return existing
        print(f"Adding '{name}' to {FORM_NAME}...")
        b = designer.Controls.Add("Forms.CommandButton.1", name)
        b.Caption = caption
        b.Left = left
        b.Top = top
        b.Width = width
        b.Height = 24
        b.BackColor = 0xCCCCCC
        print(f"  placed at ({left}, {top}) size {width}x24")
        return b

    def move_ctrl(name, left, top, width=None):
        for c in designer.Controls:
            if c.Name == name:
                c.Left = left
                c.Top = top
                if width is not None:
                    c.Width = width
                print(f"  {name} -> ({left}, {top})")
                return

    # Reflow the two bottom rows so the form (540x360, unchanged) holds the
    # two new buttons. Row A = primary copy actions; Row B = the two new
    # shortcuts + Close; lblStatus drops to its own line.
    move_ctrl("btnCopySnapshot",     12, 264, 120)
    move_ctrl("btnCopyWithTemplate", 140, 264, 200)
    move_ctrl("btnSaveTxt",          348, 264, 150)
    ensure_button("btnCopyDeckSpec", "Copy deck spec", 12, 298, 170)
    ensure_button("btnScanPalette",  "Scan palette",   190, 298, 150)
    move_ctrl("btnClose",            348, 298, 150)
    move_ctrl("lblStatus",           12, 330, 486)

    # Inject (or replace) the click handlers in the form's CodeModule.
    # Append-only: existing PromptTemplate / snapshot procs are untouched.
    code = form_comp.CodeModule
    for proc in ("btnCopyDeckSpec_Click", "btnScanPalette_Click"):
        try:
            start = code.ProcStartLine(proc, 0)
            n = code.ProcCountLines(proc, 0)
            if start > 0:
                print(f"Removing existing {proc} at lines {start}..{start + n - 1}")
                code.DeleteLines(start, n)
        except Exception:
            pass

    print("Appending click handlers...")
    code.AddFromString(CLICK_HANDLER.strip())

    print(f"Saving carrier: {CARRIER}")
    carrier.Save()

    print(f"Exporting {FORM_NAME} back to {SRC_DIR} ...")
    form_comp.Export(str(SRC_DIR / f"{FORM_NAME}.frm"))

    carrier.Close()
    print("\nDone. Decko.ai > Export Snapshot form now has "
          "'Copy deck spec' + 'Scan palette' buttons.")
    print("Future `python update_macros.py` syncs preserve them.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
