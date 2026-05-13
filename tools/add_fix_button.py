"""
add_fix_button.py — one-time installer for the "Fix This" button on frmExecute.

Adds a CommandButton named btnFixThis to the form's layout (via VBE Designer)
and injects the Click handler code (via CodeModule). Saves the carrier, then
exports frmExecute back to src/ so future update_macros.py syncs preserve the
button.

Run: python tools/add_fix_button.py
"""

import sys
from pathlib import Path

import win32com.client
import pythoncom

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
SRC_DIR = REPO_ROOT / "src"
FORM_NAME = "frmExecute"
BUTTON_NAME = "btnFixThis"

CLICK_HANDLER = r'''
Private Sub btnFixThis_Click()
    ' Post-Apply quality warnings -> clipboard. See modVerify.
    Dim n As Long
    n = modVerify.CopyWarningsPromptToClipboard()
    If n = 0 Then
        lblStatus.Caption = "No warnings to copy. Either run Apply first, or the deck is clean."
    Else
        lblStatus.Caption = n & " warning(s) copied to clipboard as LLM prompt. " & _
                            "Paste into your chat and ask the model to fix."
    End If
End Sub

Private Sub btnFixErrors_Click()
    ' Pre-Apply action validation errors -> clipboard. Reads current JSON
    ' (textbox or loaded file), runs PreviewValidate on each action, builds
    ' an LLM-ready prompt with errors + canonical guidance for each failing
    ' action type. Eliminates the user having to read INVALID lines and
    ' hand-type a correction request.
    Dim json As String: json = CurrentJson()
    If Len(json) = 0 Then
        lblStatus.Caption = "No actions JSON to validate. Paste or Load a batch first."
        Exit Sub
    End If
    Dim prompt As String
    prompt = modExecuteInstructions.BuildErrorFixPrompt(json)
    If Len(prompt) = 0 Then
        lblStatus.Caption = "All actions are valid — nothing to fix."
        Exit Sub
    End If
    Dim dobj As MSForms.DataObject
    Set dobj = New MSForms.DataObject
    dobj.SetText prompt
    dobj.PutInClipboard
    lblStatus.Caption = "Error-fix prompt copied to clipboard. Paste into your LLM chat."
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

    # Locate btnCancel for relative placement
    ref_btn = None
    for ctrl in designer.Controls:
        if ctrl.Name == "btnCancel":
            ref_btn = ctrl
            break

    def ensure_button(name: str, caption: str, left: float, top: float):
        existing = None
        for c in designer.Controls:
            if c.Name == name:
                existing = c
                break
        if existing is not None:
            # Reposition (in case prior run stacked them)
            existing.Left = left
            existing.Top = top
            existing.Width = 90
            existing.Height = 24
            existing.Caption = caption
            print(f"Button '{name}' repositioned to ({left}, {top}).")
            return existing
        print(f"Adding '{name}' to {FORM_NAME}...")
        b = designer.Controls.Add("Forms.CommandButton.1", name)
        b.Caption = caption
        b.Width = 90
        b.Height = 24
        b.Left = left
        b.Top = top
        b.BackColor = 0xCCCCCC
        print(f"  placed at ({left}, {top}) size 90x24")
        return b

    # Place both Fix buttons on the bottom row, FAR LEFT. Move Apply +
    # Cancel to the FAR RIGHT of the same row so they don't get covered.
    # Form is ~611pt wide; buttons are 80-90pt wide each.
    base_top = ref_btn.Top if ref_btn is not None else 384
    ensure_button("btnFixErrors", "Fix Errors", 10, base_top)
    ensure_button(BUTTON_NAME, "Fix This", 110, base_top)
    # Reposition Apply + Cancel to the right so they remain visible.
    for c in designer.Controls:
        if c.Name == "btnApply":
            c.Left = 410
            c.Top = base_top
        elif c.Name == "btnCancel":
            c.Left = 500
            c.Top = base_top
    print(f"  repositioned btnApply -> (410, {base_top}), btnCancel -> (500, {base_top})")

    # Inject (or replace) the click handlers in the form's CodeModule.
    code = form_comp.CodeModule
    for proc in ("btnFixThis_Click", "btnFixErrors_Click"):
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
    print("\nDone. Open Decko.ai > Execute Instructions form — 'Fix This' button now present.")
    print("After running this once, future `python update_macros.py` syncs preserve the button.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
