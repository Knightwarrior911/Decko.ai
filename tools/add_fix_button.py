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
    ' Delegates to modVerify.CopyWarningsPromptToClipboard so the logic is
    ' testable from Application.Run and reusable elsewhere.
    Dim n As Long
    n = modVerify.CopyWarningsPromptToClipboard()
    If n = 0 Then
        lblStatus.Caption = "No warnings to copy. Either run Apply first, or the deck is clean."
    Else
        lblStatus.Caption = n & " warning(s) copied to clipboard as LLM prompt. " & _
                            "Paste into your chat and ask the model to fix."
    End If
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

    # Check if button already exists
    existing = None
    for ctrl in designer.Controls:
        if ctrl.Name == BUTTON_NAME:
            existing = ctrl
            break

    if existing is not None:
        print(f"Button '{BUTTON_NAME}' already present at "
              f"({existing.Left}, {existing.Top}) — skipping Add.")
    else:
        print(f"Adding '{BUTTON_NAME}' to {FORM_NAME}...")
        # Position to right of btnCancel. Find btnCancel for reference.
        ref_btn = None
        for ctrl in designer.Controls:
            if ctrl.Name == "btnCancel":
                ref_btn = ctrl
                break

        btn = designer.Controls.Add("Forms.CommandButton.1", BUTTON_NAME)
        btn.Caption = "Fix This"
        btn.Width = 90
        btn.Height = 24
        if ref_btn is not None:
            # Place to the left of Cancel
            btn.Top = ref_btn.Top
            btn.Left = max(10, ref_btn.Left - 100)
        else:
            btn.Left = 600
            btn.Top = 600
        # Make it stand out
        btn.BackColor = 0xCCCCCC
        print(f"  placed at ({btn.Left}, {btn.Top}) size {btn.Width}x{btn.Height}")

    # Inject (or replace) the click handler in the form's CodeModule
    code = form_comp.CodeModule
    # Look for existing btnFixThis_Click sub and remove it
    try:
        start = code.ProcStartLine("btnFixThis_Click", 0)
        count = code.ProcCountLines("btnFixThis_Click", 0)
        if start > 0:
            print(f"Removing existing handler at lines {start}..{start + count - 1}")
            code.DeleteLines(start, count)
    except Exception:
        pass  # handler not present yet — fine

    print("Appending click handler...")
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
