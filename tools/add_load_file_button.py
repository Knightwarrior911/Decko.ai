"""Surgically upgrade frmExecute: bigger input box + a 'Load from file' button.

- Enlarges txtInstructions and repositions the controls below it.
- Adds btnLoadFile (Office FilePicker -> reads file -> holds raw JSON in
  mLoadedJson, which bypasses the textbox so large/long-line input is not
  mangled by MSForms' paste path).
- Replaces frmExecute's code module with the updated version.
- Re-exports src/frmExecute.frm + .frx and fixes the .frm dimensions/caption.

Only touches frmExecute. Run: python tools/add_load_file_button.py
"""
import re
import sys
import time
from pathlib import Path

import win32com.client

ROOT = Path(__file__).resolve().parent.parent
CARRIER = ROOT / "PPT_AI_Editor.pptm"
SRC = ROOT / "src"
FRM = SRC / "frmExecute.frm"

CMDBTN = "Forms.CommandButton.1"
WHITE = 16777215
BLACK = 0
DARK_BG = 0
DARK_BTN = 51 + 51 * 256 + 51 * 65536      # #333333
MID_BTN = 87 + 87 * 256 + 87 * 65536       # #575757
FONT_NAME = "Cascadia Code"
FONT_SIZE = 10

# Outer form size (points) -> .frm ClientWidth/Height twips = pt * 20
FORM_W, FORM_H = 600, 520
CAPTION = "Decko.ai • Execute Instructions"

# New control layout (points)
LAYOUT = {
    "txtInstructions": dict(Top=12, Left=12, Width=570, Height=160),
    "btnParse":        dict(Top=180, Left=12, Width=80, Height=24),
    "btnLoadFile":     dict(Top=180, Left=100, Width=150, Height=24),
    "lstActions":      dict(Top=214, Left=12, Width=570, Height=160),
    "btnApply":        dict(Top=384, Left=12, Width=80, Height=24),
    "btnCancel":       dict(Top=384, Left=102, Width=80, Height=24),
    "lblStatus":       dict(Top=416, Left=12, Width=570, Height=60),
}

FRM_EXECUTE_CODE = '''\
Option Explicit

Private mParsed As Object
Private mValid() As Boolean
' When the user loads actions from a file, the full uncorrupted JSON is held
' here and takes precedence over the textbox. MSForms textboxes can mangle
' large pastes (whitespace injected into numbers/keys); the file path avoids
' the textbox entirely. Cleared the moment the user types in the textbox.
Private mLoadedJson As String

Private Sub UserForm_Initialize()
    lblStatus.Caption = ""
    mLoadedJson = ""
End Sub

' The JSON to parse/apply: a file loaded via btnLoadFile if one is held,
' otherwise the textbox contents.
Private Function CurrentJson() As String
    If Len(mLoadedJson) > 0 Then
        CurrentJson = mLoadedJson
    Else
        CurrentJson = txtInstructions.Text
    End If
End Function

Private Sub txtInstructions_Change()
    ' Hand-editing the textbox supersedes any previously loaded file.
    mLoadedJson = ""
End Sub

Private Sub btnLoadFile_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select an actions JSON file"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "JSON / JSONL / text", "*.json;*.jsonl;*.txt"
    fd.Filters.Add "All files", "*.*"
    If fd.Show <> -1 Then Exit Sub

    Dim path As String: path = fd.SelectedItems(1)
    Dim s As String
    On Error GoTo IOFail
    Dim fnum As Integer: fnum = FreeFile
    Open path For Input As #fnum
    If LOF(fnum) > 0 Then s = Input$(LOF(fnum), fnum)
    Close #fnum
    On Error GoTo 0

    txtInstructions.Text = s          ' raises txtInstructions_Change -> clears mLoadedJson
    mLoadedJson = s                   ' ...so set it AFTER the assignment above
    lstActions.Clear
    Set mParsed = Nothing
    lblStatus.Caption = "Loaded " & Len(s) & " chars from " & path & " -- click Parse."
    Exit Sub
IOFail:
    On Error GoTo 0
    lblStatus.Caption = "Could not read file: " & Err.Description
End Sub

Private Sub btnParse_Click()
    lblStatus.Caption = ""
    lstActions.Clear

    On Error Resume Next
    Set mParsed = modJSON.ParseJson(modExecuteInstructions.SanitizeJsonInput(CurrentJson()))
    If Err.Number <> 0 Then
        lblStatus.Caption = "Invalid JSON: " & Err.Description & _
            "  (large textbox pastes can corrupt -- use 'Load from file' for big batches)"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    If Not mParsed.Exists("actions") Then
        lblStatus.Caption = "Missing top-level 'actions' array."
        Exit Sub
    End If

    Dim actions As Object: Set actions = mParsed("actions")
    ReDim mValid(1 To actions.Count)

    Dim i As Long, anyValid As Boolean: anyValid = False
    For i = 1 To actions.Count
        Dim act As Object: Set act = actions(i)
        Dim reason As String
        reason = modExecuteInstructions.PreviewValidate(act)
        mValid(i) = (Len(reason) = 0)
        If mValid(i) Then anyValid = True

        Dim row As String
        row = i & ". " & GetStrSafe(act, "type")
        If act.Exists("slide") Then row = row & " | slide=" & act("slide")
        If act.Exists("shape_id") Then row = row & " | shape_id=" & act("shape_id")
        If mValid(i) Then
            row = row & " | OK"
        Else
            row = row & " | INVALID: " & reason
        End If
        lstActions.AddItem row
    Next i

    lblStatus.Caption = actions.Count & " actions parsed. " & _
                        IIf(anyValid, "Click Apply to run valid actions.", _
                                      "No valid actions; nothing to apply.")
End Sub

Private Sub btnApply_Click()
    If mParsed Is Nothing Then
        lblStatus.Caption = "Click Parse first."
        Exit Sub
    End If
    Dim summary As String
    summary = modExecuteInstructions.ExecuteFromString(CurrentJson())
    lblStatus.Caption = summary
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Function GetStrSafe(d As Object, key As String) As String
    If d.Exists(key) Then GetStrSafe = CStr(d(key))
End Function
'''


def setp(ctrl, **kw):
    for k, v in kw.items():
        try:
            setattr(ctrl, k, v)
        except Exception as e:
            print(f"    [warn] {ctrl.Name}.{k}={v!r}: {e}")


def style_button(ctrl, primary):
    setp(ctrl, BackColor=DARK_BTN if primary else MID_BTN, ForeColor=WHITE)
    try:
        ctrl.Font.Name = FONT_NAME
        ctrl.Font.Size = FONT_SIZE
        ctrl.Font.Bold = True
    except Exception:
        pass


def style_input(ctrl):
    setp(ctrl, BackColor=WHITE, ForeColor=BLACK, SpecialEffect=2)
    try:
        ctrl.Font.Name = FONT_NAME
        ctrl.Font.Size = FONT_SIZE
    except Exception:
        pass


def main() -> int:
    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.Visible = True   # required for Designer access
    try:
        pres = app.Presentations.Open(str(CARRIER), WithWindow=True)
        time.sleep(1.5)
        try:
            comp = pres.VBProject.VBComponents("frmExecute")
            designer = comp.Designer
            try:
                designer.Caption = CAPTION
            except Exception as e:
                print(f"  [warn] caption: {e}")
            try:
                designer.BackColor = DARK_BG
                designer.Font.Name = FONT_NAME
                designer.Font.Size = FONT_SIZE
            except Exception:
                pass

            ctrls = designer.Controls

            # Add btnLoadFile if missing.
            names = {c.Name for c in ctrls}
            if "btnLoadFile" not in names:
                btn = ctrls.Add(CMDBTN, "btnLoadFile", True)
                print("  [add] btnLoadFile")
            else:
                btn = ctrls("btnLoadFile")
            setp(btn, Caption="Load from file...", **LAYOUT["btnLoadFile"])
            style_button(btn, primary=False)

            # Reposition / restyle the existing controls.
            for nm, pos in LAYOUT.items():
                if nm == "btnLoadFile":
                    continue
                try:
                    c = ctrls(nm)
                except Exception:
                    print(f"  [warn] control {nm} not found")
                    continue
                setp(c, **pos)
                if nm == "txtInstructions":
                    setp(c, MultiLine=True, ScrollBars=3, WordWrap=True,
                         MaxLength=0, EnterKeyBehavior=True, AutoTab=False)
                    style_input(c)
                elif nm == "lstActions":
                    style_input(c)
                elif nm in ("btnParse", "btnApply"):
                    style_button(c, primary=True)
                elif nm == "btnCancel":
                    style_button(c, primary=False)

            # Resize the form canvas.
            try:
                comp.Properties("Width").Value = FORM_W
                comp.Properties("Height").Value = FORM_H
            except Exception as e:
                print(f"  [warn] size_form: {e}")

            # Replace code module.
            module = comp.CodeModule
            if module.CountOfLines > 0:
                module.DeleteLines(1, module.CountOfLines)
            module.AddFromString(FRM_EXECUTE_CODE)
            print("  [code] frmExecute module replaced")

            pres.Save()
            print("  [saved] carrier")
            comp.Export(str(FRM))
            print(f"  [export] {FRM}")
        finally:
            pres.Close()
    finally:
        app.Quit()
        time.sleep(1.0)

    # Fix .frm text: ClientWidth/Height (twips) + Caption.
    content = FRM.read_bytes().decode("cp1252")
    content = re.sub(r"(   ClientHeight\s+=\s+)[^\r\n]+", lambda m: m.group(1) + str(FORM_H * 20), content)
    content = re.sub(r"(   ClientWidth\s+=\s+)[^\r\n]+", lambda m: m.group(1) + str(FORM_W * 20), content)
    cap_cp = CAPTION.encode("cp1252", errors="replace").decode("cp1252")
    content = re.sub(r'(   Caption\s+=\s+)[^\r\n]+', lambda m: m.group(1) + f'"{cap_cp}"', content)
    FRM.write_bytes(content.encode("cp1252"))
    print(f"  [fix-frm] dims={FORM_W*20}x{FORM_H*20} caption={CAPTION!r}")

    for f in ("frmExecute.frm", "frmExecute.frx"):
        fp = SRC / f
        print(f"  [verify] {f}: {'OK ' + str(fp.stat().st_size) + ' bytes' if fp.exists() else 'MISSING'}")
    print("[done]")
    return 0


if __name__ == "__main__":
    sys.exit(main())
