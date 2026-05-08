"""Build frmExport and frmExecute UserForms in PPT_AI_Editor.pptm via COM.

Adds the two UserForms programmatically, sets their controls and VBA code,
saves the carrier, and exports .frm + .frx to src/ so update_macros.py
can re-import them on every sync.

Run once: python tools/build_forms.py
"""
import re
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
SRC = REPO_ROOT / "src"

# vbext_ct_MSForm = 3
VB_FORM = 3

# ProgIDs
TEXTBOX  = "Forms.TextBox.1"
CMDBTN   = "Forms.CommandButton.1"
LABEL    = "Forms.Label.1"
LISTBOX  = "Forms.ListBox.1"


# ---------------------------------------------------------------------------
# VBA code strings
# ---------------------------------------------------------------------------

FRM_EXPORT_CODE = '''\
Option Explicit

Private Const PROMPT_TEMPLATE As String = _
"You are editing a PowerPoint presentation. Below is the current state as JSON:" & vbCrLf & vbCrLf & _
"```json" & vbCrLf & "{snapshot}" & vbCrLf & "```" & vbCrLf & vbCrLf & _
"I want the following changes:" & vbCrLf & vbCrLf & _
"[REPLACE THIS LINE WITH YOUR REQUEST]" & vbCrLf & vbCrLf & _
"Return ONLY a valid instructions JSON in this exact format. No prose, no" & vbCrLf & _
"explanation, no markdown fences:" & vbCrLf & vbCrLf & _
"{" & vbCrLf & _
"  ""actions"": [" & vbCrLf & _
"    {""type"": ""<action_type>"", ""slide"": <int>, ""shape_id"": <int>, ...}" & vbCrLf & _
"  ]" & vbCrLf & _
"}" & vbCrLf & vbCrLf & _
"Rules:" & vbCrLf & _
"- Use only shape_ids that exist in the snapshot. Do not invent ids." & vbCrLf & _
"- Slide numbers are 1-based." & vbCrLf & _
"- Colors as #RRGGBB hex." & vbCrLf & _
"- Lengths in points." & vbCrLf & _
"- Allowed action types: set_text, set_font_size, set_font_bold," & vbCrLf & _
"  set_font_italic, set_font_color, set_fill_color, move_shape," & vbCrLf & _
"  resize_shape, delete_shape, add_slide, delete_slide, duplicate_slide," & vbCrLf & _
"  set_cell_text, swap_table_columns, swap_table_rows."

Private Sub UserForm_Initialize()
    On Error Resume Next
    txtSnapshot.Text = modExportSnapshot.BuildSnapshotJson()
    If Err.Number <> 0 Then
        txtSnapshot.Text = "ERROR: " & Err.Description
        Err.Clear
    End If
End Sub

Private Sub btnCopySnapshot_Click()
    CopyToClipboard txtSnapshot.Text
    lblStatus.Caption = "Snapshot copied to clipboard."
End Sub

Private Sub btnCopyWithTemplate_Click()
    Dim payload As String
    payload = Replace(PROMPT_TEMPLATE, "{snapshot}", txtSnapshot.Text)
    CopyToClipboard payload
    lblStatus.Caption = "Snapshot + prompt template copied to clipboard."
End Sub

Private Sub btnSaveTxt_Click()
    Dim deckPath As String: deckPath = ActivePresentation.FullName
    Dim ts As String: ts = Format(Now, "yyyy-mm-dd_hhnnss")
    Dim outPath As String
    outPath = deckPath & "_snapshot_" & ts & ".txt"

    Dim f As Integer: f = FreeFile
    Open outPath For Output As #f
    Print #f, txtSnapshot.Text
    Close #f
    lblStatus.Caption = "Saved: " & outPath
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub CopyToClipboard(s As String)
    Dim doObj As Object
    Set doObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    doObj.SetText s
    doObj.PutInClipboard
End Sub
'''

FRM_EXECUTE_CODE = '''\
Option Explicit

Private mParsed As Object
Private mValid() As Boolean

Private Sub btnParse_Click()
    lblStatus.Caption = ""
    btnApply.Enabled = False
    lstActions.Clear

    On Error Resume Next
    Set mParsed = modJSON.ParseJson(txtInstructions.Text)
    If Err.Number <> 0 Then
        lblStatus.Caption = "Invalid JSON: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    If Not mParsed.Exists("actions") Then
        lblStatus.Caption = "Missing top-level \'actions\' array."
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

    btnApply.Enabled = anyValid
    lblStatus.Caption = actions.Count & " actions parsed. " & _
                        IIf(anyValid, "Click Apply to run valid actions.", _
                                      "No valid actions; nothing to apply.")
End Sub

Private Sub btnApply_Click()
    If mParsed Is Nothing Then Exit Sub
    Dim summary As String
    summary = modExecuteInstructions.ExecuteFromString(txtInstructions.Text)
    lblStatus.Caption = summary
    btnApply.Enabled = False
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Function GetStrSafe(d As Object, key As String) As String
    If d.Exists(key) Then GetStrSafe = CStr(d(key))
End Function
'''


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def remove_component(components, name: str) -> None:
    """Remove a VBComponent by name if it exists."""
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


def set_control_props(ctrl, **kwargs) -> None:
    """Set named properties on a Forms control."""
    for k, v in kwargs.items():
        try:
            setattr(ctrl, k, v)
        except Exception as e:
            print(f"    [warn] Could not set {k}={v!r}: {e}")


def build_frm_export(components, designer_mode: bool = True):
    """Add frmExport to the VBProject."""
    name = "frmExport"
    remove_component(components, name)

    print(f"  [add] {name}")
    comp = components.Add(VB_FORM)
    comp.Name = name

    # Form-level properties via Designer
    designer = comp.Designer
    try:
        designer.Caption = "PPT AI Editor — Export Snapshot"
        designer.Width = 540
        designer.Height = 360
    except Exception as e:
        print(f"  [warn] Could not set form designer props: {e}")
        # Fallback: try Properties collection
        try:
            comp.Properties("Caption").Value = "PPT AI Editor — Export Snapshot"
            comp.Properties("Width").Value = 540
            comp.Properties("Height").Value = 360
        except Exception as e2:
            print(f"  [warn] Properties fallback also failed: {e2}")

    controls = designer.Controls

    # txtSnapshot
    txt = controls.Add(TEXTBOX, "txtSnapshot", True)
    set_control_props(txt, Top=12, Left=12, Width=510, Height=240,
                      MultiLine=True, ScrollBars=3, Locked=True)

    # btnCopySnapshot
    btn1 = controls.Add(CMDBTN, "btnCopySnapshot", True)
    set_control_props(btn1, Caption="Copy snapshot only",
                      Top=264, Left=12, Width=150, Height=24)

    # btnCopyWithTemplate
    btn2 = controls.Add(CMDBTN, "btnCopyWithTemplate", True)
    set_control_props(btn2, Caption="Copy snapshot + prompt template",
                      Top=264, Left=174, Width=230, Height=24)

    # btnSaveTxt
    btn3 = controls.Add(CMDBTN, "btnSaveTxt", True)
    set_control_props(btn3, Caption="Save .txt next to deck",
                      Top=264, Left=414, Width=108, Height=24)

    # btnClose
    btn4 = controls.Add(CMDBTN, "btnClose", True)
    set_control_props(btn4, Caption="Close",
                      Top=300, Left=414, Width=108, Height=24)

    # lblStatus
    lbl = controls.Add(LABEL, "lblStatus", True)
    set_control_props(lbl, Caption="",
                      Top=300, Left=12, Width=390, Height=24)

    # Set VBA code
    module = comp.CodeModule
    if module.CountOfLines > 0:
        module.DeleteLines(1, module.CountOfLines)
    module.AddFromString(FRM_EXPORT_CODE)

    print(f"  [ok] {name} built")
    return comp


def build_frm_execute(components):
    """Add frmExecute to the VBProject."""
    name = "frmExecute"
    remove_component(components, name)

    print(f"  [add] {name}")
    comp = components.Add(VB_FORM)
    comp.Name = name

    # Form-level properties via Designer
    designer = comp.Designer
    try:
        designer.Caption = "PPT AI Editor — Execute Instructions"
        designer.Width = 600
        designer.Height = 480
    except Exception as e:
        print(f"  [warn] Could not set form designer props: {e}")
        try:
            comp.Properties("Caption").Value = "PPT AI Editor — Execute Instructions"
            comp.Properties("Width").Value = 600
            comp.Properties("Height").Value = 480
        except Exception as e2:
            print(f"  [warn] Properties fallback also failed: {e2}")

    controls = designer.Controls

    # txtInstructions
    txt = controls.Add(TEXTBOX, "txtInstructions", True)
    set_control_props(txt, Top=12, Left=12, Width=570, Height=120,
                      MultiLine=True, ScrollBars=3)

    # btnParse
    btn_parse = controls.Add(CMDBTN, "btnParse", True)
    set_control_props(btn_parse, Caption="Parse",
                      Top=144, Left=12, Width=80, Height=24)

    # lstActions
    lst = controls.Add(LISTBOX, "lstActions", True)
    set_control_props(lst, Top=180, Left=12, Width=570, Height=180)

    # btnApply
    btn_apply = controls.Add(CMDBTN, "btnApply", True)
    set_control_props(btn_apply, Caption="Apply", Enabled=False,
                      Top=372, Left=12, Width=80, Height=24)

    # btnCancel
    btn_cancel = controls.Add(CMDBTN, "btnCancel", True)
    set_control_props(btn_cancel, Caption="Cancel",
                      Top=372, Left=102, Width=80, Height=24)

    # lblStatus
    lbl = controls.Add(LABEL, "lblStatus", True)
    set_control_props(lbl, Caption="",
                      Top=408, Left=12, Width=570, Height=36)

    # Set VBA code
    module = comp.CodeModule
    if module.CountOfLines > 0:
        module.DeleteLines(1, module.CountOfLines)
    module.AddFromString(FRM_EXECUTE_CODE)

    print(f"  [ok] {name} built")
    return comp


# ---------------------------------------------------------------------------
# Post-processing helpers
# ---------------------------------------------------------------------------

def _fix_frm_dimensions(frm_path: Path, client_width: int, client_height: int) -> None:
    """Rewrite ClientWidth/ClientHeight in a .frm file to match spec dimensions.

    The COM Designer.Width/Height property is read-only when set via setattr,
    so the exported .frm contains auto-sized values. This function corrects them.
    Dimensions are in twips (1 pt = 20 twips).
    """
    content = frm_path.read_bytes().decode("cp1252")
    content = re.sub(
        r"(   ClientHeight\s+=\s+)[^\r\n]+",
        lambda m: m.group(1) + str(client_height),
        content,
    )
    content = re.sub(
        r"(   ClientWidth\s+=\s+)[^\r\n]+",
        lambda m: m.group(1) + str(client_width),
        content,
    )
    frm_path.write_bytes(content.encode("cp1252"))
    print(f"  [fix-dims] {frm_path.name}: ClientWidth={client_width}, ClientHeight={client_height}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> int:
    if not CARRIER.exists():
        print(f"ERROR: {CARRIER} not found. Run tools/build_carrier.py first.")
        return 1

    print(f"[build_forms] Opening carrier: {CARRIER}")
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    # WithWindow=True required for Designer access (UserForm visual designer)
    app.Visible = True

    try:
        pres = app.Presentations.Open(str(CARRIER), WithWindow=True)
        try:
            project = pres.VBProject
            components = project.VBComponents

            comp_export  = build_frm_export(components)
            comp_execute = build_frm_execute(components)

            pres.Save()
            print("[saved] carrier saved")

            # Export .frm + .frx to src/
            export_path_export  = str(SRC / "frmExport.frm")
            export_path_execute = str(SRC / "frmExecute.frm")

            comp_export.Export(export_path_export)
            print(f"[export] frmExport.frm -> {export_path_export}")

            comp_execute.Export(export_path_execute)
            print(f"[export] frmExecute.frm -> {export_path_execute}")

        finally:
            pres.Close()
    finally:
        app.Quit()
        time.sleep(1.0)

    # Post-process .frm files to set correct ClientWidth/ClientHeight.
    # The Designer.Width/Height setattr is read-only at COM level, so the
    # exported .frm will reflect auto-sized dimensions from the controls.
    # We fix them here to match the spec (dimensions in twips: 1pt = 20 twips).
    _fix_frm_dimensions(SRC / "frmExport.frm",
                        client_width=10800,   # 540pt * 20
                        client_height=7200)   # 360pt * 20
    _fix_frm_dimensions(SRC / "frmExecute.frm",
                        client_width=12000,   # 600pt * 20
                        client_height=9600)   # 480pt * 20

    # Verify files exist
    ok = True
    for fname in ("frmExport.frm", "frmExport.frx", "frmExecute.frm", "frmExecute.frx"):
        fp = SRC / fname
        if fp.exists():
            print(f"[verify] {fname} exists ({fp.stat().st_size} bytes)")
        else:
            print(f"[ERROR] {fname} MISSING from {SRC}")
            ok = False

    print("[done]" if ok else "[FAILED] one or more export files missing")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
