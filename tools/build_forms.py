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

# Dark palette (VBA Long via RGB(r,g,b) = r + g*256 + b*65536)
# #000000 pure black   -> form bg + label bg (labels render transparent)
# #333333 dark gray    -> primary button bg
# #575757 medium gray  -> secondary button bg
# #FFFFFF white        -> all caption text on dark surfaces
# #FFFFFF / #000000    -> white input bg with black input text
DARK_BG    = 0                                       # #000000
DARK_BTN   = 51 + 51 * 256 + 51 * 65536              # #333333 = 3355443
MID_BTN    = 87 + 87 * 256 + 87 * 65536              # #575757 = 5723991
WHITE      = 16777215                                # #FFFFFF
BLACK      = 0                                       # #000000

# Legacy aliases kept so any external script importing these still loads.
GREEN_LIGHT = DARK_BG
GREEN_MID   = MID_BTN
GREEN_DARK  = DARK_BTN

# Font: Cascadia Code (ships with Win11). Falls back to Consolas if missing.
FONT_NAME = "Cascadia Code"
FONT_SIZE_BODY = 10
FONT_SIZE_BTN  = 10


def style_input(ctrl):
    """White input bg, black text, Cascadia Code, sunken edge."""
    try:
        ctrl.BackColor = WHITE
        ctrl.ForeColor = BLACK
        ctrl.SpecialEffect = 2  # fmSpecialEffectSunken
        ctrl.Font.Name = FONT_NAME
        ctrl.Font.Size = FONT_SIZE_BODY
    except Exception as e:
        print(f"    [warn] style_input: {e}")


def style_label(ctrl):
    """Black bg (matches form), white caption text."""
    try:
        ctrl.BackColor = DARK_BG
        ctrl.ForeColor = WHITE
        # Force Font via assignment of a fresh font object to bypass
        # Forms.Label edge case where Font.Size doesn't take effect when
        # the label was imported with a previously-saved font height.
        ctrl.Font.Name = FONT_NAME
        ctrl.Font.Size = FONT_SIZE_BODY
        # Also nuke FontHeight (twips) which can override Font.Size on Labels
        try:
            ctrl.FontHeight = FONT_SIZE_BODY * 20
        except Exception:
            pass
        try:
            ctrl.AutoSize = False
            ctrl.WordWrap = False
        except Exception:
            pass
    except Exception as e:
        print(f"    [warn] style_label: {e}")


def style_button_primary(ctrl):
    """Dark gray (#333) fill, white caption, bold."""
    try:
        ctrl.BackColor = DARK_BTN
        ctrl.ForeColor = WHITE
        ctrl.Font.Name = FONT_NAME
        ctrl.Font.Size = FONT_SIZE_BTN
        ctrl.Font.Bold = True
    except Exception as e:
        print(f"    [warn] style_button_primary: {e}")


def style_button_secondary(ctrl):
    """Medium gray (#575757) fill, white caption, bold."""
    try:
        ctrl.BackColor = MID_BTN
        ctrl.ForeColor = WHITE
        ctrl.Font.Name = FONT_NAME
        ctrl.Font.Size = FONT_SIZE_BTN
        ctrl.Font.Bold = True
    except Exception as e:
        print(f"    [warn] style_button_secondary: {e}")


def style_listbox(ctrl):
    """Same look as input."""
    try:
        ctrl.BackColor = WHITE
        ctrl.ForeColor = BLACK
        ctrl.SpecialEffect = 2
        ctrl.Font.Name = FONT_NAME
        ctrl.Font.Size = FONT_SIZE_BODY
    except Exception as e:
        print(f"    [warn] style_listbox: {e}")


def style_form(designer):
    """Black background + form-default font on the form itself.

    Setting designer.Font is critical: Forms 2.0 controls (especially Labels)
    can render at the form's *default* font size when a hidden FontEffects
    "use default font" bit is set in the .frx blob. Setting the form's Font
    fixes those controls without us touching the bit directly.
    """
    try:
        designer.BackColor = DARK_BG
    except Exception as e:
        print(f"  [warn] style_form bg: {e}")
    try:
        designer.Font.Name = FONT_NAME
        designer.Font.Size = FONT_SIZE_BODY
    except Exception as e:
        print(f"  [warn] style_form font: {e}")


def size_form(comp, outer_width: int, outer_height: int):
    """Force the form's outer dimensions via comp.Properties("Width"/"Height").

    Designer.Width is read-only via setattr but the VBComponent.Properties
    collection accepts the write. This is the only reliable way to enlarge
    the form's design canvas (InsideWidth/InsideHeight) past its 215pt default.
    """
    try:
        comp.Properties("Width").Value = outer_width
        comp.Properties("Height").Value = outer_height
    except Exception as e:
        print(f"  [warn] size_form: {e}")


# ---------------------------------------------------------------------------
# VBA code strings
# ---------------------------------------------------------------------------

# NOTE: frmExport's VBA code is sourced from src/frmExport.frm — the
# project's source of truth, synced by update_macros.py. build_frm_export
# IMPORTS that .frm and PRESERVES its code; it must NOT re-stamp a copy
# here. A stale hardcoded duplicate previously silently reverted shipped
# work on every rebuild (icon-trim, DECK DESIGN PRINCIPLES, the captured-
# template manifest, the Copy-deck-spec / Scan-palette handlers).
# Do NOT reintroduce a FRM_EXPORT_CODE override.

FRM_IMPORT_SLIDES_CODE = '''\
Option Explicit

Private Sub UserForm_Initialize()
    txtPath.Text = ""
    txtRange.Text = ""
    txtPosition.Text = "1"
    lblStatus.Caption = ""
End Sub

Private Sub btnBrowse_Click()
    Dim picked As String
    On Error Resume Next
    Dim fd As Object
    Set fd = Application.FileDialog(3)
    If Not fd Is Nothing Then
        fd.Filters.Clear
        fd.Filters.Add "PowerPoint Files", "*.pptx; *.pptm"
        If fd.Show = -1 Then
            picked = fd.SelectedItems(1)
        End If
    End If
    On Error GoTo 0

    If Len(picked) = 0 Then
        picked = InputBox("Path to source deck:", "Source deck")
    End If

    If Len(picked) > 0 Then
        txtPath.Text = picked
    End If
End Sub

Private Sub btnImport_Click()
    If Len(txtPath.Text) = 0 Or Len(txtRange.Text) = 0 Or Len(txtPosition.Text) = 0 Then
        lblStatus.Caption = "Fill in source deck, slide range, and position first."
        Exit Sub
    End If
    On Error GoTo Failure
    Dim ids As Variant
    ids = ParseRange(txtRange.Text)
    Dim pos As Long: pos = CLng(txtPosition.Text)

    Dim before As Long: before = ActivePresentation.Slides.Count
    modActionsSlide.Do_import_slides_from_deck txtPath.Text, ids, pos
    Dim afterCount As Long: afterCount = ActivePresentation.Slides.Count
    lblStatus.Caption = "Imported " & (afterCount - before) & " slide(s) at position " & pos
    Exit Sub
Failure:
    lblStatus.Caption = "ERROR: " & Err.Description
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Function ParseRange(s As String) As Variant
    Dim parts() As String
    parts = Split(s, ",")
    Dim col As New Collection
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim p As String: p = Trim(parts(i))
        If InStr(p, "-") > 0 Then
            Dim ab() As String: ab = Split(p, "-")
            Dim a As Long: a = CLng(Trim(ab(0)))
            Dim b As Long: b = CLng(Trim(ab(1)))
            Dim k As Long
            For k = a To b
                col.Add k
            Next k
        Else
            col.Add CLng(p)
        End If
    Next i
    Dim arr() As Long
    ReDim arr(0 To col.Count - 1)
    Dim j As Long
    For j = 1 To col.Count
        arr(j - 1) = col(j)
    Next j
    ParseRange = arr
End Function
'''

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
        time.sleep(0.3)


def clear_controls(designer) -> None:
    """Remove all controls from a UserForm designer."""
    names = [c.Name for c in designer.Controls]
    for n in names:
        try:
            designer.Controls.Remove(n)
        except Exception as e:
            print(f"    [warn] Could not remove control {n!r}: {e}")


def get_or_add(designer, progid: str, name: str):
    """Return an existing control with `name`, or Add a new one of `progid`."""
    for c in designer.Controls:
        try:
            if c.Name == name:
                return c
        except Exception:
            pass
    return designer.Controls.Add(progid, name, True)


def set_control_props(ctrl, **kwargs) -> None:
    """Set named properties on a Forms control."""
    for k, v in kwargs.items():
        try:
            setattr(ctrl, k, v)
        except Exception as e:
            print(f"    [warn] Could not set {k}={v!r}: {e}")


def build_frm_export(components, designer_mode: bool = True):
    """Add frmExport to the VBProject.

    Strategy: Import the existing .frm (bypasses VBA name-reservation bug
    triggered by setting comp.Name on a freshly-Add'd component while the
    name is still referenced by modUI), then clear controls and rebuild.
    """
    name = "frmExport"
    remove_component(components, name)

    frm_path = SRC / "frmExport.frm"
    print(f"  [import] {name} from {frm_path.name}")
    comp = components.Import(str(frm_path))

    designer = comp.Designer
    try:
        designer.Caption = "Decko.ai \U0001F916  Export Snapshot"
    except Exception as e:
        print(f"  [warn] Caption: {e}")

    style_form(designer)
    clear_controls(designer)
    size_form(comp, outer_width=540, outer_height=360)
    controls = designer.Controls

    # txtSnapshot
    txt = controls.Add(TEXTBOX, "txtSnapshot", True)
    set_control_props(txt, Top=12, Left=12, Width=510, Height=240,
                      MultiLine=True, ScrollBars=3, Locked=True)
    style_input(txt)

    # Layout: 12pt left margin, 12pt gaps, 12pt right margin (form W=540).
    # Available content width = 540 - 12 - 12 = 516pt.
    # Geometry mirrors tools/add_export_buttons.py exactly so a rebuild
    # reproduces the same form. Row A = primary copy actions; Row B = the
    # two action-shortcut buttons + Close; lblStatus on its own line.
    # btnCopySnapshot (primary)
    btn1 = controls.Add(CMDBTN, "btnCopySnapshot", True)
    set_control_props(btn1, Caption="Copy snapshot only",
                      Top=264, Left=12, Width=120, Height=24)
    style_button_primary(btn1)

    # btnCopyWithTemplate (primary)
    btn2 = controls.Add(CMDBTN, "btnCopyWithTemplate", True)
    set_control_props(btn2, Caption="Copy snapshot + prompt template",
                      Top=264, Left=140, Width=200, Height=24)
    style_button_primary(btn2)

    # btnSaveTxt (secondary)
    btn3 = controls.Add(CMDBTN, "btnSaveTxt", True)
    set_control_props(btn3, Caption="Save to .txt",
                      Top=264, Left=348, Width=150, Height=24)
    style_button_secondary(btn3)

    # btnCopyDeckSpec (primary) — extract_spec -> clipboard
    btn5 = controls.Add(CMDBTN, "btnCopyDeckSpec", True)
    set_control_props(btn5, Caption="Copy deck spec",
                      Top=298, Left=12, Width=170, Height=24)
    style_button_primary(btn5)

    # btnScanPalette (primary) — scan_palette -> clipboard
    btn6 = controls.Add(CMDBTN, "btnScanPalette", True)
    set_control_props(btn6, Caption="Scan palette",
                      Top=298, Left=190, Width=150, Height=24)
    style_button_primary(btn6)

    # btnClose (secondary)
    btn4 = controls.Add(CMDBTN, "btnClose", True)
    set_control_props(btn4, Caption="Close",
                      Top=298, Left=348, Width=150, Height=24)
    style_button_secondary(btn4)

    # lblStatus
    lbl = controls.Add(LABEL, "lblStatus", True)
    set_control_props(lbl, Caption="",
                      Top=330, Left=12, Width=486, Height=24)
    style_label(lbl)

    # Code is PRESERVED from the imported src/frmExport.frm (source of
    # truth, synced by update_macros.py). Deliberately NOT re-stamped from
    # a hardcoded constant — see the FRM_EXPORT_CODE note near the top of
    # this file. The imported form already carries PromptTemplate, the
    # snapshot/clipboard subs, and the btnCopyDeckSpec/btnScanPalette
    # handlers; rebuilding only refreshes the control layout above.

    print(f"  [ok] {name} built (code preserved from src/frmExport.frm)")
    return comp


def build_frm_execute(components):
    """Add frmExecute via Import + clear_controls (bypasses name-reservation bug)."""
    name = "frmExecute"
    remove_component(components, name)

    frm_path = SRC / "frmExecute.frm"
    print(f"  [import] {name} from {frm_path.name}")
    comp = components.Import(str(frm_path))

    designer = comp.Designer
    try:
        designer.Caption = "Decko.ai \U0001F916  Execute Instructions"
    except Exception as e:
        print(f"  [warn] Caption: {e}")

    style_form(designer)
    clear_controls(designer)
    size_form(comp, outer_width=600, outer_height=520)
    controls = designer.Controls

    # txtInstructions — large, scrollable both ways. Note: MSForms textboxes
    # still mangle very large pastes; btnLoadFile is the reliable path for those.
    txt = controls.Add(TEXTBOX, "txtInstructions", True)
    set_control_props(txt, Top=12, Left=12, Width=570, Height=160,
                      MultiLine=True, ScrollBars=3, WordWrap=True,
                      MaxLength=0, EnterKeyBehavior=True, AutoTab=False)
    style_input(txt)

    # btnParse (primary)
    btn_parse = controls.Add(CMDBTN, "btnParse", True)
    set_control_props(btn_parse, Caption="Parse",
                      Top=180, Left=12, Width=80, Height=24)
    style_button_primary(btn_parse)

    # btnLoadFile (secondary) — load actions JSON straight from a file,
    # bypassing the textbox so large batches are not corrupted.
    btn_load = controls.Add(CMDBTN, "btnLoadFile", True)
    set_control_props(btn_load, Caption="Load from file...",
                      Top=180, Left=100, Width=150, Height=24)
    style_button_secondary(btn_load)

    # lstActions
    lst = controls.Add(LISTBOX, "lstActions", True)
    set_control_props(lst, Top=214, Left=12, Width=570, Height=160)
    style_listbox(lst)

    # btnApply (primary). Enabled=True at design-time so the screenshot is
    # crisp; UserForm_Initialize disables it at runtime until Parse succeeds.
    btn_apply = controls.Add(CMDBTN, "btnApply", True)
    set_control_props(btn_apply, Caption="Apply",
                      Top=384, Left=12, Width=80, Height=24)
    style_button_primary(btn_apply)

    # btnCancel (secondary)
    btn_cancel = controls.Add(CMDBTN, "btnCancel", True)
    set_control_props(btn_cancel, Caption="Cancel",
                      Top=384, Left=102, Width=80, Height=24)
    style_button_secondary(btn_cancel)

    # lblStatus
    lbl = controls.Add(LABEL, "lblStatus", True)
    set_control_props(lbl, Caption="",
                      Top=416, Left=12, Width=570, Height=60)
    style_label(lbl)

    # Set VBA code
    module = comp.CodeModule
    if module.CountOfLines > 0:
        module.DeleteLines(1, module.CountOfLines)
    module.AddFromString(FRM_EXECUTE_CODE)

    print(f"  [ok] {name} built")
    return comp


def build_frm_import_slides(components):
    """Add frmImportSlides via Import + clear_controls (bypasses name-reservation bug)."""
    name = "frmImportSlides"
    remove_component(components, name)

    frm_path = SRC / "frmImportSlides.frm"
    print(f"  [import] {name} from {frm_path.name}")
    comp = components.Import(str(frm_path))

    designer = comp.Designer
    try:
        designer.Caption = "Decko.ai \U0001F916  Import Slides"
    except Exception as e:
        print(f"  [warn] Caption: {e}")

    style_form(designer)
    clear_controls(designer)
    size_form(comp, outer_width=480, outer_height=320)

    controls = designer.Controls

    # lblPath
    lbl_path = get_or_add(designer, LABEL, "lblPath")
    set_control_props(lbl_path, Caption="Source deck",
                      Top=12, Left=12, Width=100, Height=20)
    style_label(lbl_path)

    # txtPath
    txt_path = get_or_add(designer, TEXTBOX, "txtPath")
    set_control_props(txt_path, Top=12, Left=120, Width=280, Height=22,
                      Locked=True)
    style_input(txt_path)

    # btnBrowse (secondary)
    btn_browse = get_or_add(designer, CMDBTN, "btnBrowse")
    set_control_props(btn_browse, Caption="Browse...",
                      Top=12, Left=408, Width=60, Height=22)
    style_button_secondary(btn_browse)

    # lblRange
    lbl_range = get_or_add(designer, LABEL, "lblRange")
    set_control_props(lbl_range, Caption="Slide range (e.g. 1-3,5,7-9)",
                      Top=50, Left=12, Width=240, Height=20)
    style_label(lbl_range)

    # txtRange
    txt_range = get_or_add(designer, TEXTBOX, "txtRange")
    set_control_props(txt_range, Top=50, Left=256, Width=212, Height=22)
    style_input(txt_range)

    # lblPosition
    lbl_pos = get_or_add(designer, LABEL, "lblPosition")
    set_control_props(lbl_pos, Caption="Insert at position",
                      Top=90, Left=12, Width=240, Height=20)
    style_label(lbl_pos)

    # txtPosition
    txt_pos = get_or_add(designer, TEXTBOX, "txtPosition")
    set_control_props(txt_pos, Top=90, Left=256, Width=60, Height=22)
    style_input(txt_pos)

    # btnImport (primary). Enabled=True at design-time so the screenshot is
    # crisp; UserForm_Initialize disables it at runtime until inputs filled.
    btn_import = get_or_add(designer, CMDBTN, "btnImport")
    set_control_props(btn_import, Caption="Import",
                      Top=250, Left=12, Width=80, Height=28)
    style_button_primary(btn_import)

    # btnCancel (secondary)
    btn_cancel = get_or_add(designer, CMDBTN, "btnCancel")
    set_control_props(btn_cancel, Caption="Cancel",
                      Top=250, Left=100, Width=80, Height=28)
    style_button_secondary(btn_cancel)

    # lblStatus
    lbl_status = get_or_add(designer, LABEL, "lblStatus")
    set_control_props(lbl_status, Caption="",
                      Top=130, Left=12, Width=456, Height=100)
    style_label(lbl_status)

    # Set VBA code
    module = comp.CodeModule
    if module.CountOfLines > 0:
        module.DeleteLines(1, module.CountOfLines)
    module.AddFromString(FRM_IMPORT_SLIDES_CODE)

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


def _fix_frm_caption(frm_path: Path, caption: str) -> None:
    """Rewrite the Caption line in a .frm file.

    Some COM Designer Caption setattr calls are silently ignored (when Width/Height
    cannot be set), leaving the auto-generated 'UserForm1' caption. This corrects it.
    """
    content = frm_path.read_bytes().decode("cp1252")
    # Caption is stored in cp1252; em-dash U+2014 -> 0x97
    caption_cp1252 = caption.encode("cp1252", errors="replace").decode("cp1252")
    content = re.sub(
        r'(   Caption\s+=\s+)[^\r\n]+',
        lambda m: m.group(1) + f'"{caption_cp1252}"',
        content,
    )
    frm_path.write_bytes(content.encode("cp1252"))
    print(f"  [fix-caption] {frm_path.name}: Caption={caption!r}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def _remove_old_forms(components) -> None:
    """Pass 1: remove existing UserForms so name reservations clear on save."""
    for name in ("frmExport", "frmExecute", "frmImportSlides"):
        remove_component(components, name)


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
        # ---- Pass 1: open, remove existing forms, save, close ----
        # PPT reserves form names persistently while the file is open. Removing
        # then re-adding in the same session triggers error -2146828213. Solve
        # by saving + closing between remove and add.
        print("[pass 1] Removing existing UserForms")
        pres = app.Presentations.Open(str(CARRIER), WithWindow=True)
        try:
            _remove_old_forms(pres.VBProject.VBComponents)
            pres.Save()
        finally:
            pres.Close()
        time.sleep(2.0)  # let Office release name reservations

        # ---- Pass 2: reopen and build fresh forms ----
        print("[pass 2] Building fresh UserForms")
        pres = app.Presentations.Open(str(CARRIER), WithWindow=True)
        time.sleep(2.0)  # let VBProject settle
        try:
            project = pres.VBProject
            components = project.VBComponents

            comp_export  = build_frm_export(components)
            comp_execute = build_frm_execute(components)
            comp_import_slides = build_frm_import_slides(components)

            pres.Save()
            print("[saved] carrier saved")

            # Export .frm + .frx to src/
            export_path_export  = str(SRC / "frmExport.frm")
            export_path_execute = str(SRC / "frmExecute.frm")
            export_path_import_slides = str(SRC / "frmImportSlides.frm")

            comp_export.Export(export_path_export)
            print(f"[export] frmExport.frm -> {export_path_export}")

            comp_execute.Export(export_path_execute)
            print(f"[export] frmExecute.frm -> {export_path_execute}")

            comp_import_slides.Export(export_path_import_slides)
            print(f"[export] frmImportSlides.frm -> {export_path_import_slides}")

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
                        client_height=10400)  # 520pt * 20
    _fix_frm_dimensions(SRC / "frmImportSlides.frm",
                        client_width=9600,    # 480pt * 20
                        client_height=6400)   # 320pt * 20
    # Caption fix: COM Designer.Caption silently fails on Imported forms,
    # leaving the imported .frm caption (e.g. old "PPT AI Editor"). We rewrite
    # the Caption line in the .frm text directly. Limited to cp1252 chars.
    _fix_frm_caption(SRC / "frmExport.frm",       "Decko.ai • Export Snapshot")
    _fix_frm_caption(SRC / "frmExecute.frm",      "Decko.ai • Execute Instructions")
    _fix_frm_caption(SRC / "frmImportSlides.frm", "Decko.ai • Import Slides")

    # Verify files exist
    ok = True
    for fname in ("frmExport.frm", "frmExport.frx",
                  "frmExecute.frm", "frmExecute.frx",
                  "frmImportSlides.frm", "frmImportSlides.frx"):
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
