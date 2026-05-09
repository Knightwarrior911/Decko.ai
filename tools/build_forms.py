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

FRM_EXPORT_CODE = '''\
Option Explicit

Private Function PromptTemplate() As String
    Dim s As String

    s = "You are editing a PowerPoint presentation. Below is the current state as JSON:" & vbCrLf & vbCrLf
    s = s & "```json" & vbCrLf & "{snapshot}" & vbCrLf & "```" & vbCrLf & vbCrLf
    s = s & "I want the following changes:" & vbCrLf & vbCrLf
    s = s & "[REPLACE THIS LINE WITH YOUR REQUEST]" & vbCrLf & vbCrLf

    s = s & "Return ONLY a valid instructions JSON. No prose, no explanation, no markdown" & vbCrLf
    s = s & "code fences. Top-level shape:" & vbCrLf & vbCrLf
    s = s & "{""actions"": [ <action>, <action>, ... ]}" & vbCrLf & vbCrLf

    s = s & "Each action is one of EXACTLY these schemas. Field names are STRICT - do not" & vbCrLf
    s = s & "rename ""value"" to ""text""/""color""/""size""/""fill"". Use names verbatim." & vbCrLf & vbCrLf

    s = s & "ATOMIC OPS (V1):" & vbCrLf
    s = s & "  {""type"":""set_text"",""slide"":1,""shape_id"":3,""value"":""Hello""}" & vbCrLf
    s = s & "  {""type"":""set_font_size"",""slide"":1,""shape_id"":3,""value"":28}" & vbCrLf
    s = s & "  {""type"":""set_font_bold"",""slide"":1,""shape_id"":3,""value"":true}" & vbCrLf
    s = s & "  {""type"":""set_font_italic"",""slide"":1,""shape_id"":3,""value"":false}" & vbCrLf
    s = s & "  {""type"":""set_font_color"",""slide"":1,""shape_id"":3,""value"":""#FF0000""}" & vbCrLf
    s = s & "  {""type"":""set_fill_color"",""slide"":1,""shape_id"":4,""value"":""#2E75B6""}" & vbCrLf
    s = s & "  {""type"":""move_shape"",""slide"":1,""shape_id"":4,""left"":100,""top"":200}" & vbCrLf
    s = s & "  {""type"":""resize_shape"",""slide"":1,""shape_id"":4,""width"":250,""height"":80}" & vbCrLf
    s = s & "  {""type"":""delete_shape"",""slide"":1,""shape_id"":7}" & vbCrLf
    s = s & "  {""type"":""add_slide"",""position"":3,""layout_index"":1}" & vbCrLf
    s = s & "  {""type"":""delete_slide"",""slide"":4}" & vbCrLf
    s = s & "  {""type"":""duplicate_slide"",""slide"":2}" & vbCrLf
    s = s & "  {""type"":""set_cell_text"",""slide"":1,""shape_id"":5,""row"":2,""col"":1,""value"":""Revenue""}" & vbCrLf
    s = s & "  {""type"":""swap_table_columns"",""slide"":1,""shape_id"":5,""col_a"":1,""col_b"":2}" & vbCrLf
    s = s & "  {""type"":""swap_table_rows"",""slide"":1,""shape_id"":5,""row_a"":1,""row_b"":2}" & vbCrLf & vbCrLf

    s = s & "GRANULAR TEXT:" & vbCrLf
    s = s & "  {""type"":""set_paragraph_text"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":""...""}" & vbCrLf
    s = s & "  {""type"":""add_paragraph"",""slide"":1,""shape_id"":3,""after_paragraph_index"":-1,""value"":""...""}" & vbCrLf
    s = s & "  {""type"":""delete_paragraph"",""slide"":1,""shape_id"":3,""paragraph_index"":2}" & vbCrLf
    s = s & "  {""type"":""set_bullet_style"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":""disc""}" & vbCrLf
    s = s & "  {""type"":""set_indent_level"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":1}" & vbCrLf
    s = s & "  {""type"":""set_paragraph_font_size"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":24}" & vbCrLf
    s = s & "  {""type"":""set_paragraph_font_color"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":""#1F4E79""}" & vbCrLf
    s = s & "  {""type"":""find_replace_text"",""scope"":""deck"",""find"":""ACME"",""replace"":""NewCo""}" & vbCrLf & vbCrLf

    s = s & "LAYOUT / COMPOSITION:" & vbCrLf
    s = s & "  {""type"":""align_shapes"",""slide"":1,""shape_ids"":[3,4,5],""anchor"":""top""}" & vbCrLf
    s = s & "  {""type"":""distribute_horizontal"",""slide"":1,""shape_ids"":[3,4,5]}" & vbCrLf
    s = s & "  {""type"":""distribute_vertical"",""slide"":1,""shape_ids"":[3,4,5]}" & vbCrLf
    s = s & "  {""type"":""tile_grid"",""slide"":1,""shape_ids"":[3,4,5,6],""cols"":2,""gap_pt"":10}" & vbCrLf
    s = s & "  {""type"":""fit_to_slide_margins"",""slide"":1,""shape_id"":3,""margin_pt"":36}" & vbCrLf
    s = s & "  {""type"":""add_line"",""slide"":1,""x1"":50,""y1"":50,""x2"":700,""y2"":50,""color"":""#000000"",""weight_pt"":1.5}" & vbCrLf
    s = s & "  {""type"":""add_shape"",""slide"":1,""kind"":""capsule"",""pos"":{""left"":100,""top"":100,""width"":200,""height"":60},""fill"":""#1F4E79"",""stroke"":null,""stroke_weight_pt"":0}" & vbCrLf
    s = s & "  {""type"":""set_shape_kind"",""slide"":1,""shape_id"":4,""kind"":""capsule""}" & vbCrLf
    s = s & "  {""type"":""clear_slide"",""slide"":1,""keep_shape_ids"":[3]}" & vbCrLf
    s = s & "  {""type"":""move_shape_relative"",""slide"":1,""shape_id"":4,""dx_pt"":0,""dy_pt"":50}" & vbCrLf & vbCrLf

    s = s & "CROSS-CUTTING BATCH:" & vbCrLf
    s = s & "  {""type"":""recolor_fill_match"",""scope"":""deck"",""from"":""#0000FF"",""to"":""#FF0000""}" & vbCrLf
    s = s & "  {""type"":""recolor_font_match"",""scope"":""slide:2"",""from"":""#000000"",""to"":""#1F4E79""}" & vbCrLf
    s = s & "  {""type"":""delete_shapes_match"",""scope"":""deck"",""text_contains"":""Confidential""}" & vbCrLf & vbCrLf

    s = s & "SPEAKER NOTES:" & vbCrLf
    s = s & "  {""type"":""set_speaker_notes"",""slide"":1,""value"":""Talking points...""}" & vbCrLf
    s = s & "  {""type"":""append_speaker_notes"",""slide"":1,""value"":""Add this too""}" & vbCrLf & vbCrLf

    s = s & "IMAGES (LOCAL FILE PATHS ONLY - no URLs):" & vbCrLf
    s = s & "  {""type"":""insert_picture"",""slide"":1,""path"":""C:\\\\path\\\\to\\\\img.png"",""pos"":{""left"":50,""top"":50,""width"":200,""height"":150}}" & vbCrLf
    s = s & "  {""type"":""replace_picture"",""slide"":1,""shape_id"":7,""path"":""C:\\\\path\\\\to\\\\new.png""}" & vbCrLf & vbCrLf

    s = s & "SLIDE STRUCTURE:" & vbCrLf
    s = s & "  {""type"":""move_slide"",""from"":3,""to"":1}" & vbCrLf
    s = s & "  {""type"":""extract_slides"",""slide_indices"":[1,3,5],""output_path"":""C:\\\\path\\\\out.pptx""}" & vbCrLf
    s = s & "  {""type"":""import_slides_from_deck"",""source_path"":""C:\\\\path\\\\other.pptx"",""slide_indices"":[1,2],""target_position"":3}" & vbCrLf & vbCrLf

    s = s & "TABLES:" & vbCrLf
    s = s & "  {""type"":""add_table_row"",""slide"":1,""shape_id"":5,""after_row"":2}" & vbCrLf
    s = s & "  {""type"":""delete_table_row"",""slide"":1,""shape_id"":5,""row"":3}" & vbCrLf
    s = s & "  {""type"":""add_table_col"",""slide"":1,""shape_id"":5,""after_col"":2}" & vbCrLf
    s = s & "  {""type"":""delete_table_col"",""slide"":1,""shape_id"":5,""col"":3}" & vbCrLf
    s = s & "  {""type"":""merge_cells"",""slide"":1,""shape_id"":5,""row_a"":1,""col_a"":1,""row_b"":1,""col_b"":3}" & vbCrLf & vbCrLf

    s = s & "GROUPS:" & vbCrLf
    s = s & "  {""type"":""group_shapes"",""slide"":1,""shape_ids"":[3,4,5]}" & vbCrLf
    s = s & "  {""type"":""ungroup"",""slide"":1,""shape_id"":12}" & vbCrLf & vbCrLf

    s = s & "CONNECTORS:" & vbCrLf
    s = s & "  {""type"":""add_connector"",""slide"":1,""from_shape_id"":3,""to_shape_id"":7,""kind"":""elbow"",""arrow_end"":""filled"",""color"":""#000000"",""weight_pt"":1.5}" & vbCrLf & vbCrLf

    s = s & "NATIVE CHARTS (Shape.HasChart=True only; pasted images skipped):" & vbCrLf
    s = s & "  {""type"":""set_chart_type"",""slide"":1,""shape_id"":4,""value"":""barClustered""}" & vbCrLf
    s = s & "  {""type"":""set_chart_title"",""slide"":1,""shape_id"":4,""value"":""Q3 Revenue"",""enabled"":true}" & vbCrLf
    s = s & "  {""type"":""set_chart_axis_title"",""slide"":1,""shape_id"":4,""axis"":""x"",""value"":""Quarter""}" & vbCrLf
    s = s & "  {""type"":""set_chart_legend_position"",""slide"":1,""shape_id"":4,""value"":""bottom""}" & vbCrLf
    s = s & "  {""type"":""set_series_color"",""slide"":1,""shape_id"":4,""series_index"":1,""value"":""#1F4E79""}" & vbCrLf & vbCrLf

    s = s & "RULES:" & vbCrLf
    s = s & "- Use only shape_ids that exist in the snapshot. Do not invent ids." & vbCrLf
    s = s & "- Slide / row / col / series / position numbers are 1-based." & vbCrLf
    s = s & "- paragraph_index is 0-based to match the snapshot's paragraphs[].index." & vbCrLf
    s = s & "- after_paragraph_index = -1 means insert at top." & vbCrLf
    s = s & "- Colors are #RRGGBB hex strings." & vbCrLf
    s = s & "- Lengths in points." & vbCrLf
    s = s & "- Booleans are JSON true / false (lowercase, no quotes)." & vbCrLf
    s = s & "- For 'every X with property P -> Y' requests: enumerate matching shape_ids" & vbCrLf
    s = s & "  in the snapshot and emit one action per match (or use a *_match helper)." & vbCrLf
    s = s & "- For 'rebuild this slide': use clear_slide first, then a sequence of" & vbCrLf
    s = s & "  add_shape / set_text / move_shape actions to populate." & vbCrLf
    s = s & "- File paths must be absolute and use double backslashes in JSON strings." & vbCrLf
    s = s & "- One field name per action - never substitute aliases."

    PromptTemplate = s
End Function

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
    payload = Replace(PromptTemplate(), "{snapshot}", txtSnapshot.Text)
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

Private Sub UserForm_Initialize()
    lblStatus.Caption = ""
End Sub

Private Sub btnParse_Click()
    lblStatus.Caption = ""
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
    summary = modExecuteInstructions.ExecuteFromString(txtInstructions.Text)
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
    # btnCopySnapshot (primary)
    btn1 = controls.Add(CMDBTN, "btnCopySnapshot", True)
    set_control_props(btn1, Caption="Copy snapshot only",
                      Top=264, Left=12, Width=140, Height=24)
    style_button_primary(btn1)

    # btnCopyWithTemplate (primary)
    btn2 = controls.Add(CMDBTN, "btnCopyWithTemplate", True)
    set_control_props(btn2, Caption="Copy snapshot + prompt template",
                      Top=264, Left=160, Width=215, Height=24)
    style_button_primary(btn2)

    # btnSaveTxt (secondary)
    btn3 = controls.Add(CMDBTN, "btnSaveTxt", True)
    set_control_props(btn3, Caption="Save to .txt",
                      Top=264, Left=383, Width=110, Height=24)
    style_button_secondary(btn3)

    # btnClose (secondary)
    btn4 = controls.Add(CMDBTN, "btnClose", True)
    set_control_props(btn4, Caption="Close",
                      Top=300, Left=383, Width=110, Height=24)
    style_button_secondary(btn4)

    # lblStatus
    lbl = controls.Add(LABEL, "lblStatus", True)
    set_control_props(lbl, Caption="",
                      Top=300, Left=12, Width=363, Height=24)
    style_label(lbl)

    # Set VBA code
    module = comp.CodeModule
    if module.CountOfLines > 0:
        module.DeleteLines(1, module.CountOfLines)
    module.AddFromString(FRM_EXPORT_CODE)

    print(f"  [ok] {name} built")
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
    size_form(comp, outer_width=600, outer_height=480)
    controls = designer.Controls

    # txtInstructions
    txt = controls.Add(TEXTBOX, "txtInstructions", True)
    set_control_props(txt, Top=12, Left=12, Width=570, Height=120,
                      MultiLine=True, ScrollBars=3)
    style_input(txt)

    # btnParse (primary)
    btn_parse = controls.Add(CMDBTN, "btnParse", True)
    set_control_props(btn_parse, Caption="Parse",
                      Top=144, Left=12, Width=80, Height=24)
    style_button_primary(btn_parse)

    # lstActions
    lst = controls.Add(LISTBOX, "lstActions", True)
    set_control_props(lst, Top=180, Left=12, Width=570, Height=180)
    style_listbox(lst)

    # btnApply (primary). Enabled=True at design-time so the screenshot is
    # crisp; UserForm_Initialize disables it at runtime until Parse succeeds.
    btn_apply = controls.Add(CMDBTN, "btnApply", True)
    set_control_props(btn_apply, Caption="Apply",
                      Top=372, Left=12, Width=80, Height=24)
    style_button_primary(btn_apply)

    # btnCancel (secondary)
    btn_cancel = controls.Add(CMDBTN, "btnCancel", True)
    set_control_props(btn_cancel, Caption="Cancel",
                      Top=372, Left=102, Width=80, Height=24)
    style_button_secondary(btn_cancel)

    # lblStatus
    lbl = controls.Add(LABEL, "lblStatus", True)
    set_control_props(lbl, Caption="",
                      Top=408, Left=12, Width=570, Height=36)
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
                        client_height=9600)   # 480pt * 20
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
