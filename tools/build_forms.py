"""Build frmExport, frmExecute, and frmImportSlides UserForms in PPT_AI_Editor.pptm via COM.

Adds the three UserForms programmatically, sets their controls and VBA code,
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
# macOS-style design system constants
# OLE colors are BGR-packed: color = R + G*256 + B*65536
# ---------------------------------------------------------------------------

COLOR_BG_WINDOW  = 245 + 245 * 256 + 247 * 65536   # RGB(245, 245, 247) – window bg
COLOR_BG_CARD    = 255 + 255 * 256 + 255 * 65536   # RGB(255, 255, 255) – card / input bg
COLOR_BORDER     = 229 + 229 * 256 + 234 * 65536   # RGB(229, 229, 234) – subtle border
COLOR_TEXT_PRI   = 28  +  28 * 256 +  30 * 65536   # RGB(28, 28, 30)    – primary text
COLOR_TEXT_SEC   = 99  +  99 * 256 + 102 * 65536   # RGB(99, 99, 102)   – secondary text
COLOR_TEXT_TER   = 142 + 142 * 256 + 147 * 65536   # RGB(142, 142, 147) – tertiary / hint
COLOR_ACCENT     =   0 + 122 * 256 + 255 * 65536   # RGB(0, 122, 255)   – macOS blue
COLOR_SUCCESS    =  52 + 199 * 256 +  89 * 65536   # RGB(52, 199, 89)   – success green
COLOR_ERROR      = 255 +  59 * 256 +  48 * 65536   # RGB(255, 59, 48)   – error red
COLOR_DISABLED   = 199 + 199 * 256 + 204 * 65536   # RGB(199, 199, 204) – disabled gray
COLOR_WHITE      = 255 + 255 * 256 + 255 * 65536   # pure white

FONT_BODY    = ("Segoe UI", 13, False)
FONT_HEADING = ("Segoe UI", 16, True)
FONT_SECTION = ("Segoe UI", 13, True)
FONT_BUTTON  = ("Segoe UI", 13, True)
FONT_HINT    = ("Segoe UI", 11, False)


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
    btnImport.Enabled = False
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
        UpdateImportButton
    End If
End Sub

Private Sub txtRange_Change()
    UpdateImportButton
End Sub

Private Sub txtPosition_Change()
    UpdateImportButton
End Sub

Private Sub UpdateImportButton()
    btnImport.Enabled = (Len(txtPath.Text) > 0 And Len(txtRange.Text) > 0 And Len(txtPosition.Text) > 0)
End Sub

Private Sub btnImport_Click()
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
# Style helpers
# ---------------------------------------------------------------------------

def _set(ctrl, **kwargs):
    """Set named properties on a Forms control, swallowing individual failures."""
    for k, v in kwargs.items():
        try:
            setattr(ctrl, k, v)
        except Exception as e:
            print(f"    [warn] Could not set {k}={v!r}: {e}")


def _font(ctrl, name, size, bold=False):
    """Apply font properties to a control's Font object."""
    try:
        ctrl.Font.Name = name
        ctrl.Font.Size = size
        ctrl.Font.Bold = bold
    except Exception as e:
        print(f"    [warn] Font set failed: {e}")


def style_form(designer, title, width_pt):
    """Add a macOS-style title strip + 1pt divider to the form."""
    designer.BackColor = COLOR_BG_WINDOW

    # Title strip (card-white bar across the top)
    title_lbl = designer.Controls.Add(LABEL, "lblFormTitle", True)
    _set(title_lbl,
         Caption=title,
         Top=0, Left=0, Width=width_pt, Height=48,
         BackColor=COLOR_BG_CARD, BackStyle=1,
         ForeColor=COLOR_TEXT_PRI,
         TextAlign=2)
    _font(title_lbl, FONT_HEADING[0], FONT_HEADING[1], FONT_HEADING[2])

    # 1pt divider line
    divider = designer.Controls.Add(LABEL, "lblFormDivider", True)
    _set(divider,
         Caption="",
         Top=48, Left=0, Width=width_pt, Height=1,
         BackColor=COLOR_BORDER, BackStyle=1)


def style_textbox(ctrl, multiline=False, locked=False):
    _set(ctrl,
         BackColor=COLOR_BG_CARD, BackStyle=1,
         ForeColor=COLOR_TEXT_PRI,
         BorderStyle=1)
    try:
        ctrl.SpecialEffect = 0   # fmSpecialEffectFlat
    except Exception as e:
        print(f"    [warn] SpecialEffect=0 failed: {e}")
    _font(ctrl, FONT_BODY[0], FONT_BODY[1])
    _set(ctrl, MultiLine=multiline, Locked=locked)
    if multiline:
        _set(ctrl, ScrollBars=3)


def style_listbox(ctrl):
    _set(ctrl,
         BackColor=COLOR_BG_CARD, BackStyle=1,
         ForeColor=COLOR_TEXT_PRI,
         BorderStyle=1)
    try:
        ctrl.SpecialEffect = 0
    except Exception as e:
        print(f"    [warn] SpecialEffect=0 on listbox failed: {e}")
    _font(ctrl, FONT_BODY[0], FONT_BODY[1])


def style_label(ctrl, text, secondary=False, hint=False, bold=False):
    _set(ctrl, Caption=text, BackStyle=0, ForeColor=COLOR_TEXT_PRI)
    if hint:
        _font(ctrl, FONT_HINT[0], FONT_HINT[1])
        _set(ctrl, ForeColor=COLOR_TEXT_TER)
    elif secondary:
        _font(ctrl, FONT_BODY[0], FONT_BODY[1])
        _set(ctrl, ForeColor=COLOR_TEXT_SEC)
    else:
        _font(ctrl, FONT_BODY[0], FONT_BODY[1], bold)


def style_button_primary(ctrl, caption):
    _set(ctrl,
         Caption=caption,
         BackColor=COLOR_ACCENT, BackStyle=1,
         ForeColor=COLOR_WHITE)
    _font(ctrl, FONT_BUTTON[0], FONT_BUTTON[1], True)


def style_button_secondary(ctrl, caption):
    _set(ctrl,
         Caption=caption,
         BackColor=COLOR_BG_CARD, BackStyle=1,
         ForeColor=COLOR_TEXT_PRI)
    _font(ctrl, FONT_BUTTON[0], FONT_BUTTON[1], False)


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


def set_control_props(ctrl, **kwargs) -> None:
    """Set named properties on a Forms control."""
    for k, v in kwargs.items():
        try:
            setattr(ctrl, k, v)
        except Exception as e:
            print(f"    [warn] Could not set {k}={v!r}: {e}")


# ---------------------------------------------------------------------------
# Form builders
# ---------------------------------------------------------------------------

def build_frm_export(components):
    """Add frmExport to the VBProject — macOS aesthetic.

    Strategy: Import the existing .frm (bypasses VBA name-reservation bug),
    clear all controls, add new styled controls, update VBA code.
    """
    name = "frmExport"
    remove_component(components, name)

    frm_path = SRC / "frmExport.frm"
    print(f"  [import] {name} from {frm_path.name}")
    comp = components.Import(str(frm_path))

    # Form dimensions: 540 x 460
    W, H = 540, 460
    designer = comp.Designer
    try:
        designer.Caption = "Export Snapshot"
    except Exception as e:
        print(f"  [warn] Caption: {e}")

    designer.BackColor = COLOR_BG_WINDOW

    # Clear old controls and rebuild
    clear_controls(designer)
    controls = designer.Controls

    # ---- Title strip + divider ----
    style_form(designer, "Export Snapshot", W)

    # Content starts at Top=64 (48 strip + 1pt divider + 15 padding)
    # Hint label
    lbl_hint = controls.Add(LABEL, "lblHint", True)
    _set(lbl_hint, Top=64, Left=24, Width=492, Height=20)
    style_label(lbl_hint,
                "Snapshot of the active deck. Copy and paste into your LLM tool.",
                hint=True)

    # txtSnapshot – large read-only text area
    txt = controls.Add(TEXTBOX, "txtSnapshot", True)
    _set(txt, Top=92, Left=24, Width=492, Height=240)
    style_textbox(txt, multiline=True, locked=True)

    # Status label
    lbl_status = controls.Add(LABEL, "lblStatus", True)
    _set(lbl_status, Top=344, Left=24, Width=492, Height=20)
    style_label(lbl_status, "", secondary=True)

    # ---- Bottom button row (Top=388) ----
    # macOS convention: secondary buttons left → primary rightmost
    btn_save = controls.Add(CMDBTN, "btnSaveTxt", True)
    _set(btn_save, Top=388, Left=24, Width=140, Height=32)
    style_button_secondary(btn_save, "Save .txt")

    btn_copy = controls.Add(CMDBTN, "btnCopySnapshot", True)
    _set(btn_copy, Top=388, Left=176, Width=140, Height=32)
    style_button_secondary(btn_copy, "Copy snapshot")

    btn_close = controls.Add(CMDBTN, "btnClose", True)
    _set(btn_close, Top=388, Left=332, Width=80, Height=32)
    style_button_secondary(btn_close, "Close")

    btn_copy_tmpl = controls.Add(CMDBTN, "btnCopyWithTemplate", True)
    _set(btn_copy_tmpl, Top=388, Left=420, Width=96, Height=32)
    style_button_primary(btn_copy_tmpl, "Copy + Prompt")

    # Set VBA code
    module = comp.CodeModule
    if module.CountOfLines > 0:
        module.DeleteLines(1, module.CountOfLines)
    module.AddFromString(FRM_EXPORT_CODE)

    print(f"  [ok] {name} built")
    return comp


def build_frm_execute(components):
    """Add frmExecute to the VBProject — macOS aesthetic."""
    name = "frmExecute"
    remove_component(components, name)

    frm_path = SRC / "frmExecute.frm"
    print(f"  [import] {name} from {frm_path.name}")
    comp = components.Import(str(frm_path))

    W, H = 600, 540
    designer = comp.Designer
    try:
        designer.Caption = "Execute Instructions"
    except Exception as e:
        print(f"  [warn] Caption: {e}")

    designer.BackColor = COLOR_BG_WINDOW

    # Clear old controls and rebuild
    clear_controls(designer)
    controls = designer.Controls

    # ---- Title strip + divider ----
    style_form(designer, "Execute Instructions", W)

    # Hint label
    lbl_hint = controls.Add(LABEL, "lblHint", True)
    _set(lbl_hint, Top=64, Left=24, Width=552, Height=20)
    style_label(lbl_hint,
                "Paste the instructions JSON from your LLM. Review parsed actions, then Apply.",
                hint=True)

    # txtInstructions
    txt = controls.Add(TEXTBOX, "txtInstructions", True)
    _set(txt, Top=92, Left=24, Width=552, Height=120)
    style_textbox(txt, multiline=True)

    # btnParse
    btn_parse = controls.Add(CMDBTN, "btnParse", True)
    _set(btn_parse, Top=224, Left=24, Width=100, Height=32)
    style_button_secondary(btn_parse, "Parse")

    # Section label "Parsed Actions"
    lbl_section = controls.Add(LABEL, "lblParsedActions", True)
    _set(lbl_section, Top=268, Left=24, Width=552, Height=20)
    style_label(lbl_section, "Parsed Actions", bold=True)

    # lstActions
    lst = controls.Add(LISTBOX, "lstActions", True)
    _set(lst, Top=292, Left=24, Width=552, Height=144)
    style_listbox(lst)

    # Status label
    lbl_status = controls.Add(LABEL, "lblStatus", True)
    _set(lbl_status, Top=448, Left=24, Width=552, Height=36)
    style_label(lbl_status, "", secondary=True)

    # ---- Bottom button row (Top=492) ----
    btn_cancel = controls.Add(CMDBTN, "btnCancel", True)
    _set(btn_cancel, Top=492, Left=412, Width=80, Height=32)
    style_button_secondary(btn_cancel, "Cancel")

    btn_apply = controls.Add(CMDBTN, "btnApply", True)
    _set(btn_apply, Top=492, Left=500, Width=76, Height=32, Enabled=False)
    style_button_primary(btn_apply, "Apply")

    # Set VBA code
    module = comp.CodeModule
    if module.CountOfLines > 0:
        module.DeleteLines(1, module.CountOfLines)
    module.AddFromString(FRM_EXECUTE_CODE)

    print(f"  [ok] {name} built")
    return comp


def build_frm_import_slides(components):
    """Add frmImportSlides to the VBProject — macOS aesthetic."""
    name = "frmImportSlides"
    remove_component(components, name)

    frm_path = SRC / "frmImportSlides.frm"
    print(f"  [import] {name} from {frm_path.name}")
    comp = components.Import(str(frm_path))

    W, H = 480, 360
    designer = comp.Designer
    try:
        designer.Caption = "Import Slides"
    except Exception as e:
        print(f"  [warn] Caption: {e}")

    designer.BackColor = COLOR_BG_WINDOW

    # Clear old controls and rebuild
    clear_controls(designer)
    controls = designer.Controls

    # ---- Title strip + divider ----
    style_form(designer, "Import Slides", W)

    # Source deck row
    lbl_path = controls.Add(LABEL, "lblPath", True)
    _set(lbl_path, Top=76, Left=24, Width=100, Height=22)
    style_label(lbl_path, "Source deck")

    txt_path = controls.Add(TEXTBOX, "txtPath", True)
    _set(txt_path, Top=76, Left=132, Width=256, Height=22)
    style_textbox(txt_path, locked=True)

    btn_browse = controls.Add(CMDBTN, "btnBrowse", True)
    _set(btn_browse, Top=76, Left=396, Width=60, Height=22)
    style_button_secondary(btn_browse, "Browse...")

    # Slide range row
    lbl_range = controls.Add(LABEL, "lblRange", True)
    _set(lbl_range, Top=116, Left=24, Width=224, Height=22)
    style_label(lbl_range, "Slide range (e.g. 1-3,5,7-9)")

    txt_range = controls.Add(TEXTBOX, "txtRange", True)
    _set(txt_range, Top=116, Left=252, Width=204, Height=22)
    style_textbox(txt_range)

    # Position row
    lbl_pos = controls.Add(LABEL, "lblPosition", True)
    _set(lbl_pos, Top=156, Left=24, Width=224, Height=22)
    style_label(lbl_pos, "Insert at position")

    txt_pos = controls.Add(TEXTBOX, "txtPosition", True)
    _set(txt_pos, Top=156, Left=252, Width=60, Height=22)
    style_textbox(txt_pos)

    # Status label
    lbl_status = controls.Add(LABEL, "lblStatus", True)
    _set(lbl_status, Top=200, Left=24, Width=432, Height=60)
    style_label(lbl_status, "", secondary=True)

    # ---- Bottom button row (Top=296) ----
    btn_cancel = controls.Add(CMDBTN, "btnCancel", True)
    _set(btn_cancel, Top=296, Left=312, Width=76, Height=32)
    style_button_secondary(btn_cancel, "Cancel")

    btn_import = controls.Add(CMDBTN, "btnImport", True)
    _set(btn_import, Top=296, Left=392, Width=72, Height=32, Enabled=False)
    style_button_primary(btn_import, "Import")

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

def main() -> int:
    if not CARRIER.exists():
        print(f"ERROR: {CARRIER} not found. Run tools/build_carrier.py first.")
        return 1

    print(f"[build_forms] Opening carrier: {CARRIER}")
    import win32com.client

    # Single-pass: open, remove old forms, import existing .frm (bypasses VBA
    # name-reservation bug), clear controls, rebuild with macOS style, save, export.
    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.Visible = True

    try:
        pres = app.Presentations.Open(str(CARRIER), WithWindow=True)
        try:
            project = pres.VBProject
            components = project.VBComponents

            comp_export        = build_frm_export(components)
            comp_execute       = build_frm_execute(components)
            comp_import_slides = build_frm_import_slides(components)

            pres.Save()
            print("[saved] carrier saved")

            # Export .frm + .frx to src/
            export_path_export        = str(SRC / "frmExport.frm")
            export_path_execute       = str(SRC / "frmExecute.frm")
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
                        client_width=10800,    # 540pt * 20
                        client_height=9200)    # 460pt * 20
    _fix_frm_dimensions(SRC / "frmExecute.frm",
                        client_width=12000,    # 600pt * 20
                        client_height=10800)   # 540pt * 20
    _fix_frm_dimensions(SRC / "frmImportSlides.frm",
                        client_width=9600,     # 480pt * 20
                        client_height=7200)    # 360pt * 20

    # Caption fix: COM Designer.Caption is not always set via setattr
    _fix_frm_caption(SRC / "frmExport.frm",        "Export Snapshot")
    _fix_frm_caption(SRC / "frmExecute.frm",        "Execute Instructions")
    _fix_frm_caption(SRC / "frmImportSlides.frm",   "Import Slides")

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
