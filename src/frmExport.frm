VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExport 
   Caption         =   "Export Snapshot"
   ClientHeight    =   9200
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   10800
   OleObjectBlob   =   "frmExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    s = s & "  {""type"":""insert_picture"",""slide"":1,""path"":""C:\\path\\to\\img.png"",""pos"":{""left"":50,""top"":50,""width"":200,""height"":150}}" & vbCrLf
    s = s & "  {""type"":""replace_picture"",""slide"":1,""shape_id"":7,""path"":""C:\\path\\to\\new.png""}" & vbCrLf & vbCrLf

    s = s & "SLIDE STRUCTURE:" & vbCrLf
    s = s & "  {""type"":""move_slide"",""from"":3,""to"":1}" & vbCrLf
    s = s & "  {""type"":""extract_slides"",""slide_indices"":[1,3,5],""output_path"":""C:\\path\\out.pptx""}" & vbCrLf
    s = s & "  {""type"":""import_slides_from_deck"",""source_path"":""C:\\path\\other.pptx"",""slide_indices"":[1,2],""target_position"":3}" & vbCrLf & vbCrLf

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

