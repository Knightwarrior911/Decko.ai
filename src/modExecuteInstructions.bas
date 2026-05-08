Attribute VB_Name = "modExecuteInstructions"
Option Explicit

' Parse instructions JSON, validate each action, run backup, dispatch
' valid actions, log per-action result. Returns a summary string.
Public Function ExecuteFromString(jsonText As String) As String
    Dim parsed As Object
    On Error Resume Next
    Set parsed = modJSON.ParseJson(jsonText)
    If Err.Number <> 0 Then
        ExecuteFromString = "ERROR: invalid JSON: " & Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    If Not parsed.Exists("actions") Then
        ExecuteFromString = "ERROR: missing top-level 'actions' array"
        Exit Function
    End If

    Dim actions As Object
    Set actions = parsed("actions")

    Dim deckPath As String
    deckPath = ActivePresentation.FullName

    Dim backupPath As String
    On Error Resume Next
    backupPath = modBackup.BackupActiveDeck()
    If Err.Number <> 0 Then
        ExecuteFromString = "ERROR: backup failed: " & Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    Dim applied As Long, skipped As Long
    applied = 0: skipped = 0

    Dim i As Long
    For i = 1 To actions.Count
        Dim act As Object
        Set act = actions(i)

        Dim invalidReason As String
        invalidReason = ValidateAction(act)

        Dim paramsJson As String
        paramsJson = modJSON.ConvertToJson(act)

        If Len(invalidReason) > 0 Then
            modBackup.LogAction deckPath, GetStr(act, "type"), _
                                GetVar(act, "slide"), GetVar(act, "shape_id"), _
                                paramsJson, "skipped", invalidReason
            skipped = skipped + 1
        Else
            On Error Resume Next
            DispatchAction act
            If Err.Number <> 0 Then
                modBackup.LogAction deckPath, GetStr(act, "type"), _
                                    GetVar(act, "slide"), GetVar(act, "shape_id"), _
                                    paramsJson, "error", Err.Description
                Err.Clear
                skipped = skipped + 1
            Else
                modBackup.LogAction deckPath, GetStr(act, "type"), _
                                    GetVar(act, "slide"), GetVar(act, "shape_id"), _
                                    paramsJson, "ok", ""
                applied = applied + 1
            End If
            On Error GoTo 0
        End If
    Next i

    On Error Resume Next
    ActivePresentation.Save
    On Error GoTo 0

    ExecuteFromString = applied & " applied, " & skipped & " skipped. " & _
                        "Log: " & deckPath & ".action_log.jsonl. " & _
                        "Backup: " & backupPath
End Function

' Returns "" if valid, else a reason string.
Private Function ValidateAction(act As Object) As String
    If Not act.Exists("type") Then
        ValidateAction = "missing_field: type"
        Exit Function
    End If
    Dim t As String: t = act("type")

    Select Case t
        Case "set_text"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_font_size"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_font_bold", "set_font_italic"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_font_color", "set_fill_color"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "move_shape"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "left", "top"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "resize_shape"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "width", "height"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "delete_shape"
            ValidateAction = RequireFields(act, Array("slide", "shape_id"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "add_slide"
            ValidateAction = RequireFields(act, Array("position", "layout_index"))
        Case "delete_slide", "duplicate_slide"
            ValidateAction = RequireFields(act, Array("slide"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "set_cell_text"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "swap_table_columns"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "col_a", "col_b"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "swap_table_rows"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row_a", "row_b"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "add_paragraph"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "after_paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "delete_paragraph"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_bullet_style"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_indent_level"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_paragraph_font_size"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_paragraph_font_color"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "find_replace_text"
            ValidateAction = RequireFields(act, Array("scope", "find", "replace"))
        Case "set_paragraph_text"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "align_shapes"
            ValidateAction = RequireFields(act, Array("slide", "shape_ids", "anchor"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "distribute_horizontal", "distribute_vertical"
            ValidateAction = RequireFields(act, Array("slide", "shape_ids"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "tile_grid"
            ValidateAction = RequireFields(act, Array("slide", "shape_ids", "cols", "gap_pt"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "fit_to_slide_margins"
            ValidateAction = RequireFields(act, Array("slide", "shape_id"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "add_line"
            ValidateAction = RequireFields(act, Array("slide", "x1", "y1", "x2", "y2", "color", "weight_pt"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "add_shape"
            ValidateAction = RequireFields(act, Array("slide", "kind", "pos"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "set_shape_kind"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "kind"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "clear_slide"
            ValidateAction = RequireFields(act, Array("slide"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "move_shape_relative"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "dx_pt", "dy_pt"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "recolor_fill_match", "recolor_font_match"
            ValidateAction = RequireFields(act, Array("scope", "from", "to"))
        Case "delete_shapes_match"
            ValidateAction = RequireFields(act, Array("scope"))
        Case "set_speaker_notes", "append_speaker_notes"
            ValidateAction = RequireFields(act, Array("slide", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "insert_picture"
            ValidateAction = RequireFields(act, Array("slide", "path", "pos"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "replace_picture"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "path"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "move_slide"
            ValidateAction = RequireFields(act, Array("from", "to"))
        Case "extract_slides"
            ValidateAction = RequireFields(act, Array("slide_indices", "output_path"))
        Case "import_slides_from_deck"
            ValidateAction = RequireFields(act, Array("source_path", "slide_indices", "target_position"))
        Case "add_table_row"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "after_row"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "delete_table_row"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "add_table_col"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "after_col"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "delete_table_col"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "col"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "merge_cells"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row_a", "col_a", "row_b", "col_b"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "group_shapes"
            ValidateAction = RequireFields(act, Array("slide", "shape_ids"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "ungroup"
            ValidateAction = RequireFields(act, Array("slide", "shape_id"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "add_connector"
            ValidateAction = RequireFields(act, Array("slide", "from_shape_id", "to_shape_id", "kind"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "set_chart_type"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_chart_title"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_chart_axis_title"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "axis", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_chart_legend_position"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_series_color"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "series_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case Else
            ValidateAction = "unknown_type: " & t
    End Select
End Function

Private Function RequireFields(act As Object, fields As Variant) As String
    Dim i As Long
    For i = LBound(fields) To UBound(fields)
        If Not act.Exists(fields(i)) Then
            RequireFields = "missing_field: " & fields(i)
            Exit Function
        End If
    Next i
    RequireFields = ""
End Function

Private Function ValidateSlide(act As Object) As String
    Dim n As Long: n = CLng(act("slide"))
    If n < 1 Or n > ActivePresentation.Slides.Count Then
        ValidateSlide = "slide_out_of_range"
    End If
End Function

Private Function ValidateShape(act As Object) As String
    Dim slideErr As String: slideErr = ValidateSlide(act)
    If Len(slideErr) > 0 Then
        ValidateShape = slideErr
        Exit Function
    End If
    Dim sh As Shape
    Set sh = modActions.FindShape(CLng(act("slide")), CLng(act("shape_id")))
    If sh Is Nothing Then ValidateShape = "shape_not_found"
End Function

Private Sub DispatchAction(act As Object)
    Dim t As String: t = act("type")
    Select Case t
        Case "set_text"
            modActions.Do_set_text CLng(act("slide")), CLng(act("shape_id")), CStr(act("value"))
        Case "set_font_size"
            modActions.Do_set_font_size CLng(act("slide")), CLng(act("shape_id")), CLng(act("value"))
        Case "set_font_bold"
            modActions.Do_set_font_bold CLng(act("slide")), CLng(act("shape_id")), CBool(act("value"))
        Case "set_font_italic"
            modActions.Do_set_font_italic CLng(act("slide")), CLng(act("shape_id")), CBool(act("value"))
        Case "set_font_color"
            modActions.Do_set_font_color CLng(act("slide")), CLng(act("shape_id")), CStr(act("value"))
        Case "set_fill_color"
            modActions.Do_set_fill_color CLng(act("slide")), CLng(act("shape_id")), CStr(act("value"))
        Case "move_shape"
            modActions.Do_move_shape CLng(act("slide")), CLng(act("shape_id")), _
                                     CSng(act("left")), CSng(act("top"))
        Case "resize_shape"
            modActions.Do_resize_shape CLng(act("slide")), CLng(act("shape_id")), _
                                       CSng(act("width")), CSng(act("height"))
        Case "delete_shape"
            modActions.Do_delete_shape CLng(act("slide")), CLng(act("shape_id"))
        Case "add_slide"
            modActions.Do_add_slide CLng(act("position")), CLng(act("layout_index"))
        Case "delete_slide"
            modActions.Do_delete_slide CLng(act("slide"))
        Case "duplicate_slide"
            modActions.Do_duplicate_slide CLng(act("slide"))
        Case "set_cell_text"
            modActions.Do_set_cell_text CLng(act("slide")), CLng(act("shape_id")), _
                                        CLng(act("row")), CLng(act("col")), CStr(act("value"))
        Case "swap_table_columns"
            modActions.Do_swap_table_columns CLng(act("slide")), CLng(act("shape_id")), _
                                             CLng(act("col_a")), CLng(act("col_b"))
        Case "swap_table_rows"
            modActions.Do_swap_table_rows CLng(act("slide")), CLng(act("shape_id")), _
                                          CLng(act("row_a")), CLng(act("row_b"))
        Case "set_paragraph_text"
            modActionsText.Do_set_paragraph_text CLng(act("slide")), CLng(act("shape_id")), _
                                                 CLng(act("paragraph_index")), CStr(act("value"))
        Case "add_paragraph"
            modActionsText.Do_add_paragraph CLng(act("slide")), CLng(act("shape_id")), _
                                            CLng(act("after_paragraph_index")), CStr(act("value"))
        Case "delete_paragraph"
            modActionsText.Do_delete_paragraph CLng(act("slide")), CLng(act("shape_id")), _
                                               CLng(act("paragraph_index"))
        Case "set_bullet_style"
            modActionsText.Do_set_bullet_style CLng(act("slide")), CLng(act("shape_id")), _
                                               CLng(act("paragraph_index")), CStr(act("value"))
        Case "set_indent_level"
            modActionsText.Do_set_indent_level CLng(act("slide")), CLng(act("shape_id")), _
                                               CLng(act("paragraph_index")), CLng(act("value"))
        Case "set_paragraph_font_size"
            modActionsText.Do_set_paragraph_font_size CLng(act("slide")), CLng(act("shape_id")), _
                                                       CLng(act("paragraph_index")), CLng(act("value"))
        Case "set_paragraph_font_color"
            modActionsText.Do_set_paragraph_font_color CLng(act("slide")), CLng(act("shape_id")), _
                                                        CLng(act("paragraph_index")), CStr(act("value"))
        Case "find_replace_text"
            modActionsText.Do_find_replace_text CStr(act("scope")), CStr(act("find")), CStr(act("replace"))
        Case "align_shapes"
            modActionsLayout.Do_align_shapes CLng(act("slide")), act("shape_ids"), CStr(act("anchor"))
        Case "distribute_horizontal"
            modActionsLayout.Do_distribute_horizontal CLng(act("slide")), act("shape_ids")
        Case "distribute_vertical"
            modActionsLayout.Do_distribute_vertical CLng(act("slide")), act("shape_ids")
        Case "tile_grid"
            modActionsLayout.Do_tile_grid CLng(act("slide")), act("shape_ids"), _
                                          CLng(act("cols")), CSng(act("gap_pt"))
        Case "fit_to_slide_margins"
            Dim m As Single: m = 36.0
            If act.Exists("margin_pt") Then m = CSng(act("margin_pt"))
            modActionsLayout.Do_fit_to_slide_margins CLng(act("slide")), CLng(act("shape_id")), m
        Case "add_line"
            modActionsLayout.Do_add_line CLng(act("slide")), CSng(act("x1")), CSng(act("y1")), _
                                         CSng(act("x2")), CSng(act("y2")), _
                                         CStr(act("color")), CSng(act("weight_pt"))
        Case "add_shape"
            Dim posDict As Object: Set posDict = act("pos")
            Dim fh As String: fh = ""
            Dim shex As String: shex = ""
            Dim swt As Single: swt = 1.0
            If act.Exists("fill") Then If Not IsNull(act("fill")) Then fh = CStr(act("fill"))
            If act.Exists("stroke") Then If Not IsNull(act("stroke")) Then shex = CStr(act("stroke"))
            If act.Exists("stroke_weight_pt") Then swt = CSng(act("stroke_weight_pt"))
            modActionsLayout.Do_add_shape CLng(act("slide")), CStr(act("kind")), _
                                          CSng(posDict("left")), CSng(posDict("top")), _
                                          CSng(posDict("width")), CSng(posDict("height")), _
                                          fh, shex, swt
        Case "set_shape_kind"
            modActionsLayout.Do_set_shape_kind CLng(act("slide")), CLng(act("shape_id")), CStr(act("kind"))
        Case "clear_slide"
            Dim keep As Variant
            If act.Exists("keep_shape_ids") Then
                keep = act("keep_shape_ids")
            Else
                keep = Array()
            End If
            modActionsLayout.Do_clear_slide CLng(act("slide")), keep
        Case "move_shape_relative"
            modActionsLayout.Do_move_shape_relative CLng(act("slide")), CLng(act("shape_id")), _
                                                    CSng(act("dx_pt")), CSng(act("dy_pt"))
        Case "recolor_fill_match"
            modActionsLayout.Do_recolor_fill_match CStr(act("scope")), CStr(act("from")), CStr(act("to"))
        Case "recolor_font_match"
            modActionsLayout.Do_recolor_font_match CStr(act("scope")), CStr(act("from")), CStr(act("to"))
        Case "delete_shapes_match"
            Dim kf As String, ff As String, tc As String
            kf = "" : ff = "" : tc = ""
            If act.Exists("kind") Then kf = CStr(act("kind"))
            If act.Exists("fill") Then ff = CStr(act("fill"))
            If act.Exists("text_contains") Then tc = CStr(act("text_contains"))
            modActionsLayout.Do_delete_shapes_match CStr(act("scope")), kf, ff, tc
        Case "set_speaker_notes"
            modActions.Do_set_speaker_notes CLng(act("slide")), CStr(act("value"))
        Case "append_speaker_notes"
            modActions.Do_append_speaker_notes CLng(act("slide")), CStr(act("value"))
        Case "insert_picture"
            Dim ipos As Object: Set ipos = act("pos")
            modActionsImage.Do_insert_picture CLng(act("slide")), CStr(act("path")), _
                                              CSng(ipos("left")), CSng(ipos("top")), _
                                              CSng(ipos("width")), CSng(ipos("height"))
        Case "replace_picture"
            modActionsImage.Do_replace_picture CLng(act("slide")), CLng(act("shape_id")), CStr(act("path"))
        Case "move_slide"
            modActionsSlide.Do_move_slide CLng(act("from")), CLng(act("to"))
        Case "extract_slides"
            modActionsSlide.Do_extract_slides act("slide_indices"), CStr(act("output_path"))
        Case "import_slides_from_deck"
            modActionsSlide.Do_import_slides_from_deck CStr(act("source_path")), _
                                                       act("slide_indices"), _
                                                       CLng(act("target_position"))
        Case "add_table_row"
            modActionsTable.Do_add_table_row CLng(act("slide")), CLng(act("shape_id")), CLng(act("after_row"))
        Case "delete_table_row"
            modActionsTable.Do_delete_table_row CLng(act("slide")), CLng(act("shape_id")), CLng(act("row"))
        Case "add_table_col"
            modActionsTable.Do_add_table_col CLng(act("slide")), CLng(act("shape_id")), CLng(act("after_col"))
        Case "delete_table_col"
            modActionsTable.Do_delete_table_col CLng(act("slide")), CLng(act("shape_id")), CLng(act("col"))
        Case "merge_cells"
            modActionsTable.Do_merge_cells CLng(act("slide")), CLng(act("shape_id")), _
                                           CLng(act("row_a")), CLng(act("col_a")), _
                                           CLng(act("row_b")), CLng(act("col_b"))
        Case "group_shapes"
            modActionsGroup.Do_group_shapes CLng(act("slide")), act("shape_ids")
        Case "ungroup"
            modActionsGroup.Do_ungroup CLng(act("slide")), CLng(act("shape_id"))
        Case "add_connector"
            Dim ae As String, cc As String, cw As Single
            ae = "filled" : cc = "#000000" : cw = 1.0
            If act.Exists("arrow_end") Then ae = CStr(act("arrow_end"))
            If act.Exists("color") Then cc = CStr(act("color"))
            If act.Exists("weight_pt") Then cw = CSng(act("weight_pt"))
            modActionsConnector.Do_add_connector CLng(act("slide")), _
                                                 CLng(act("from_shape_id")), _
                                                 CLng(act("to_shape_id")), _
                                                 CStr(act("kind")), ae, cc, cw
        Case "set_chart_type"
            modActionsChart.Do_set_chart_type CLng(act("slide")), CLng(act("shape_id")), CStr(act("value"))
        Case "set_chart_title"
            Dim cte As Boolean: cte = True
            If act.Exists("enabled") Then cte = CBool(act("enabled"))
            modActionsChart.Do_set_chart_title CLng(act("slide")), CLng(act("shape_id")), _
                                               CStr(act("value")), cte
        Case "set_chart_axis_title"
            modActionsChart.Do_set_chart_axis_title CLng(act("slide")), CLng(act("shape_id")), _
                                                    CStr(act("axis")), CStr(act("value"))
        Case "set_chart_legend_position"
            modActionsChart.Do_set_chart_legend_position CLng(act("slide")), CLng(act("shape_id")), _
                                                          CStr(act("value"))
        Case "set_series_color"
            modActionsChart.Do_set_series_color CLng(act("slide")), CLng(act("shape_id")), _
                                                CLng(act("series_index")), CStr(act("value"))
    End Select
End Sub

Private Function GetStr(d As Object, key As String) As String
    If d.Exists(key) Then GetStr = CStr(d(key))
End Function

Private Function GetVar(d As Object, key As String) As Variant
    If d.Exists(key) Then
        GetVar = d(key)
    Else
        GetVar = Null
    End If
End Function

' Public wrapper around private validation, for the dry-run preview UI.
Public Function PreviewValidate(act As Object) As String
    PreviewValidate = ValidateAction(act)
End Function
