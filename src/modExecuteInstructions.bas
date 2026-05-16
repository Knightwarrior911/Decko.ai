Attribute VB_Name = "modExecuteInstructions"
Option Explicit

' Parse instructions JSON, validate each action, dispatch valid actions,
' log per-action result. Returns a summary string. NO auto-backup, NO save.
Public Function ExecuteFromString(jsonText As String) As String
    ' LLMs frequently disobey "no prose" instructions and prepend/append text.
    ' Sanitize before parsing so the user does not have to clean the output by hand.
    Dim cleaned As String: cleaned = SanitizeJsonInput(jsonText)
    Dim parsed As Object
    On Error Resume Next
    Set parsed = modJSON.ParseJson(cleaned)
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
    Set actions = ReorderForRunIndexSafety(actions)

    Dim deckPath As String
    deckPath = ActivePresentation.FullName

    Dim applied As Long, skipped As Long
    applied = 0: skipped = 0
    ' Per-failure report appended to the returned summary so the caller
    ' sees exactly which action failed and why -- not just a count, and
    ' never a silent swallow. Index is the executed (post-reorder) order.
    Dim failReport As String: failReport = ""

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
            failReport = failReport & vbCrLf & "action #" & i & " " & _
                         GetStr(act, "type") & ": " & invalidReason
        Else
            On Error Resume Next
            DispatchAction act
            If Err.Number <> 0 Then
                Dim errDesc As String: errDesc = Err.Description
                modBackup.LogAction deckPath, GetStr(act, "type"), _
                                    GetVar(act, "slide"), GetVar(act, "shape_id"), _
                                    paramsJson, "error", errDesc
                Err.Clear
                skipped = skipped + 1
                failReport = failReport & vbCrLf & "action #" & i & " " & _
                             GetStr(act, "type") & ": " & errDesc
            Else
                modBackup.LogAction deckPath, GetStr(act, "type"), _
                                    GetVar(act, "slide"), GetVar(act, "shape_id"), _
                                    paramsJson, "ok", ""
                applied = applied + 1
            End If
            On Error GoTo 0
        End If
    Next i

    ' --- Automatic verification sweep ----------------------------------
    ' Runs by default; caller may set "verify_after": false at the top level
    ' to skip (e.g. when chaining batches and only the last should verify),
    ' or "verify_scope": "slide:N" / "deck" to scope the sweep.
    Dim verifyOn As Boolean: verifyOn = True
    If parsed.Exists("verify_after") Then verifyOn = modActions.ToBool(parsed("verify_after"))
    Dim verifyScope As String: verifyScope = "deck"
    If parsed.Exists("verify_scope") Then verifyScope = CStr(parsed("verify_scope"))

    Dim verifyMsg As String: verifyMsg = ""
    Dim warningsJson As String: warningsJson = ""
    If verifyOn Then
        Dim warnings As Collection
        Set warnings = modVerify.RunVerificationLoop(verifyScope, 100)
        verifyMsg = " | " & modVerify.FormatWarningsSummary(warnings)
        warningsJson = modVerify.WarningsToJson(warnings)
        If warnings.Count > 0 Then
            ' Persist warnings to a sidecar file so the LLM / user can read details.
            WriteWarningsSidecar deckPath, warningsJson
            verifyMsg = verifyMsg & " (details: " & deckPath & ".warnings.json)"
        End If
    End If

    ExecuteFromString = applied & " applied, " & skipped & " skipped. " & _
                        "Log: " & deckPath & ".action_log.jsonl" & verifyMsg
    If Len(failReport) > 0 Then
        ExecuteFromString = ExecuteFromString & vbCrLf & _
                            "FAILURES (" & skipped & "):" & failReport
    End If
End Function

' Write the warnings JSON to <deckPath>.warnings.json so a human or downstream
' agent can read the full payload without parsing the action_log.
Private Sub WriteWarningsSidecar(deckPath As String, warningsJson As String)
    On Error Resume Next
    Dim path As String: path = deckPath & ".warnings.json"
    Dim fnum As Integer: fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, "{""warnings"":" & warningsJson & "}"
    Close #fnum
    On Error GoTo 0
End Sub

' =============================================================================
' Plan preview: turn a (possibly messy) action batch into a numbered,
' plain-language description of exactly what Apply WILL do -- WITHOUT
' mutating the deck or calling any Do_* handler. Mirrors the executor's
' sanitize -> parse -> run-order so the preview matches reality. The
' executor has no undo and no auto-backup, so this is the user's only
' chance to catch intent drift before it is destructive.
' =============================================================================
Public Function BuildActionPlanSummary(jsonText As String) As String
    Dim cleaned As String: cleaned = SanitizeJsonInput(jsonText)
    Dim parsed As Object
    On Error Resume Next
    Set parsed = modJSON.ParseJson(cleaned)
    If Err.Number <> 0 Then
        BuildActionPlanSummary = "ERROR: invalid JSON: " & Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    If parsed Is Nothing Then
        BuildActionPlanSummary = "ERROR: invalid JSON"
        Exit Function
    End If
    If Not parsed.Exists("actions") Then
        BuildActionPlanSummary = "ERROR: missing top-level 'actions' array"
        Exit Function
    End If

    Dim actions As Object
    Set actions = parsed("actions")
    Set actions = ReorderForRunIndexSafety(actions)

    If actions.Count = 0 Then
        BuildActionPlanSummary = "(no actions)"
        Exit Function
    End If

    Dim out As String
    Dim i As Long
    For i = 1 To actions.Count
        Dim act As Object: Set act = actions(i)
        If i > 1 Then out = out & vbCrLf
        out = out & CStr(i) & ". " & DescribeAction(act)
    Next i
    BuildActionPlanSummary = out
End Function

' One precise sentence for a single action. Specific per type; never a
' bare "action: <type>" fallback. Unknown -> explicit UNKNOWN ACTION line.
Private Function DescribeAction(act As Object) As String
    Dim t As String: t = GetStr(act, "type")
    If Len(t) = 0 Then
        DescribeAction = "UNKNOWN ACTION (missing type)"
        Exit Function
    End If

    Dim sp As String: sp = PlanSlidePrefix(act)
    Dim tg As String: tg = PlanTarget(act)

    Select Case t
        ' ---- text / content ----
        Case "set_text"
            DescribeAction = sp & "set text" & PlanOn(tg) & " -> " & PlanQ(GetStr(act, "text"))
        Case "set_cell_text", "append_cell_text", "clear_cell_text", "set_cell"
            DescribeAction = sp & PlanWords(t) & " in cell (" & PlanV(act, "row") & "," & PlanV(act, "col") & ")" & PlanOn(tg)
        Case "set_paragraph_text", "set_cell_paragraph_text"
            DescribeAction = sp & PlanWords(t) & PlanOn(tg) & " -> " & PlanQ(GetStr(act, "text"))
        Case "set_run_text", "add_run"
            DescribeAction = sp & PlanWords(t) & PlanOn(tg) & " -> " & PlanQ(GetStr(act, "text"))
        Case "find_replace_text", "find_replace_regex"
            DescribeAction = sp & PlanWords(t) & " -> replace " & PlanQ(GetStr(act, "find")) & " with " & PlanQ(GetStr(act, "replace"))
        Case "set_speaker_notes", "append_speaker_notes", "clear_speaker_notes"
            DescribeAction = sp & PlanWords(t)

        ' ---- font / run / paragraph / cell formatting ----
        Case "set_font_size", "set_paragraph_font_size", "set_run_font_size", _
             "set_cell_font_size", "set_row_font_size", "set_column_font_size", _
             "set_table_font_size", "set_cell_paragraph_font_size", "set_notes_font_size"
            DescribeAction = sp & PlanWords(t) & PlanOn(tg) & " -> " & PlanV(act, "size") & "pt"
        Case "set_font_color", "set_paragraph_font_color", "set_run_font_color", _
             "set_cell_font_color", "set_row_font_color", "set_column_font_color", _
             "set_table_font_color", "set_cell_paragraph_font_color", "set_notes_font_color"
            DescribeAction = sp & PlanWords(t) & PlanOn(tg) & " -> " & PlanColor(act)
        Case "set_paragraph_font_name", "set_run_font_name", "set_cell_font_name", _
             "set_table_font_name", "set_notes_font_name", "set_theme_font"
            DescribeAction = sp & PlanWords(t) & PlanOn(tg) & " -> " & PlanQ(GetStr(act, "name"))
        Case "set_font_bold", "set_paragraph_bold", "set_run_bold", "set_cell_font_bold", _
             "set_row_font_bold", "set_column_font_bold", "set_cell_paragraph_bold", "set_notes_font_bold", _
             "set_font_italic", "set_paragraph_italic", "set_run_italic", "set_cell_font_italic", _
             "set_cell_paragraph_italic", "set_notes_font_italic", _
             "set_run_underline", "set_paragraph_underline", "set_cell_font_underline", _
             "set_run_strikethrough", "set_run_subscript", "set_run_superscript"
            DescribeAction = sp & PlanWords(t) & PlanOn(tg) & " -> " & PlanBool(act)
        Case "set_run_highlight", "set_run_kerning", "set_run_baseline_offset", _
             "set_paragraph_alignment", "set_cell_paragraph_alignment", "set_paragraph_line_spacing", _
             "set_paragraph_space_before", "set_paragraph_space_after", "clear_paragraph_formatting", _
             "set_bullet_style", "set_cell_bullet_style", "set_indent_level", "set_cell_indent_level", _
             "set_bullet_start_number", "set_run_hyperlink"
            DescribeAction = sp & PlanWords(t) & PlanOn(tg) & PlanArrowVal(act)

        ' ---- geometry / layout ----
        Case "move_shape", "set_pos", "move_shape_relative"
            DescribeAction = sp & PlanWords(t) & PlanOn(tg) & " -> left " & PlanV(act, "left") & ", top " & PlanV(act, "top")
        Case "resize_shape"
            DescribeAction = sp & "resize" & PlanOn(tg) & " -> " & PlanV(act, "width") & " x " & PlanV(act, "height")
        Case "rotate_shape", "set_3d_rotation"
            DescribeAction = sp & PlanWords(t) & PlanOn(tg) & " -> angle " & PlanV(act, "angle")
        Case "delete_shape", "duplicate_shape", "group_shapes", "ungroup", "flip_shape", _
             "lock_aspect_ratio", "z_order", "set_shape_name", "set_shape_alt_text", _
             "set_shape_kind", "set_shape_visible", "set_shape_hyperlink", "set_shape_adjustment", _
             "snap_to_grid", "align_to_slide_center", "nudge", "fit_to_content", "set_shape_picture_fill"
            DescribeAction = sp & PlanWords(t) & PlanOn(tg)
        Case "align_shapes", "distribute_horizontal", "distribute_vertical", "tile_grid", _
             "fit_to_slide_margins", "match_size", "uniform_size", "smart_spacing", _
             "equalize_spacing", "match_position", "swap_positions", "group_by_overlap"
            DescribeAction = sp & PlanWords(t) & " on selected shapes"
        Case "add_line", "add_shape", "add_text_box", "add_connector", "reconnect_connector"
            DescribeAction = sp & PlanWords(t) & PlanArrowVal(act)
        Case "clear_slide"
            DescribeAction = sp & "clear all shapes on the slide"

        ' ---- effects ----
        Case "set_line_color", "set_line_weight", "set_line_style", "set_shadow", "set_glow", _
             "set_reflection", "set_transparency", "set_gradient_fill", "set_3d_bevel", _
             "apply_preset_effect", "set_soft_edge", "set_fill_color", "clear_fill", "clear_line", _
             "set_fill_visible", "set_line_visible", "clear_shadow", "clear_glow", "clear_reflection", _
             "clear_all_effects"
            DescribeAction = sp & PlanWords(t) & PlanOn(tg) & PlanArrowVal(act)

        ' ---- pictures ----
        Case "insert_picture", "replace_picture", "insert_icon", "crop_picture", "recolor_picture", _
             "set_brightness", "set_contrast", "apply_picture_artistic_effect", "reset_picture", _
             "download_image", "fetch_page_images", "open_image_picker", "build_image_picker_slide", _
             "bulk_insert_image", "build_image_grid_table"
            DescribeAction = sp & PlanWords(t) & PlanOn(tg)

        ' ---- slides / deck / sections ----
        Case "add_slide", "delete_slide", "duplicate_slide", "move_slide", "set_slide_hidden", _
             "set_slide_name", "set_slide_transition", "change_slide_layout", "set_slide_background_color", _
             "insert_slide_number", "apply_layout_to_slides", "set_slide_size", "apply_theme", _
             "extract_slides", "import_slides_from_deck", "bulk_insert_text_box", "copy_formatting"
            DescribeAction = sp & PlanWords(t) & PlanArrowVal(act)
        Case "add_section", "delete_section", "rename_section", "move_section"
            DescribeAction = PlanWords(t) & " " & PlanQ(GetStr(act, "name"))
        Case "swap_font_deck_wide", "recolor_palette_deck_wide", "recolor_deck", "scan_palette", _
             "recolor_fill_match", "recolor_font_match", "delete_shapes_match"
            DescribeAction = PlanWords(t) & PlanArrowVal(act) & " (deck-wide)"

        ' ---- tables ----
        Case "add_table", "add_table_row", "delete_table_row", "add_table_col", "delete_table_col", _
             "merge_cells", "unmerge_cells", "swap_table_columns", "swap_table_rows", _
             "set_table_col_width", "set_table_row_height", "set_cell_border", "set_cell_text_align", _
             "set_cell_fill", "apply_table_style", "set_cell_padding", "set_table_style_options", _
             "populate_table_row", "populate_table_column", "populate_table_cells", _
             "set_cell_text_orientation", "set_row_fill", "set_column_fill", "clear_row_text", _
             "clear_column_text", "auto_fit_table_text", "set_table_borders", "set_row_borders", _
             "set_column_borders", "fit_cell_to_content", "add_cell_paragraph", "delete_cell_paragraph"
            DescribeAction = sp & PlanWords(t) & PlanOn(tg)

        ' ---- charts ----
        Case "add_chart", "set_chart_type", "set_chart_title", "set_chart_axis_title", _
             "set_chart_legend_position", "set_chart_legend", "set_series_color", "set_series_values", _
             "set_chart_categories", "set_series_name", "set_chart_axis", "set_chart_gridlines", _
             "set_chart_format", "set_chart_series", "add_chart_trendline", "set_chart_error_bars", _
             "set_chart_data_table", "set_line_smoothing", "delete_series", "add_series", _
             "set_data_label_text"
            DescribeAction = sp & PlanWords(t) & PlanOn(tg) & PlanArrowVal(act)

        ' ---- text frame / autofit ----
        Case "set_text_vertical_align", "set_text_autofit", "enable_text_shrink_for_overflow", _
             "set_text_margin"
            DescribeAction = sp & PlanWords(t) & PlanOn(tg) & PlanArrowVal(act)

        ' ---- verification ----
        Case "run_verification"
            DescribeAction = "run the slide-quality verification sweep"
        Case "apply_template"
            DescribeAction = sp & "apply the " & PlanQ(GetStr(act, "template")) & " slide template"

        Case Else
            If PlanIsKnownType(t) Then
                ' Known action with no bespoke template: still specific --
                ' the verb is derived from the type, plus slide/target/value.
                DescribeAction = sp & PlanWords(t) & PlanOn(tg) & PlanArrowVal(act)
            Else
                DescribeAction = "UNKNOWN ACTION " & t
            End If
    End Select
End Function

' True if t is a dispatched/known action type (reuses the canonical CSV).
Private Function PlanIsKnownType(t As String) As Boolean
    PlanIsKnownType = (InStr("," & GetAllActionTypes() & ",", "," & t & ",") > 0)
End Function

' "Slide N: " when a slide is given, else "".
Private Function PlanSlidePrefix(act As Object) As String
    If act.Exists("slide") Then
        PlanSlidePrefix = "Slide " & CStr(act("slide")) & ": "
    End If
End Function

' Quoted shape_name, else "shape #id", else "".
Private Function PlanTarget(act As Object) As String
    If act.Exists("shape_name") Then
        PlanTarget = PlanQ(CStr(act("shape_name")))
    ElseIf act.Exists("shape_id") Then
        PlanTarget = "shape #" & CStr(act("shape_id"))
    End If
End Function

Private Function PlanOn(tg As String) As String
    If Len(tg) > 0 Then PlanOn = " on " & tg
End Function

' Humanize an action type into a verb phrase: drop trailing noise and
' turn underscores into spaces. Distinct per type, so always specific.
Private Function PlanWords(t As String) As String
    PlanWords = Replace(t, "_", " ")
End Function

Private Function PlanQ(s As String) As String
    PlanQ = """" & s & """"
End Function

' Value of a key, or "?" if absent (keeps the sentence readable).
Private Function PlanV(act As Object, key As String) As String
    If act.Exists(key) Then
        PlanV = CStr(act(key))
    Else
        PlanV = "?"
    End If
End Function

' " -> true/false" for boolean-ish actions (value/bold/enabled/on keys).
Private Function PlanBool(act As Object) As String
    Dim k As Variant
    For Each k In Array("value", "enabled", "on", "bold", "italic", "underline")
        If act.Exists(k) Then
            PlanBool = LCase(CStr(act(k)))
            Exit Function
        End If
    Next k
    PlanBool = "(toggle)"
End Function

Private Function PlanColor(act As Object) As String
    If act.Exists("value") Then
        PlanColor = CStr(act("value"))
    ElseIf act.Exists("color") Then
        PlanColor = CStr(act("color"))
    Else
        PlanColor = "(color)"
    End If
End Function

' " -> <salient value>" using the first present well-known value key.
Private Function PlanArrowVal(act As Object) As String
    Dim k As Variant
    For Each k In Array("value", "text", "name", "color", "layout", _
                        "size", "count", "amount", "angle", "width", "height", _
                        "address", "style", "position", "scope")
        If act.Exists(k) Then
            PlanArrowVal = " -> " & CStr(act(k))
            Exit Function
        End If
    Next k
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
        Case "insert_icon"
            ValidateAction = RequireFields(act, Array("slide", "icon", "left", "top", "width", "height"))
        Case "insert_picture"
            ' Accept "path" or "picture_path" (LLMs and bulk_insert_image use the latter)
            If Not act.Exists("path") And Not act.Exists("picture_path") Then
                ValidateAction = "missing_field: path or picture_path"
            ElseIf Not act.Exists("pos") Then
                ValidateAction = "missing_field: pos"
            Else
                ValidateAction = ValidateSlide(act)
            End If
        Case "replace_picture"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "path"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "fetch_page_images"
            ValidateAction = RequireFields(act, Array("url"))
        Case "open_image_picker"
            ' folder optional - falls back to last fetch
            ValidateAction = ""
        Case "build_image_picker_slide"
            ValidateAction = ""   ' folder optional - falls back to g_LastFetchFolder
        Case "download_image"
            ValidateAction = RequireFields(act, Array("url", "dest_path"))
        Case "build_image_grid_table"
            ValidateAction = RequireFields(act, Array("slide", "pos", "rows"))
        Case "move_slide"
            ' Accept from/to OR from_slide/to_slide (less ambiguous vs add_connector)
            If Not act.Exists("from") And Not act.Exists("from_slide") Then
                ValidateAction = "missing_field: from or from_slide"
            ElseIf Not act.Exists("to") And Not act.Exists("to_slide") Then
                ValidateAction = "missing_field: to or to_slide"
            End If
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
            If Not act.Exists("from_shape_id") And Not act.Exists("from_shape_name") Then
                ValidateAction = "missing: from_shape_id or from_shape_name"
            ElseIf Not act.Exists("to_shape_id") And Not act.Exists("to_shape_name") Then
                ValidateAction = "missing: to_shape_id or to_shape_name"
            ElseIf Not act.Exists("kind") Then
                ValidateAction = "missing: kind"
            Else
                If Not act.Exists("slide") Then
                    ValidateAction = "missing: slide"
                Else
                    ValidateAction = ValidateSlide(act)
                End If
            End If
        Case "add_chart"
            ValidateAction = RequireFields(act, Array("slide", "chart_type", "pos", "categories", "series"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "set_chart_axis"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "axis", "props"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_chart_gridlines"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "props"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_chart_format"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "props"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "add_chart_trendline"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "series_index", "props"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_chart_error_bars"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "series_index", "props"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_chart_series"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "series_index", "props"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_chart_legend"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "props"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
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
        Case "set_run_bold", "set_run_italic", "set_run_underline", _
             "set_run_subscript", "set_run_superscript"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "run_index", "value"))
        Case "set_run_font_color"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "run_index", "value"))
        Case "set_run_font_size"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "run_index", "value"))
            If ValidateAction = "" Then
                If Not IsNumeric(act("value")) Or CDbl(act("value")) <= 0 Then
                    ValidateAction = "value: must be a positive number"
                End If
            End If
        Case "set_run_font_name"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "run_index", "value"))
            If ValidateAction = "" Then
                If Len(Trim(CStr(act("value")))) = 0 Then ValidateAction = "value: empty font name"
            End If
        Case "set_run_text"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "run_index", "value"))
        Case "add_run"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
        Case "set_run_hyperlink"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "run_index", "value"))
            If ValidateAction = "" Then
                Dim u As String: u = CStr(act("value"))
                If Len(u) > 0 _
                    And Not (LCase(Left(u, 7)) = "http://") _
                    And Not (LCase(Left(u, 8)) = "https://") _
                    And Not (LCase(Left(u, 7)) = "mailto:") _
                    And Not (Left(u, 7) = "#slide:") Then
                    ValidateAction = "value: invalid hyperlink URL"
                End If
            End If
        Case "set_paragraph_alignment"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If ValidateAction = "" Then
                Dim av As String: av = LCase(CStr(act("value")))
                If av <> "left" And av <> "center" And av <> "right" And av <> "justify" Then
                    ValidateAction = "value: must be one of left, center, right, justify"
                End If
            End If
        Case "set_paragraph_line_spacing"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If ValidateAction = "" Then
                If Not IsNumeric(act("value")) Or CDbl(act("value")) <= 0 Then
                    ValidateAction = "value: must be a positive number"
                End If
            End If
        Case "set_text_vertical_align"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If ValidateAction = "" Then
                Dim ax As String: ax = LCase(CStr(act("value")))
                If ax <> "top" And ax <> "middle" And ax <> "bottom" Then
                    ValidateAction = "value: must be one of top, middle, bottom"
                End If
            End If
        Case "set_text_autofit"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "mode"))
            If ValidateAction = "" Then
                Dim mf As String: mf = LCase(CStr(act("mode")))
                If mf <> "none" And mf <> "shrink" And mf <> "resize" Then _
                    ValidateAction = "mode: must be none/shrink/resize"
            End If
        Case "enable_text_shrink_for_overflow"
            ValidateAction = RequireFields(act, Array("scope"))
        Case "set_text_margin"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "left", "right", "top", "bottom"))
            If ValidateAction = "" Then
                Dim k As Variant
                For Each k In Array("left", "right", "top", "bottom")
                    If Not IsNumeric(act(k)) Or CDbl(act(k)) < 0 Then
                        ValidateAction = CStr(k) & ": must be a number >= 0"
                        Exit For
                    End If
                Next k
            End If
        Case "snap_to_grid"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "grid_pt"))
            If ValidateAction = "" Then
                If Not IsNumeric(act("grid_pt")) Or CDbl(act("grid_pt")) <= 0 Then _
                    ValidateAction = "grid_pt: must be > 0"
            End If
        Case "align_to_slide_center"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "axis"))
            If ValidateAction = "" Then
                Dim ac As String: ac = LCase(CStr(act("axis")))
                If ac <> "h" And ac <> "v" And ac <> "both" Then _
                    ValidateAction = "axis: must be h, v, or both"
            End If
        Case "nudge"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "direction", "amount_pt"))
            If ValidateAction = "" Then
                Dim dr As String: dr = LCase(CStr(act("direction")))
                If dr <> "l" And dr <> "r" And dr <> "u" And dr <> "d" Then _
                    ValidateAction = "direction: must be l, r, u, or d"
                If ValidateAction = "" Then
                    If Not IsNumeric(act("amount_pt")) Or CDbl(act("amount_pt")) < 0 Then _
                        ValidateAction = "amount_pt: must be a number >= 0"
                End If
            End If
        Case "fit_to_content"
            ValidateAction = RequireFields(act, Array("slide", "shape_id"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "match_size"
            ValidateAction = RequireFields(act, Array("slide", "ref_shape_id", "target_shape_ids"))
        Case "uniform_size"
            ValidateAction = RequireFields(act, Array("slide", "shape_ids", "width_pt", "height_pt"))
            If ValidateAction = "" Then
                If Not IsNumeric(act("width_pt")) Or CDbl(act("width_pt")) <= 0 _
                    Or Not IsNumeric(act("height_pt")) Or CDbl(act("height_pt")) <= 0 Then _
                    ValidateAction = "width_pt/height_pt: must be > 0"
            End If
        Case "smart_spacing"
            ValidateAction = RequireFields(act, Array("slide", "shape_ids", "gap_pt", "axis"))
            If ValidateAction = "" Then
                Dim sx As String: sx = LCase(CStr(act("axis")))
                If sx <> "h" And sx <> "v" Then ValidateAction = "axis: must be h or v"
                If ValidateAction = "" Then
                    If Not IsNumeric(act("gap_pt")) Or CDbl(act("gap_pt")) < 0 Then _
                        ValidateAction = "gap_pt: must be a number >= 0"
                End If
            End If
        Case "equalize_spacing"
            ValidateAction = RequireFields(act, Array("slide", "shape_ids", "axis"))
            If ValidateAction = "" Then
                Dim ex As String: ex = LCase(CStr(act("axis")))
                If ex <> "h" And ex <> "v" Then ValidateAction = "axis: must be h or v"
            End If
        Case "match_position"
            ValidateAction = RequireFields(act, Array("slide", "ref_shape_id", "target_shape_id", "edge"))
            If ValidateAction = "" Then
                Dim eg As String: eg = LCase(CStr(act("edge")))
                If eg <> "left" And eg <> "right" And eg <> "top" And eg <> "bottom" _
                    And eg <> "hcenter" And eg <> "vcenter" Then _
                    ValidateAction = "edge: must be left/right/top/bottom/hcenter/vcenter"
            End If
        Case "swap_positions"
            ValidateAction = RequireFields(act, Array("slide", "shape_a_id", "shape_b_id"))
        Case "group_by_overlap"
            ValidateAction = RequireFields(act, Array("slide", "shape_ids"))
        Case "find_replace_regex"
            ValidateAction = RequireFields(act, Array("scope", "pattern", "replacement"))
            If ValidateAction = "" Then
                If Len(CStr(act("pattern"))) = 0 Then ValidateAction = "pattern: empty"
            End If
        Case "swap_font_deck_wide"
            ValidateAction = RequireFields(act, Array("from_name", "to_name"))
            If ValidateAction = "" Then
                If Len(Trim(CStr(act("from_name")))) = 0 Or Len(Trim(CStr(act("to_name")))) = 0 Then _
                    ValidateAction = "from_name/to_name: empty"
            End If
        Case "recolor_palette_deck_wide"
            ValidateAction = RequireFields(act, Array("from_hex", "to_hex", "target"))
            If ValidateAction = "" Then
                Dim tg As String: tg = LCase(CStr(act("target")))
                If tg <> "fill" And tg <> "font" And tg <> "both" Then _
                    ValidateAction = "target: must be fill/font/both"
            End If
        Case "recolor_deck"
            ValidateAction = RequireFields(act, Array("mappings"))
            If ValidateAction = "" Then
                Dim rdMappings As Object: Set rdMappings = act("mappings")
                If rdMappings Is Nothing Or rdMappings.Count = 0 Then _
                    ValidateAction = "mappings: must be non-empty array of {from,to} objects"
            End If
        Case "scan_palette"
            ' no required fields; validate optional scope
            If act.Exists("scope") Then
                Dim spScope As String: spScope = LCase(CStr(act("scope")))
                If spScope <> "deck" And Left(spScope, 6) <> "slide:" Then _
                    ValidateAction = "scope: must be 'deck' or 'slide:N'"
            End If
        Case "apply_theme"
            ValidateAction = RequireFields(act, Array("theme_path"))
            If ValidateAction = "" Then
                If Len(CStr(act("theme_path"))) = 0 Then ValidateAction = "theme_path: empty"
            End If
        Case "set_slide_size"
            ' Either (width_pt + height_pt) OR (preset) — not both
            Dim hasDims As Boolean: hasDims = act.Exists("width_pt") And act.Exists("height_pt")
            Dim hasPreset As Boolean: hasPreset = act.Exists("preset")
            If hasDims And hasPreset Then
                ValidateAction = "specify dims OR preset, not both"
            ElseIf Not hasDims And Not hasPreset Then
                ValidateAction = "missing_field: width_pt+height_pt or preset"
            ElseIf hasDims Then
                If Not IsNumeric(act("width_pt")) Or CDbl(act("width_pt")) <= 0 _
                    Or Not IsNumeric(act("height_pt")) Or CDbl(act("height_pt")) <= 0 Then _
                    ValidateAction = "width_pt/height_pt: must be > 0"
            Else
                Dim ps As String: ps = LCase(CStr(act("preset")))
                If ps <> "16:9" And ps <> "4:3" Then ValidateAction = "preset: must be 16:9 or 4:3"
            End If
        Case "set_theme_font"
            ' At least one of major/minor must be present and non-empty
            Dim hasMajor As Boolean: hasMajor = act.Exists("major") And Len(CStr(act("major"))) > 0
            Dim hasMinor As Boolean: hasMinor = act.Exists("minor") And Len(CStr(act("minor"))) > 0
            If Not hasMajor And Not hasMinor Then ValidateAction = "set_theme_font: need major or minor"
        Case "bulk_insert_image"
            ValidateAction = RequireFields(act, Array("slide_indices", "picture_path", "left", "top", "width", "height"))
            If ValidateAction = "" Then
                If Len(CStr(act("picture_path"))) = 0 Then ValidateAction = "picture_path: empty"
            End If
        Case "bulk_insert_text_box"
            ValidateAction = RequireFields(act, Array("slide_indices", "text", "left", "top", "width", "height"))
        Case "apply_layout_to_slides"
            ValidateAction = RequireFields(act, Array("slide_indices", "layout_index"))
            If ValidateAction = "" Then
                If Not IsNumeric(act("layout_index")) Or CLng(act("layout_index")) < 0 Then _
                    ValidateAction = "layout_index: must be >= 0"
            End If
        Case "rotate_shape"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "degrees"))
        Case "set_shape_adjustment"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "index", "value"))
        Case "flip_shape"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "axis"))
            If ValidateAction = "" Then
                Dim fa As String: fa = LCase(CStr(act("axis")))
                If fa <> "h" And fa <> "v" Then ValidateAction = "axis: must be h or v"
            End If
        Case "set_line_color"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
        Case "set_line_weight"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "weight_pt"))
            If ValidateAction = "" Then
                If Not IsNumeric(act("weight_pt")) Or CDbl(act("weight_pt")) <= 0 Then _
                    ValidateAction = "weight_pt: must be > 0"
            End If
        Case "set_line_style"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "style"))
            If ValidateAction = "" Then
                Dim ls As String: ls = LCase(CStr(act("style")))
                If ls <> "solid" And ls <> "dash" And ls <> "dot" And ls <> "dashdot" Then _
                    ValidateAction = "style: must be solid/dash/dot/dashdot"
            End If
        Case "set_shadow"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "offset_x", "offset_y", "blur", "color", "transparency"))
        Case "set_glow"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "color", "radius", "transparency"))
        Case "set_reflection"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "size", "transparency", "distance"))
        Case "set_transparency"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If ValidateAction = "" Then
                If Not IsNumeric(act("value")) Or CDbl(act("value")) < 0 Or CDbl(act("value")) > 1 Then _
                    ValidateAction = "value: must be 0..1"
            End If
        Case "set_gradient_fill"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "color1", "color2", "angle"))
        Case "set_3d_bevel"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "bevel_type", "depth_pt"))
            If ValidateAction = "" Then
                Dim bv As String: bv = LCase(CStr(act("bevel_type")))
                If bv <> "circle" And bv <> "slope" And bv <> "cross" And bv <> "angle" And bv <> "softround" Then _
                    ValidateAction = "bevel_type: must be circle/slope/cross/angle/softround"
            End If
        Case "apply_preset_effect"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "preset_index"))
            If ValidateAction = "" Then
                Dim pi As Long: pi = CLng(act("preset_index"))
                If pi < 1 Or pi > 24 Then ValidateAction = "preset_index: must be 1..24"
            End If
        Case "crop_picture"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "left", "right", "top", "bottom"))
        Case "recolor_picture"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "color_type"))
            If ValidateAction = "" Then
                Dim ct As String: ct = LCase(CStr(act("color_type")))
                If ct <> "grayscale" And ct <> "sepia" And ct <> "washout" And ct <> "bw" And ct <> "auto" Then _
                    ValidateAction = "color_type: must be grayscale/sepia/washout/bw/auto"
            End If
        Case "set_brightness", "set_contrast"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If ValidateAction = "" Then
                If Not IsNumeric(act("value")) Or CDbl(act("value")) < -1 Or CDbl(act("value")) > 1 Then _
                    ValidateAction = "value: must be -1..1"
            End If
        Case "add_text_box"
            ValidateAction = RequireFields(act, Array("slide", "text", "pos"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "z_order"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "order"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
            If Len(ValidateAction) = 0 Then
                Dim zo As String: zo = LCase(CStr(act("order")))
                If zo <> "front" And zo <> "back" And zo <> "forward" And zo <> "backward" Then _
                    ValidateAction = "order: must be front/back/forward/backward"
            End If
        Case "duplicate_shape"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "left", "top"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "add_table"
            ValidateAction = RequireFields(act, Array("slide", "rows", "cols", "pos"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "set_table_col_width"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "col", "width_pt"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_table_row_height"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "height_pt"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_cell_border"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "side"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_cell_text_align"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_cell_fill"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "color"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "apply_table_style"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "style_id"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_series_values"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "series_index", "values"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_chart_categories"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "categories"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_series_name"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "series_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_slide_background_color"
            ValidateAction = RequireFields(act, Array("slide", "color"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "insert_slide_number"
            ValidateAction = RequireFields(act, Array("slide", "pos"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "copy_formatting"
            ValidateAction = RequireFields(act, Array("slide", "source_shape_id", "target_shape_id"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "set_run_strikethrough"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "run_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        ' --- New granular actions (paragraph / shape / effects / slide / table / chart) ---
        Case "set_paragraph_bold", "set_paragraph_italic", "set_paragraph_underline"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_paragraph_font_name"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then
                If Len(Trim(CStr(act("value")))) = 0 Then ValidateAction = "value: empty font name"
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_paragraph_space_before", "set_paragraph_space_after"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then
                If Not IsNumeric(act("value")) Then ValidateAction = "value: must be a number"
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "clear_paragraph_formatting"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_run_highlight"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "run_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_shape_name"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then
                If Len(Trim(CStr(act("value")))) = 0 Then ValidateAction = "value: empty shape name"
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_pos"
            ' At least one of left/top/width/height must be present
            ValidateAction = RequireFields(act, Array("slide", "shape_id"))
            If Len(ValidateAction) = 0 Then
                If Not (act.Exists("left") Or act.Exists("top") Or act.Exists("width") Or act.Exists("height")) Then
                    ValidateAction = "set_pos: at least one of left/top/width/height required"
                End If
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_shape_alt_text"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "lock_aspect_ratio"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "clear_shadow", "clear_glow", "clear_reflection", "clear_all_effects"
            ValidateAction = RequireFields(act, Array("slide", "shape_id"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_soft_edge"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "radius_pt"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_3d_rotation"
            ValidateAction = RequireFields(act, Array("slide", "shape_id"))
            If Len(ValidateAction) = 0 Then
                If Not (act.Exists("x") Or act.Exists("y") Or act.Exists("z")) Then
                    ValidateAction = "set_3d_rotation: at least one of x/y/z required"
                End If
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_slide_hidden"
            ValidateAction = RequireFields(act, Array("slide", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "clear_speaker_notes"
            ValidateAction = RequireFields(act, Array("slide"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "set_slide_name"
            ValidateAction = RequireFields(act, Array("slide", "value"))
            If Len(ValidateAction) = 0 Then
                If Len(Trim(CStr(act("value")))) = 0 Then ValidateAction = "value: empty slide name"
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "set_cell_padding"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "left", "right", "top", "bottom"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "clear_cell_text"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_table_style_options"
            ValidateAction = RequireFields(act, Array("slide", "shape_id"))
            If Len(ValidateAction) = 0 Then
                If Not (act.Exists("header_row") Or act.Exists("total_row") Or _
                        act.Exists("banded_rows") Or act.Exists("first_column") Or _
                        act.Exists("last_column") Or act.Exists("banded_columns")) Then
                    ValidateAction = "set_table_style_options: pass at least one of header_row/total_row/banded_rows/first_column/last_column/banded_columns"
                End If
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_chart_data_table"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "visible"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_line_smoothing"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "series_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "delete_series"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "series_index"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "add_series"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "name", "values"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        ' --- GRANULAR TABLE ACTIONS ---------------------------------------
        Case "populate_table_row"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "values"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "populate_table_column"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "col", "values"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "populate_table_cells"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "start_row", "start_col", "values"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_cell_font_size"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "value"))
            If Len(ValidateAction) = 0 Then
                If Not IsNumeric(act("value")) Or CLng(act("value")) <= 0 Then _
                    ValidateAction = "value: must be a positive integer"
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_cell_font_color", "set_cell_font_name"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_cell_font_bold", "set_cell_font_italic", "set_cell_font_underline"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_cell_text_orientation"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "value"))
            If Len(ValidateAction) = 0 Then
                Dim ctoVal As String: ctoVal = LCase(CStr(act("value")))
                If ctoVal <> "horizontal" And ctoVal <> "vertical_90" And _
                   ctoVal <> "vertical_270" And ctoVal <> "stacked" Then _
                    ValidateAction = "value: must be horizontal/vertical_90/vertical_270/stacked"
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_row_fill", "set_row_font_color"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_column_fill", "set_column_font_color"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "col", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_row_font_size", "set_row_font_bold"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_column_font_size", "set_column_font_bold"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "col", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "clear_row_text"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "clear_column_text"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "col"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_table_font_size"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then
                If Not IsNumeric(act("value")) Or CLng(act("value")) <= 0 Then _
                    ValidateAction = "value: must be a positive integer"
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_table_font_name", "set_table_font_color"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "auto_fit_table_text"
            ValidateAction = RequireFields(act, Array("slide", "shape_id"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_table_borders"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "side"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_row_borders"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "side"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_column_borders"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "col", "side"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "unmerge_cells"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        ' --- CELL PARAGRAPH ACTIONS ---------------------------------------
        Case "set_cell_paragraph_text"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_cell_paragraph_font_size"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then
                If Not IsNumeric(act("value")) Or CLng(act("value")) <= 0 Then _
                    ValidateAction = "value: must be a positive integer"
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_cell_paragraph_font_color"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_cell_paragraph_bold", "set_cell_paragraph_italic"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_cell_paragraph_alignment"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then
                Dim cpaVal As String: cpaVal = LCase(CStr(act("value")))
                If cpaVal <> "left" And cpaVal <> "center" And cpaVal <> "right" And cpaVal <> "justify" Then _
                    ValidateAction = "value: must be left/center/right/justify"
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_cell_bullet_style"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "add_cell_paragraph"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "after_paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "delete_cell_paragraph"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "paragraph_index"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_cell_indent_level"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "append_cell_text"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_cell"
            ' Mega-action: at least one of text/font_*/fill/h_align/v_align must be present
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col"))
            If Len(ValidateAction) = 0 Then
                If Not (act.Exists("text") Or act.Exists("font_size") Or act.Exists("font_color") _
                    Or act.Exists("font_bold") Or act.Exists("font_italic") Or act.Exists("font_underline") _
                    Or act.Exists("font_name") Or act.Exists("fill") _
                    Or act.Exists("h_align") Or act.Exists("v_align")) Then
                    ValidateAction = "set_cell: pass at least one of text/font_*/fill/h_align/v_align"
                End If
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        ' --- Shape fill / line / hyperlink / picture-fill ----------------
        Case "clear_fill", "clear_line"
            ValidateAction = RequireFields(act, Array("slide", "shape_id"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_fill_visible", "set_line_visible"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_shape_hyperlink"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then
                Dim shu As String: shu = CStr(act("value"))
                If Len(shu) > 0 _
                    And Not (LCase(Left(shu, 7)) = "http://") _
                    And Not (LCase(Left(shu, 8)) = "https://") _
                    And Not (LCase(Left(shu, 7)) = "mailto:") _
                    And Not (Left(shu, 7) = "#slide:") Then
                    ValidateAction = "value: invalid hyperlink URL"
                End If
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_shape_picture_fill"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "picture_path"))
            If Len(ValidateAction) = 0 Then
                If Len(Trim(CStr(act("picture_path")))) = 0 Then ValidateAction = "picture_path: empty"
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        ' --- WAVE 3 -----------------------------------------------------
        Case "set_slide_transition"
            ValidateAction = RequireFields(act, Array("slide", "effect"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "change_slide_layout"
            ValidateAction = RequireFields(act, Array("slide", "layout_index"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "add_section"
            ValidateAction = RequireFields(act, Array("before_slide", "name"))
        Case "delete_section"
            ValidateAction = RequireFields(act, Array("section_index"))
        Case "rename_section"
            ValidateAction = RequireFields(act, Array("section_index", "name"))
        Case "move_section"
            ValidateAction = RequireFields(act, Array("section_index", "to_position"))
        Case "apply_picture_artistic_effect"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "effect"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "reset_picture"
            ValidateAction = RequireFields(act, Array("slide", "shape_id"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_shape_visible"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "reconnect_connector"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "from_shape_id", "to_shape_id"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_run_kerning"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "run_index", "value"))
            If Len(ValidateAction) = 0 Then
                If Not IsNumeric(act("value")) Then ValidateAction = "value: must be a number"
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_run_baseline_offset"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "run_index", "value"))
            If Len(ValidateAction) = 0 Then
                If Not IsNumeric(act("value")) Or CDbl(act("value")) < -1 Or CDbl(act("value")) > 1 Then _
                    ValidateAction = "value: must be -1..1"
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_bullet_start_number"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then
                If Not IsNumeric(act("value")) Or CLng(act("value")) < 1 Then _
                    ValidateAction = "value: must be a positive integer"
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_notes_font_size"
            ValidateAction = RequireFields(act, Array("slide", "value"))
            If Len(ValidateAction) = 0 Then
                If Not IsNumeric(act("value")) Or CLng(act("value")) <= 0 Then _
                    ValidateAction = "value: must be a positive integer"
            End If
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "set_notes_font_color", "set_notes_font_name"
            ValidateAction = RequireFields(act, Array("slide", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "set_notes_font_bold", "set_notes_font_italic"
            ValidateAction = RequireFields(act, Array("slide", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "fit_cell_to_content"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row", "col"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_data_label_text"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "series_index", "point_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "run_verification"
            ' Optional scope (default "deck") and max_warnings (default 100); no required fields.
            ValidateAction = ""
        Case "apply_template"
            If Not act.Exists("template") Then
                ValidateAction = "missing_field: template"
            ElseIf InStr("," & modActionsTemplate.TemplateNames() & ",", _
                         "," & LCase(CStr(act("template"))) & ",") = 0 Then
                ValidateAction = "template: must be one of " & modActionsTemplate.TemplateNames()
            ElseIf Not act.Exists("content") Then
                ValidateAction = "missing_field: content"
            Else
                ValidateAction = modActionsTemplate.ValidateTemplateSlots( _
                    LCase(CStr(act("template"))), act("content"))
            End If
        Case Else
            ValidateAction = "unknown_type: " & t
    End Select
End Function

' Headless, read-only batch validator: sanitize+parse a batch and run
' ValidateAction on each action, returning JSON
' {"results":[{"index":N,"type":"...","reason":"..."}]} where reason ""
' means valid. Does NOT mutate the deck or call any Do_* handler.
' (ValidateShape may inject a numeric shape_id into the in-memory act
' dict when a shape_name is given -- that is in-memory only.)
Public Function ValidateBatchJson(jsonText As String) As String
    Dim res As Object: Set res = CreateObject("Scripting.Dictionary")
    Dim arr As Object: Set arr = New Collection
    res.Add "results", arr

    Dim cleaned As String: cleaned = SanitizeJsonInput(jsonText)
    Dim parsed As Object
    On Error Resume Next
    Set parsed = modJSON.ParseJson(cleaned)
    If Err.Number <> 0 Then
        arr.Add MakeValRow(0, "", "ERROR: invalid JSON: " & Err.Description)
        Err.Clear
        On Error GoTo 0
        ValidateBatchJson = modJSON.ConvertToJson(res)
        Exit Function
    End If
    On Error GoTo 0

    If parsed Is Nothing Then
        arr.Add MakeValRow(0, "", "ERROR: invalid JSON")
        ValidateBatchJson = modJSON.ConvertToJson(res)
        Exit Function
    End If
    If Not parsed.Exists("actions") Then
        arr.Add MakeValRow(0, "", "ERROR: missing top-level 'actions' array")
        ValidateBatchJson = modJSON.ConvertToJson(res)
        Exit Function
    End If

    Dim actions As Object: Set actions = parsed("actions")
    Dim i As Long
    For i = 1 To actions.Count
        Dim act As Object: Set act = actions(i)
        Dim t As String: t = GetStr(act, "type")
        Dim reason As String
        On Error Resume Next
        reason = ValidateAction(act)
        If Err.Number <> 0 Then
            reason = "ERROR: validator raised: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        arr.Add MakeValRow(i, t, reason)
    Next i

    ValidateBatchJson = modJSON.ConvertToJson(res)
End Function

Private Function MakeValRow(idx As Long, t As String, reason As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.Add "index", idx
    d.Add "type", t
    d.Add "reason", reason
    Set MakeValRow = d
End Function

Private Function RequireFields(act As Object, fields As Variant) As String
    Dim i As Long
    Dim hasShapeId As Boolean: hasShapeId = False
    For i = LBound(fields) To UBound(fields)
        Dim f As String: f = CStr(fields(i))
        Dim altName As String: altName = ""
        ' Any *_shape_ids field accepts *_shape_names alternative
        ' Any *_shape_id  field accepts *_shape_name  alternative
        If InStr(f, "shape_ids") > 0 Then
            altName = Replace(f, "shape_ids", "shape_names")
        ElseIf InStr(f, "shape_id") > 0 Then
            altName = Replace(f, "shape_id", "shape_name")
        End If
        If f = "shape_id" Then hasShapeId = True
        If Len(altName) > 0 Then
            If Not act.Exists(f) And Not act.Exists(altName) Then
                RequireFields = "missing_field: " & f & " or " & altName
                Exit Function
            End If
        ElseIf Not act.Exists(f) Then
            RequireFields = "missing_field: " & f
            Exit Function
        End If
    Next i
    ' If singular shape_id was required, run ValidateShape so shape_name
    ' aliases get resolved to a numeric shape_id injected into the act dict.
    ' Idempotent — explicit ValidateShape calls in case bodies still safe.
    If hasShapeId Then
        Dim shapeErr As String: shapeErr = ValidateShape(act)
        If Len(shapeErr) > 0 Then
            RequireFields = shapeErr
            Exit Function
        End If
    End If
    RequireFields = ""
End Function

' Resolve a shape ref field (any *_shape_id) to a numeric Id, accepting either
' the *_id (numeric or string-of-digits) or the *_name (ref_name) variant.
Private Function ResolveActShapeId(act As Object, idKey As String) As Long
    Dim raw As Variant
    Dim nameKey As String: nameKey = Replace(idKey, "shape_id", "shape_name")
    If act.Exists(idKey) Then
        raw = act(idKey)
    ElseIf act.Exists(nameKey) Then
        raw = act(nameKey)
    Else
        Err.Raise vbObjectError + 2060, "ResolveActShapeId", _
                  "missing field " & idKey & " or " & nameKey
    End If
    ResolveActShapeId = modActions.ResolveShapeRef(CLng(act("slide")), raw, idKey)
End Function

' Resolve an array shape-ref field (*_shape_ids) into act("...shape_ids") with
' numeric Ids in place, so existing dispatch code can pass it to NormalizeIdsArray
' unchanged. Accepts strings (ref_names) or numbers per element.
Private Function ResolveActShapeIdArray(act As Object, idsKey As String) As Variant
    Dim namesKey As String: namesKey = Replace(idsKey, "shape_ids", "shape_names")
    ' JSON arrays parse to Collection (Object). Variant return holding Object needs Set.
    Dim key As String
    If act.Exists(idsKey) Then
        key = idsKey
    ElseIf act.Exists(namesKey) Then
        key = namesKey
    Else
        Err.Raise vbObjectError + 2061, "ResolveActShapeIdArray", _
                  "missing field " & idsKey & " or " & namesKey
    End If
    If IsObject(act(key)) Then
        Set ResolveActShapeIdArray = act(key)
    Else
        ResolveActShapeIdArray = act(key)
    End If
End Function

Private Function ValidateSlide(act As Object) As String
    If Not act.Exists("slide") Then ValidateSlide = "missing_field: slide": Exit Function
    If Not IsNumeric(act("slide")) Then
        ValidateSlide = "slide must be numeric (1-based), got: " & CStr(act("slide"))
        Exit Function
    End If
    Dim n As Long: n = CLng(act("slide"))
    If n < 1 Or n > ActivePresentation.Slides.Count Then
        ValidateSlide = "slide_out_of_range: " & n & " (deck has " & ActivePresentation.Slides.Count & ")"
    End If
End Function

Private Function ValidateShape(act As Object) As String
    Dim slideErr As String: slideErr = ValidateSlide(act)
    If Len(slideErr) > 0 Then
        ValidateShape = slideErr
        Exit Function
    End If
    Dim slideN As Long: slideN = CLng(act("slide"))
    ' Three accepted forms: shape_id (numeric), shape_id (string ref_name), shape_name (string)
    ' All normalize to numeric shape_id so downstream dispatch code is uniform.
    If act.Exists("shape_id") Then
        If Not IsNumeric(act("shape_id")) Then
            Dim shByIdName As Shape
            Set shByIdName = modActions.FindShapeByName(slideN, CStr(act("shape_id")))
            If shByIdName Is Nothing Then
                ValidateShape = "shape_id '" & CStr(act("shape_id")) & "': not found as Id or ref_name"
                Exit Function
            End If
            act.Remove "shape_id"
            act.Add "shape_id", CLng(shByIdName.Id)
        End If
    ElseIf act.Exists("shape_name") Then
        Dim sh As Shape
        Set sh = modActions.FindShapeByName(slideN, CStr(act("shape_name")))
        If sh Is Nothing Then
            ValidateShape = "shape_name '" & CStr(act("shape_name")) & "': not found"
            Exit Function
        End If
        act.Add "shape_id", CLng(sh.Id)
    Else
        ValidateShape = "shape_id or shape_name: required"
        Exit Function
    End If
    Dim shCheck As Shape
    Set shCheck = modActions.FindShape(slideN, CLng(act("shape_id")))
    If shCheck Is Nothing Then ValidateShape = "shape_not_found: id=" & CLng(act("shape_id"))
End Function

Private Sub DispatchAction(act As Object)
    Dim t As String: t = act("type")
    Select Case t
        Case "set_text"
            modActions.Do_set_text CLng(act("slide")), CLng(act("shape_id")), CStr(act("value"))
        Case "set_font_size"
            modActions.Do_set_font_size CLng(act("slide")), CLng(act("shape_id")), CLng(act("value"))
        Case "set_font_bold"
            modActions.Do_set_font_bold CLng(act("slide")), CLng(act("shape_id")), modActions.ToBool(act("value"))
        Case "set_font_italic"
            modActions.Do_set_font_italic CLng(act("slide")), CLng(act("shape_id")), modActions.ToBool(act("value"))
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
            Dim frCase As Boolean: frCase = False
            Dim frWord As Boolean: frWord = False
            Dim frNotes As Boolean: frNotes = False
            If act.Exists("case_sensitive") Then frCase = modActions.ToBool(act("case_sensitive"))
            If act.Exists("whole_word") Then frWord = modActions.ToBool(act("whole_word"))
            If act.Exists("include_notes") Then frNotes = modActions.ToBool(act("include_notes"))
            modActionsText.Do_find_replace_text CStr(act("scope")), CStr(act("find")), CStr(act("replace")), _
                frCase, frWord, frNotes
        Case "align_shapes"
            modActionsLayout.Do_align_shapes CLng(act("slide")), ResolveActShapeIdArray(act, "shape_ids"), CStr(act("anchor"))
        Case "distribute_horizontal"
            modActionsLayout.Do_distribute_horizontal CLng(act("slide")), ResolveActShapeIdArray(act, "shape_ids")
        Case "distribute_vertical"
            modActionsLayout.Do_distribute_vertical CLng(act("slide")), ResolveActShapeIdArray(act, "shape_ids")
        Case "tile_grid"
            modActionsLayout.Do_tile_grid CLng(act("slide")), ResolveActShapeIdArray(act, "shape_ids"), _
                                          CLng(act("cols")), CSng(act("gap_pt"))
        Case "fit_to_slide_margins"
            Dim m As Single: m = 36.0
            If act.Exists("margin_pt") Then m = CSng(act("margin_pt"))
            modActionsLayout.Do_fit_to_slide_margins CLng(act("slide")), CLng(act("shape_id")), m
        Case "add_line"
            Dim alAE As String: alAE = "none"
            Dim alAS As String: alAS = "none"
            Dim alDS As String: alDS = "solid"
            If act.Exists("arrow_end") Then alAE = CStr(act("arrow_end"))
            If act.Exists("arrow_start") Then alAS = CStr(act("arrow_start"))
            If act.Exists("dash_style") Then alDS = CStr(act("dash_style"))
            modActionsLayout.Do_add_line CLng(act("slide")), CSng(act("x1")), CSng(act("y1")), _
                                         CSng(act("x2")), CSng(act("y2")), _
                                         CStr(act("color")), CSng(act("weight_pt")), _
                                         alAE, alAS, alDS
        Case "add_shape"
            Dim posDict As Object: Set posDict = act("pos")
            Dim fh As String: fh = ""
            Dim shex As String: shex = ""
            Dim swt As Single: swt = 1.0
            Dim asRef As String: asRef = ""
            Dim asTxt As String: asTxt = ""
            Dim asFc As String: asFc = ""
            Dim asFs As Long: asFs = 0
            Dim asBold As Boolean: asBold = False
            Dim asAlign As String: asAlign = "center"
            Dim asVAlign As String: asVAlign = "middle"
            If act.Exists("fill") Then If Not IsNull(act("fill")) Then fh = CStr(act("fill"))
            If act.Exists("stroke") Then If Not IsNull(act("stroke")) Then shex = CStr(act("stroke"))
            If act.Exists("stroke_weight_pt") Then swt = CSng(act("stroke_weight_pt"))
            If act.Exists("ref_name") Then asRef = CStr(act("ref_name"))
            If act.Exists("text") Then asTxt = CStr(act("text"))
            If act.Exists("font_color") Then asFc = CStr(act("font_color"))
            If act.Exists("font_size") Then asFs = CLng(act("font_size"))
            If act.Exists("font_bold") Then asBold = modActions.ToBool(act("font_bold"))
            If act.Exists("h_align") Then asAlign = CStr(act("h_align"))
            If act.Exists("v_align") Then asVAlign = CStr(act("v_align"))
            Dim asSuper As String: asSuper = ""
            Dim asSub As String: asSub = ""
            If act.Exists("super_suffix") Then asSuper = CStr(act("super_suffix"))
            If act.Exists("sub_suffix") Then asSub = CStr(act("sub_suffix"))
            modActionsLayout.Do_add_shape CLng(act("slide")), CStr(act("kind")), _
                                          CSng(posDict("left")), CSng(posDict("top")), _
                                          CSng(posDict("width")), CSng(posDict("height")), _
                                          fh, shex, swt, asRef, asTxt, asFc, asFs, asBold, asAlign, asVAlign, _
                                          asSuper, asSub
        Case "set_shape_kind"
            modActionsLayout.Do_set_shape_kind CLng(act("slide")), CLng(act("shape_id")), CStr(act("kind"))
        Case "clear_slide"
            Dim keep As Variant
            If act.Exists("keep_shape_ids") Then
                keep = ResolveActShapeIdArray(act, "keep_shape_ids")
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
        Case "insert_icon"
            modActionsIcon.Do_insert_icon act
        Case "insert_picture"
            Dim ipos As Object: Set ipos = act("pos")
            Dim ipPath As String
            If act.Exists("path") Then ipPath = CStr(act("path")) Else ipPath = CStr(act("picture_path"))
            modActionsImage.Do_insert_picture CLng(act("slide")), ipPath, _
                                              CSng(ipos("left")), CSng(ipos("top")), _
                                              CSng(ipos("width")), CSng(ipos("height"))
        Case "replace_picture"
            modActionsImage.Do_replace_picture CLng(act("slide")), CLng(act("shape_id")), CStr(act("path"))
        Case "fetch_page_images"
            Dim fpiFolder As String: fpiFolder = ""
            Dim fpiRef As String: fpiRef = ""
            If act.Exists("dest_folder") Then fpiFolder = CStr(act("dest_folder"))
            If act.Exists("ref_name") Then fpiRef = CStr(act("ref_name"))
            modActionsWeb.Do_fetch_page_images CStr(act("url")), fpiFolder, fpiRef
        Case "open_image_picker"
            Dim oipFolder As String: oipFolder = ""
            If act.Exists("folder") Then oipFolder = CStr(act("folder"))
            modActionsWeb.Do_open_image_picker oipFolder
        Case "build_image_picker_slide"
            Dim bipCols As Long: bipCols = 4
            Dim bipAt As Long: bipAt = 0
            Dim bipMax As Long: bipMax = 24
            If act.Exists("cols") Then bipCols = CLng(act("cols"))
            If act.Exists("insert_at") Then bipAt = CLng(act("insert_at"))
            If act.Exists("max_per_slide") Then bipMax = CLng(act("max_per_slide"))
            Dim bipFolder As String: bipFolder = ""
            If act.Exists("folder") Then bipFolder = CStr(act("folder"))
            modActionsImage.Do_build_image_picker_slide bipFolder, bipCols, bipAt, bipMax
        Case "download_image"
            modActionsWeb.Do_download_image CStr(act("url")), CStr(act("dest_path"))
        Case "build_image_grid_table"
            modActionsTable.Do_build_image_grid_table_act act
        Case "move_slide"
            Dim msFrom As Variant, msTo As Variant
            If act.Exists("from_slide") Then msFrom = act("from_slide") Else msFrom = act("from")
            If act.Exists("to_slide") Then msTo = act("to_slide") Else msTo = act("to")
            modActionsSlide.Do_move_slide CLng(msFrom), CLng(msTo)
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
            Dim grpRef As String: grpRef = ""
            If act.Exists("ref_name") Then grpRef = CStr(act("ref_name"))
            modActionsGroup.Do_group_shapes CLng(act("slide")), _
                                            ResolveActShapeIdArray(act, "shape_ids"), _
                                            grpRef
        Case "ungroup"
            modActionsGroup.Do_ungroup CLng(act("slide")), CLng(act("shape_id"))
        Case "add_connector"
            Dim ae As String, cc As String, cw As Single
            Dim astart As String, asize As String, fp As String, tp As String, ds As String
            ae = "filled": cc = "#000000": cw = 1.0
            astart = "none": asize = "medium": fp = "auto": tp = "auto": ds = "solid"
            If act.Exists("arrow_end") Then ae = CStr(act("arrow_end"))
            If act.Exists("color") Then cc = CStr(act("color"))
            If act.Exists("weight_pt") Then cw = CSng(act("weight_pt"))
            If act.Exists("arrow_start") Then astart = CStr(act("arrow_start"))
            If act.Exists("arrow_size") Then asize = CStr(act("arrow_size"))
            If act.Exists("from_point") Then fp = CStr(act("from_point"))
            If act.Exists("to_point") Then tp = CStr(act("to_point"))
            If act.Exists("dash_style") Then ds = CStr(act("dash_style"))
            Dim connSlide As Long: connSlide = CLng(act("slide"))
            Dim fromId As Long, toId As Long
            Dim shFr As Shape, shTo2 As Shape
            Dim fromRaw As String, toRaw As String
            If act.Exists("from_shape_id") Then
                fromRaw = CStr(act("from_shape_id"))
            Else
                fromRaw = CStr(act("from_shape_name"))
            End If
            If act.Exists("to_shape_id") Then
                toRaw = CStr(act("to_shape_id"))
            Else
                toRaw = CStr(act("to_shape_name"))
            End If
            If IsNumeric(fromRaw) Then
                fromId = CLng(fromRaw)
            Else
                Set shFr = modActions.FindShapeByName(connSlide, fromRaw)
                If shFr Is Nothing Then Err.Raise vbObjectError + 10002, "add_connector", "from shape not found: " & fromRaw
                fromId = shFr.Id
            End If
            If IsNumeric(toRaw) Then
                toId = CLng(toRaw)
            Else
                Set shTo2 = modActions.FindShapeByName(connSlide, toRaw)
                If shTo2 Is Nothing Then Err.Raise vbObjectError + 10002, "add_connector", "to shape not found: " & toRaw
                toId = shTo2.Id
            End If
            modActionsConnector.Do_add_connector connSlide, fromId, toId, _
                                                 CStr(act("kind")), ae, cc, cw, astart, asize, fp, tp, ds
        Case "add_chart"
            Dim acPos As Object: Set acPos = act("pos")
            Dim acRef As String: acRef = ""
            Dim acLeg As Boolean: acLeg = True
            Dim acVals As Boolean: acVals = False
            Dim acTitle As String: acTitle = ""
            Dim acClean As Boolean: acClean = False
            Dim acFmt As String: acFmt = ""
            If act.Exists("ref_name") Then acRef = CStr(act("ref_name"))
            If act.Exists("show_legend") Then acLeg = modActions.ToBool(act("show_legend"))
            If act.Exists("show_values") Then acVals = modActions.ToBool(act("show_values"))
            If act.Exists("title") Then acTitle = CStr(act("title"))
            If act.Exists("clean_style") Then acClean = modActions.ToBool(act("clean_style"))
            If act.Exists("value_format") Then acFmt = CStr(act("value_format"))
            modActionsChart.Do_add_chart CLng(act("slide")), CStr(act("chart_type")), _
                                          CSng(acPos("left")), CSng(acPos("top")), _
                                          CSng(acPos("width")), CSng(acPos("height")), _
                                          act("categories"), act("series"), _
                                          acRef, acLeg, acVals, acTitle, _
                                          acClean, acFmt
        Case "set_chart_axis"
            Dim caxProps As Object: Set caxProps = act("props")
            modActionsChart.Do_set_chart_axis CLng(act("slide")), CLng(act("shape_id")), _
                                               CStr(act("axis")), caxProps
        Case "set_chart_gridlines"
            Dim cglProps As Object: Set cglProps = act("props")
            Dim cglAxis As String: cglAxis = "y"
            If act.Exists("axis") Then cglAxis = CStr(act("axis"))
            modActionsChart.Do_set_chart_gridlines CLng(act("slide")), CLng(act("shape_id")), _
                                                   cglAxis, cglProps
        Case "set_chart_format"
            Dim cfmtProps As Object: Set cfmtProps = act("props")
            modActionsChart.Do_set_chart_format CLng(act("slide")), CLng(act("shape_id")), cfmtProps
        Case "add_chart_trendline"
            Dim ctlProps As Object: Set ctlProps = act("props")
            modActionsChart.Do_add_chart_trendline CLng(act("slide")), CLng(act("shape_id")), _
                                                    CLng(act("series_index")), ctlProps
        Case "set_chart_error_bars"
            Dim cebProps As Object: Set cebProps = act("props")
            modActionsChart.Do_set_chart_error_bars CLng(act("slide")), CLng(act("shape_id")), _
                                                     CLng(act("series_index")), cebProps
        Case "set_chart_series"
            Dim cssProps As Object: Set cssProps = act("props")
            modActionsChart.Do_set_chart_series CLng(act("slide")), CLng(act("shape_id")), _
                                                 CLng(act("series_index")), cssProps
        Case "set_chart_legend"
            Dim cleProps As Object: Set cleProps = act("props")
            modActionsChart.Do_set_chart_legend CLng(act("slide")), CLng(act("shape_id")), cleProps
        Case "set_chart_type"
            modActionsChart.Do_set_chart_type CLng(act("slide")), CLng(act("shape_id")), CStr(act("value"))
        Case "set_chart_title"
            Dim cte As Boolean: cte = True
            If act.Exists("enabled") Then cte = modActions.ToBool(act("enabled"))
            Dim cttProps As Object
            If act.Exists("props") Then Set cttProps = act("props")
            modActionsChart.Do_set_chart_title CLng(act("slide")), CLng(act("shape_id")), _
                                               CStr(act("value")), cte, cttProps
        Case "set_chart_axis_title"
            Dim catProps As Object
            If act.Exists("props") Then Set catProps = act("props")
            modActionsChart.Do_set_chart_axis_title CLng(act("slide")), CLng(act("shape_id")), _
                                                    CStr(act("axis")), CStr(act("value")), catProps
        Case "set_chart_legend_position"
            modActionsChart.Do_set_chart_legend_position CLng(act("slide")), CLng(act("shape_id")), _
                                                          CStr(act("value"))
        Case "set_series_color"
            modActionsChart.Do_set_series_color CLng(act("slide")), CLng(act("shape_id")), _
                                                CLng(act("series_index")), CStr(act("value"))
        Case "set_run_bold"
            modActionsRun.Do_set_run_bold CLng(act("slide")), CLng(act("shape_id")), _
                                          CLng(act("paragraph_index")), CLng(act("run_index")), _
                                          modActions.ToBool(act("value"))
        Case "set_run_italic"
            modActionsRun.Do_set_run_italic CLng(act("slide")), CLng(act("shape_id")), _
                                            CLng(act("paragraph_index")), CLng(act("run_index")), _
                                            modActions.ToBool(act("value"))
        Case "set_run_underline"
            modActionsRun.Do_set_run_underline CLng(act("slide")), CLng(act("shape_id")), _
                                               CLng(act("paragraph_index")), CLng(act("run_index")), _
                                               modActions.ToBool(act("value"))
        Case "set_run_subscript"
            modActionsRun.Do_set_run_subscript CLng(act("slide")), CLng(act("shape_id")), _
                                               CLng(act("paragraph_index")), CLng(act("run_index")), _
                                               modActions.ToBool(act("value"))
        Case "set_run_superscript"
            modActionsRun.Do_set_run_superscript CLng(act("slide")), CLng(act("shape_id")), _
                                                 CLng(act("paragraph_index")), CLng(act("run_index")), _
                                                 modActions.ToBool(act("value"))
        Case "set_run_font_color"
            modActionsRun.Do_set_run_font_color CLng(act("slide")), CLng(act("shape_id")), _
                                                CLng(act("paragraph_index")), CLng(act("run_index")), _
                                                CStr(act("value"))
        Case "set_run_font_size"
            modActionsRun.Do_set_run_font_size CLng(act("slide")), CLng(act("shape_id")), _
                                               CLng(act("paragraph_index")), CLng(act("run_index")), _
                                               CLng(act("value"))
        Case "set_run_font_name"
            modActionsRun.Do_set_run_font_name CLng(act("slide")), CLng(act("shape_id")), _
                                               CLng(act("paragraph_index")), CLng(act("run_index")), _
                                               CStr(act("value"))
        Case "set_run_text"
            modActionsRun.Do_set_run_text CLng(act("slide")), CLng(act("shape_id")), _
                                          CLng(act("paragraph_index")), CLng(act("run_index")), _
                                          CStr(act("value"))
        Case "add_run"
            Dim arBold As Long: arBold = -1
            Dim arItalic As Long: arItalic = -1
            Dim arUnder As Long: arUnder = -1
            If act.Exists("bold") Then arBold = IIf(CBool(act("bold")), 1, 0)
            If act.Exists("italic") Then arItalic = IIf(CBool(act("italic")), 1, 0)
            If act.Exists("underline") Then arUnder = IIf(CBool(act("underline")), 1, 0)
            Dim arColor As String: arColor = ""
            Dim arFont As String: arFont = ""
            Dim arSize As Long: arSize = 0
            If act.Exists("color") Then arColor = CStr(act("color"))
            If act.Exists("font_name") Then arFont = CStr(act("font_name"))
            If act.Exists("font_size") Then arSize = CLng(act("font_size"))
            modActionsRun.Do_add_run CLng(act("slide")), CLng(act("shape_id")), _
                                     CLng(act("paragraph_index")), CStr(act("value")), _
                                     arBold, arItalic, arColor, arFont, arSize, arUnder
        Case "set_run_hyperlink"
            modActionsRun.Do_set_run_hyperlink CLng(act("slide")), CLng(act("shape_id")), _
                                               CLng(act("paragraph_index")), CLng(act("run_index")), _
                                               CStr(act("value"))
        Case "set_paragraph_alignment"
            modActionsText.Do_set_paragraph_alignment CLng(act("slide")), CLng(act("shape_id")), _
                                                      CLng(act("paragraph_index")), CStr(act("value"))
        Case "set_paragraph_line_spacing"
            modActionsText.Do_set_paragraph_line_spacing CLng(act("slide")), CLng(act("shape_id")), _
                                                         CLng(act("paragraph_index")), CDbl(act("value"))
        Case "set_text_vertical_align"
            modActionsText.Do_set_text_vertical_align CLng(act("slide")), CLng(act("shape_id")), _
                                                      CStr(act("value"))
        Case "set_text_autofit"
            modActionsText.Do_set_text_autofit CLng(act("slide")), CLng(act("shape_id")), CStr(act("mode"))
        Case "enable_text_shrink_for_overflow"
            Dim incTitlesArg As String: incTitlesArg = "false"
            If act.Exists("include_titles") Then incTitlesArg = CStr(act("include_titles"))
            modActionsText.Do_enable_text_shrink_for_overflow CStr(act("scope")), incTitlesArg
        Case "set_text_margin"
            modActionsText.Do_set_text_margin CLng(act("slide")), CLng(act("shape_id")), _
                                              CDbl(act("left")), CDbl(act("right")), _
                                              CDbl(act("top")), CDbl(act("bottom"))
        Case "snap_to_grid"
            modActionsLayout.Do_snap_to_grid CLng(act("slide")), CLng(act("shape_id")), _
                                             CDbl(act("grid_pt"))
        Case "align_to_slide_center"
            modActionsLayout.Do_align_to_slide_center CLng(act("slide")), CLng(act("shape_id")), _
                                                      CStr(act("axis"))
        Case "nudge"
            modActionsLayout.Do_nudge CLng(act("slide")), CLng(act("shape_id")), _
                                      CStr(act("direction")), CDbl(act("amount_pt"))
        Case "fit_to_content"
            modActionsLayout.Do_fit_to_content CLng(act("slide")), CLng(act("shape_id"))
        Case "match_size"
            modActionsLayout.Do_match_size CLng(act("slide")), _
                                           ResolveActShapeId(act, "ref_shape_id"), _
                                           ResolveActShapeIdArray(act, "target_shape_ids")
        Case "uniform_size"
            modActionsLayout.Do_uniform_size CLng(act("slide")), ResolveActShapeIdArray(act, "shape_ids"), _
                                             CDbl(act("width_pt")), CDbl(act("height_pt"))
        Case "smart_spacing"
            modActionsLayout.Do_smart_spacing CLng(act("slide")), ResolveActShapeIdArray(act, "shape_ids"), _
                                              CDbl(act("gap_pt")), CStr(act("axis"))
        Case "equalize_spacing"
            modActionsLayout.Do_equalize_spacing CLng(act("slide")), ResolveActShapeIdArray(act, "shape_ids"), _
                                                 CStr(act("axis"))
        Case "match_position"
            modActionsLayout.Do_match_position CLng(act("slide")), _
                                               ResolveActShapeId(act, "ref_shape_id"), _
                                               ResolveActShapeId(act, "target_shape_id"), _
                                               CStr(act("edge"))
        Case "swap_positions"
            modActionsLayout.Do_swap_positions CLng(act("slide")), _
                                               ResolveActShapeId(act, "shape_a_id"), _
                                               ResolveActShapeId(act, "shape_b_id")
        Case "group_by_overlap"
            modActionsLayout.Do_group_by_overlap CLng(act("slide")), ResolveActShapeIdArray(act, "shape_ids")
        Case "find_replace_regex"
            modActionsDeck.Do_find_replace_regex CStr(act("scope")), CStr(act("pattern")), CStr(act("replacement"))
        Case "swap_font_deck_wide"
            modActionsDeck.Do_swap_font_deck_wide CStr(act("from_name")), CStr(act("to_name"))
        Case "recolor_palette_deck_wide"
            modActionsDeck.Do_recolor_palette_deck_wide CStr(act("from_hex")), CStr(act("to_hex")), CStr(act("target"))
        Case "recolor_deck"
            Dim rdMap As Object: Set rdMap = act("mappings")
            Dim rdScope As String: rdScope = ""
            If act.Exists("scope") Then rdScope = CStr(act("scope"))
            modActionsDeck.Do_recolor_deck rdMap, rdScope
        Case "scan_palette"
            Dim spSc As String: spSc = "deck"
            If act.Exists("scope") Then spSc = CStr(act("scope"))
            modActionsDeck.Do_scan_palette spSc
        Case "apply_theme"
            modActionsDeck.Do_apply_theme CStr(act("theme_path"))
        Case "set_slide_size"
            If act.Exists("preset") Then
                modActionsDeck.Do_set_slide_size_preset CStr(act("preset"))
            Else
                modActionsDeck.Do_set_slide_size_dims CDbl(act("width_pt")), CDbl(act("height_pt"))
            End If
        Case "set_theme_font"
            Dim mj As String: mj = ""
            Dim mn As String: mn = ""
            If act.Exists("major") Then mj = CStr(act("major"))
            If act.Exists("minor") Then mn = CStr(act("minor"))
            modActionsDeck.Do_set_theme_font mj, mn
        Case "bulk_insert_image"
            modActionsDeck.Do_bulk_insert_image act("slide_indices"), CStr(act("picture_path")), _
                                                CDbl(act("left")), CDbl(act("top")), _
                                                CDbl(act("width")), CDbl(act("height"))
        Case "bulk_insert_text_box"
            modActionsDeck.Do_bulk_insert_text_box act("slide_indices"), CStr(act("text")), _
                                                   CDbl(act("left")), CDbl(act("top")), _
                                                   CDbl(act("width")), CDbl(act("height"))
        Case "apply_layout_to_slides"
            modActionsDeck.Do_apply_layout_to_slides act("slide_indices"), CLng(act("layout_index"))
        Case "rotate_shape"
            modActionsEffects.Do_rotate_shape CLng(act("slide")), CLng(act("shape_id")), CDbl(act("degrees"))
        Case "set_shape_adjustment"
            modActionsEffects.Do_set_shape_adjustment CLng(act("slide")), CLng(act("shape_id")), _
                                                       CLng(act("index")), CDbl(act("value"))
        Case "flip_shape"
            modActionsEffects.Do_flip_shape CLng(act("slide")), CLng(act("shape_id")), CStr(act("axis"))
        Case "set_line_color"
            modActionsEffects.Do_set_line_color CLng(act("slide")), CLng(act("shape_id")), CStr(act("value"))
        Case "set_line_weight"
            modActionsEffects.Do_set_line_weight CLng(act("slide")), CLng(act("shape_id")), CDbl(act("weight_pt"))
        Case "set_line_style"
            modActionsEffects.Do_set_line_style CLng(act("slide")), CLng(act("shape_id")), CStr(act("style"))
        Case "set_shadow"
            modActionsEffects.Do_set_shadow CLng(act("slide")), CLng(act("shape_id")), _
                CDbl(act("offset_x")), CDbl(act("offset_y")), CDbl(act("blur")), _
                CStr(act("color")), CDbl(act("transparency"))
        Case "set_glow"
            modActionsEffects.Do_set_glow CLng(act("slide")), CLng(act("shape_id")), _
                CStr(act("color")), CDbl(act("radius")), CDbl(act("transparency"))
        Case "set_reflection"
            modActionsEffects.Do_set_reflection CLng(act("slide")), CLng(act("shape_id")), _
                CDbl(act("size")), CDbl(act("transparency")), CDbl(act("distance"))
        Case "set_transparency"
            modActionsEffects.Do_set_transparency CLng(act("slide")), CLng(act("shape_id")), CDbl(act("value"))
        Case "set_gradient_fill"
            modActionsEffects.Do_set_gradient_fill CLng(act("slide")), CLng(act("shape_id")), _
                CStr(act("color1")), CStr(act("color2")), CDbl(act("angle"))
        Case "set_3d_bevel"
            modActionsEffects.Do_set_3d_bevel CLng(act("slide")), CLng(act("shape_id")), _
                CStr(act("bevel_type")), CDbl(act("depth_pt"))
        Case "apply_preset_effect"
            modActionsEffects.Do_apply_preset_effect CLng(act("slide")), CLng(act("shape_id")), CLng(act("preset_index"))
        Case "crop_picture"
            modActionsEffects.Do_crop_picture CLng(act("slide")), CLng(act("shape_id")), _
                CDbl(act("left")), CDbl(act("right")), CDbl(act("top")), CDbl(act("bottom"))
        Case "recolor_picture"
            modActionsEffects.Do_recolor_picture CLng(act("slide")), CLng(act("shape_id")), CStr(act("color_type"))
        Case "set_brightness"
            modActionsEffects.Do_set_brightness CLng(act("slide")), CLng(act("shape_id")), CDbl(act("value"))
        Case "set_contrast"
            modActionsEffects.Do_set_contrast CLng(act("slide")), CLng(act("shape_id")), CDbl(act("value"))
        Case "add_text_box"
            Dim tbPos As Object: Set tbPos = act("pos")
            Dim tbRef As String: tbRef = ""
            Dim tbFc As String: tbFc = ""
            Dim tbFs As Long: tbFs = 0
            Dim tbBold As Boolean: tbBold = False
            Dim tbItalic As Boolean: tbItalic = False
            Dim tbAlign As String: tbAlign = ""
            Dim tbFill As String: tbFill = ""
            Dim tbStroke As String: tbStroke = ""
            Dim tbSw As Single: tbSw = 1.0
            If act.Exists("ref_name") Then tbRef = CStr(act("ref_name"))
            If act.Exists("font_color") Then tbFc = CStr(act("font_color"))
            If act.Exists("font_size") Then tbFs = CLng(act("font_size"))
            If act.Exists("font_bold") Then tbBold = modActions.ToBool(act("font_bold"))
            If act.Exists("font_italic") Then tbItalic = modActions.ToBool(act("font_italic"))
            If act.Exists("h_align") Then tbAlign = CStr(act("h_align"))
            If act.Exists("fill") Then If Not IsNull(act("fill")) Then tbFill = CStr(act("fill"))
            If act.Exists("stroke") Then If Not IsNull(act("stroke")) Then tbStroke = CStr(act("stroke"))
            If act.Exists("stroke_weight_pt") Then tbSw = CSng(act("stroke_weight_pt"))
            Dim tbSuper As String: tbSuper = ""
            Dim tbSub As String: tbSub = ""
            If act.Exists("super_suffix") Then tbSuper = CStr(act("super_suffix"))
            If act.Exists("sub_suffix") Then tbSub = CStr(act("sub_suffix"))
            modActionsLayout.Do_add_text_box CLng(act("slide")), CStr(act("text")), _
                CSng(tbPos("left")), CSng(tbPos("top")), CSng(tbPos("width")), CSng(tbPos("height")), _
                tbRef, tbFc, tbFs, tbBold, tbItalic, tbAlign, tbFill, tbStroke, tbSw, _
                tbSuper, tbSub
        Case "z_order"
            modActionsLayout.Do_z_order CLng(act("slide")), CLng(act("shape_id")), CStr(act("order"))
        Case "duplicate_shape"
            Dim dsRef As String: dsRef = ""
            If act.Exists("ref_name") Then dsRef = CStr(act("ref_name"))
            modActionsLayout.Do_duplicate_shape CLng(act("slide")), CLng(act("shape_id")), _
                CSng(act("left")), CSng(act("top")), dsRef
        Case "add_table"
            Dim atPos As Object: Set atPos = act("pos")
            Dim atRef As String: atRef = ""
            If act.Exists("ref_name") Then atRef = CStr(act("ref_name"))
            modActionsTable.Do_add_table CLng(act("slide")), CLng(act("rows")), CLng(act("cols")), _
                CSng(atPos("left")), CSng(atPos("top")), CSng(atPos("width")), CSng(atPos("height")), atRef
        Case "set_table_col_width"
            modActionsTable.Do_set_table_col_width CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("col")), CSng(act("width_pt"))
        Case "set_table_row_height"
            modActionsTable.Do_set_table_row_height CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CSng(act("height_pt"))
        Case "set_cell_border"
            Dim cbColor As String: cbColor = ""
            Dim cbWeight As Single: cbWeight = 0
            Dim cbVisible As Boolean: cbVisible = True
            Dim cbDash As String: cbDash = ""
            If act.Exists("color") Then cbColor = CStr(act("color"))
            If act.Exists("weight_pt") Then cbWeight = CSng(act("weight_pt"))
            If act.Exists("visible") Then cbVisible = modActions.ToBool(act("visible"))
            If act.Exists("dash_style") Then cbDash = CStr(act("dash_style"))
            modActionsTable.Do_set_cell_border CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CStr(act("side")), cbColor, cbWeight, cbVisible, cbDash
        Case "set_cell_text_align"
            Dim cthAlign As String: cthAlign = ""
            Dim ctvAlign As String: ctvAlign = ""
            If act.Exists("h_align") Then cthAlign = CStr(act("h_align"))
            If act.Exists("v_align") Then ctvAlign = CStr(act("v_align"))
            modActionsTable.Do_set_cell_text_align CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), cthAlign, ctvAlign
        Case "set_cell_fill"
            modActionsTable.Do_set_cell_fill CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CStr(act("color"))
        Case "apply_table_style"
            modActionsTable.Do_apply_table_style CLng(act("slide")), CLng(act("shape_id")), _
                CStr(act("style_id"))
        Case "set_series_values"
            modActionsChart.Do_set_series_values CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("series_index")), act("values")
        Case "set_chart_categories"
            modActionsChart.Do_set_chart_categories CLng(act("slide")), CLng(act("shape_id")), _
                act("categories")
        Case "set_series_name"
            modActionsChart.Do_set_series_name CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("series_index")), CStr(act("value"))
        Case "set_slide_background_color"
            modActionsSlide.Do_set_slide_background_color CLng(act("slide")), CStr(act("color"))
        Case "insert_slide_number"
            Dim isnPos As Object: Set isnPos = act("pos")
            Dim isnRef As String: isnRef = ""
            Dim isnFc As String: isnFc = ""
            Dim isnFs As Long: isnFs = 0
            If act.Exists("ref_name") Then isnRef = CStr(act("ref_name"))
            If act.Exists("font_color") Then isnFc = CStr(act("font_color"))
            If act.Exists("font_size") Then isnFs = CLng(act("font_size"))
            modActionsSlide.Do_insert_slide_number CLng(act("slide")), _
                CSng(isnPos("left")), CSng(isnPos("top")), _
                CSng(isnPos("width")), CSng(isnPos("height")), isnRef, isnFc, isnFs
        Case "copy_formatting"
            modActionsLayout.Do_copy_formatting CLng(act("slide")), _
                ResolveActShapeId(act, "source_shape_id"), _
                ResolveActShapeId(act, "target_shape_id")
        Case "set_run_strikethrough"
            modActionsRun.Do_set_run_strikethrough CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("paragraph_index")), CLng(act("run_index")), modActions.ToBool(act("value"))
        ' --- Granular paragraph-level toggles -----------------------------
        Case "set_paragraph_bold"
            modActionsText.Do_set_paragraph_bold CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("paragraph_index")), modActions.ToBool(act("value"))
        Case "set_paragraph_italic"
            modActionsText.Do_set_paragraph_italic CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("paragraph_index")), modActions.ToBool(act("value"))
        Case "set_paragraph_underline"
            modActionsText.Do_set_paragraph_underline CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("paragraph_index")), modActions.ToBool(act("value"))
        Case "set_paragraph_font_name"
            modActionsText.Do_set_paragraph_font_name CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("paragraph_index")), CStr(act("value"))
        Case "set_paragraph_space_before"
            modActionsText.Do_set_paragraph_space_before CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("paragraph_index")), CDbl(act("value"))
        Case "set_paragraph_space_after"
            modActionsText.Do_set_paragraph_space_after CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("paragraph_index")), CDbl(act("value"))
        Case "clear_paragraph_formatting"
            modActionsText.Do_clear_paragraph_formatting CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("paragraph_index"))
        ' --- Run-level highlight ------------------------------------------
        Case "set_run_highlight"
            modActionsRun.Do_set_run_highlight CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("paragraph_index")), CLng(act("run_index")), CStr(act("value"))
        ' --- Shape-level granular -----------------------------------------
        Case "set_shape_name"
            modActions.Do_set_shape_name CLng(act("slide")), CLng(act("shape_id")), CStr(act("value"))
        Case "set_pos"
            Dim spLeft As Single: spLeft = 0
            Dim spTop As Single: spTop = 0
            Dim spWidth As Single: spWidth = 0
            Dim spHeight As Single: spHeight = 0
            Dim spHasLeft As Boolean: spHasLeft = act.Exists("left")
            Dim spHasTop As Boolean: spHasTop = act.Exists("top")
            Dim spHasWidth As Boolean: spHasWidth = act.Exists("width")
            Dim spHasHeight As Boolean: spHasHeight = act.Exists("height")
            If spHasLeft Then spLeft = CSng(act("left"))
            If spHasTop Then spTop = CSng(act("top"))
            If spHasWidth Then spWidth = CSng(act("width"))
            If spHasHeight Then spHeight = CSng(act("height"))
            modActions.Do_set_pos CLng(act("slide")), CLng(act("shape_id")), _
                spLeft, spTop, spWidth, spHeight, _
                spHasLeft, spHasTop, spHasWidth, spHasHeight
        Case "set_shape_alt_text"
            modActions.Do_set_shape_alt_text CLng(act("slide")), CLng(act("shape_id")), CStr(act("value"))
        Case "lock_aspect_ratio"
            modActions.Do_lock_aspect_ratio CLng(act("slide")), CLng(act("shape_id")), _
                modActions.ToBool(act("value"))
        ' --- Effects clearers ---------------------------------------------
        Case "clear_shadow"
            modActionsEffects.Do_clear_shadow CLng(act("slide")), CLng(act("shape_id"))
        Case "clear_glow"
            modActionsEffects.Do_clear_glow CLng(act("slide")), CLng(act("shape_id"))
        Case "clear_reflection"
            modActionsEffects.Do_clear_reflection CLng(act("slide")), CLng(act("shape_id"))
        Case "clear_all_effects"
            modActionsEffects.Do_clear_all_effects CLng(act("slide")), CLng(act("shape_id"))
        Case "set_soft_edge"
            modActionsEffects.Do_set_soft_edge CLng(act("slide")), CLng(act("shape_id")), _
                CDbl(act("radius_pt"))
        Case "set_3d_rotation"
            Dim rotX As Double: rotX = 0
            Dim rotY As Double: rotY = 0
            Dim rotZ As Double: rotZ = 0
            Dim hasRX As Boolean: hasRX = act.Exists("x")
            Dim hasRY As Boolean: hasRY = act.Exists("y")
            Dim hasRZ As Boolean: hasRZ = act.Exists("z")
            If hasRX Then rotX = CDbl(act("x"))
            If hasRY Then rotY = CDbl(act("y"))
            If hasRZ Then rotZ = CDbl(act("z"))
            modActionsEffects.Do_set_3d_rotation CLng(act("slide")), CLng(act("shape_id")), _
                rotX, rotY, rotZ, hasRX, hasRY, hasRZ
        ' --- Slide-level granular -----------------------------------------
        Case "set_slide_hidden"
            modActionsSlide.Do_set_slide_hidden CLng(act("slide")), modActions.ToBool(act("value"))
        Case "clear_speaker_notes"
            modActionsSlide.Do_clear_speaker_notes CLng(act("slide"))
        Case "set_slide_name"
            modActionsSlide.Do_set_slide_name CLng(act("slide")), CStr(act("value"))
        ' --- Table-level granular -----------------------------------------
        Case "set_cell_padding"
            modActionsTable.Do_set_cell_padding CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), _
                CDbl(act("left")), CDbl(act("right")), _
                CDbl(act("top")), CDbl(act("bottom"))
        Case "clear_cell_text"
            modActionsTable.Do_clear_cell_text CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col"))
        Case "set_table_style_options"
            Dim tsHasHeader As Boolean: tsHasHeader = act.Exists("header_row")
            Dim tsHasTotal As Boolean: tsHasTotal = act.Exists("total_row")
            Dim tsHasBandR As Boolean: tsHasBandR = act.Exists("banded_rows")
            Dim tsHasFirstC As Boolean: tsHasFirstC = act.Exists("first_column")
            Dim tsHasLastC As Boolean: tsHasLastC = act.Exists("last_column")
            Dim tsHasBandC As Boolean: tsHasBandC = act.Exists("banded_columns")
            Dim tsHeader As Boolean: tsHeader = False
            Dim tsTotal As Boolean: tsTotal = False
            Dim tsBandR As Boolean: tsBandR = False
            Dim tsFirstC As Boolean: tsFirstC = False
            Dim tsLastC As Boolean: tsLastC = False
            Dim tsBandC As Boolean: tsBandC = False
            If tsHasHeader Then tsHeader = modActions.ToBool(act("header_row"))
            If tsHasTotal Then tsTotal = modActions.ToBool(act("total_row"))
            If tsHasBandR Then tsBandR = modActions.ToBool(act("banded_rows"))
            If tsHasFirstC Then tsFirstC = modActions.ToBool(act("first_column"))
            If tsHasLastC Then tsLastC = modActions.ToBool(act("last_column"))
            If tsHasBandC Then tsBandC = modActions.ToBool(act("banded_columns"))
            modActionsTable.Do_set_table_style_options CLng(act("slide")), CLng(act("shape_id")), _
                tsHasHeader, tsHeader, tsHasTotal, tsTotal, _
                tsHasBandR, tsBandR, tsHasFirstC, tsFirstC, _
                tsHasLastC, tsLastC, tsHasBandC, tsBandC
        ' --- Chart-level granular -----------------------------------------
        Case "set_chart_data_table"
            Dim cdtProps As Object
            If act.Exists("props") Then Set cdtProps = act("props")
            modActionsChart.Do_set_chart_data_table CLng(act("slide")), CLng(act("shape_id")), _
                modActions.ToBool(act("visible")), cdtProps
        Case "set_line_smoothing"
            modActionsChart.Do_set_line_smoothing CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("series_index")), modActions.ToBool(act("value"))
        Case "delete_series"
            modActionsChart.Do_delete_series CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("series_index"))
        Case "add_series"
            Dim asColor As String: asColor = ""
            If act.Exists("color") Then asColor = CStr(act("color"))
            modActionsChart.Do_add_series CLng(act("slide")), CLng(act("shape_id")), _
                CStr(act("name")), act("values"), asColor
        ' --- GRANULAR TABLE ACTIONS -------------------------------------------
        Case "populate_table_row"
            modActionsTable.Do_populate_table_row CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), act("values")
        Case "populate_table_column"
            modActionsTable.Do_populate_table_column CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("col")), act("values")
        Case "populate_table_cells"
            modActionsTable.Do_populate_table_cells CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("start_row")), CLng(act("start_col")), act("values")
        Case "set_cell_font_size"
            modActionsTable.Do_set_cell_font_size CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CLng(act("value"))
        Case "set_cell_font_color"
            modActionsTable.Do_set_cell_font_color CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CStr(act("value"))
        Case "set_cell_font_bold"
            modActionsTable.Do_set_cell_font_bold CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), modActions.ToBool(act("value"))
        Case "set_cell_font_italic"
            modActionsTable.Do_set_cell_font_italic CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), modActions.ToBool(act("value"))
        Case "set_cell_font_underline"
            modActionsTable.Do_set_cell_font_underline CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), modActions.ToBool(act("value"))
        Case "set_cell_font_name"
            modActionsTable.Do_set_cell_font_name CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CStr(act("value"))
        Case "set_cell_text_orientation"
            modActionsTable.Do_set_cell_text_orientation CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CStr(act("value"))
        Case "set_row_fill"
            modActionsTable.Do_set_row_fill CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CStr(act("value"))
        Case "set_column_fill"
            modActionsTable.Do_set_column_fill CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("col")), CStr(act("value"))
        Case "set_row_font_size"
            modActionsTable.Do_set_row_font_size CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("value"))
        Case "set_column_font_size"
            modActionsTable.Do_set_column_font_size CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("col")), CLng(act("value"))
        Case "set_row_font_color"
            modActionsTable.Do_set_row_font_color CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CStr(act("value"))
        Case "set_column_font_color"
            modActionsTable.Do_set_column_font_color CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("col")), CStr(act("value"))
        Case "set_row_font_bold"
            modActionsTable.Do_set_row_font_bold CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), modActions.ToBool(act("value"))
        Case "set_column_font_bold"
            modActionsTable.Do_set_column_font_bold CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("col")), modActions.ToBool(act("value"))
        Case "clear_row_text"
            modActionsTable.Do_clear_row_text CLng(act("slide")), CLng(act("shape_id")), CLng(act("row"))
        Case "clear_column_text"
            modActionsTable.Do_clear_column_text CLng(act("slide")), CLng(act("shape_id")), CLng(act("col"))
        Case "set_table_font_size"
            modActionsTable.Do_set_table_font_size CLng(act("slide")), CLng(act("shape_id")), CLng(act("value"))
        Case "set_table_font_name"
            modActionsTable.Do_set_table_font_name CLng(act("slide")), CLng(act("shape_id")), CStr(act("value"))
        Case "set_table_font_color"
            modActionsTable.Do_set_table_font_color CLng(act("slide")), CLng(act("shape_id")), CStr(act("value"))
        Case "auto_fit_table_text"
            modActionsTable.Do_auto_fit_table_text CLng(act("slide")), CLng(act("shape_id"))
        Case "set_table_borders"
            Dim tbBdColor As String: tbBdColor = ""
            Dim tbBdWeight As Single: tbBdWeight = 0
            Dim tbBdVisible As Boolean: tbBdVisible = True
            Dim tbBdDash As String: tbBdDash = ""
            If act.Exists("color") Then tbBdColor = CStr(act("color"))
            If act.Exists("weight_pt") Then tbBdWeight = CSng(act("weight_pt"))
            If act.Exists("visible") Then tbBdVisible = modActions.ToBool(act("visible"))
            If act.Exists("dash_style") Then tbBdDash = CStr(act("dash_style"))
            modActionsTable.Do_set_table_borders CLng(act("slide")), CLng(act("shape_id")), _
                CStr(act("side")), tbBdColor, tbBdWeight, tbBdVisible, tbBdDash
        Case "set_row_borders"
            Dim rbColor As String: rbColor = ""
            Dim rbWeight As Single: rbWeight = 0
            Dim rbVisible As Boolean: rbVisible = True
            Dim rbDash As String: rbDash = ""
            If act.Exists("color") Then rbColor = CStr(act("color"))
            If act.Exists("weight_pt") Then rbWeight = CSng(act("weight_pt"))
            If act.Exists("visible") Then rbVisible = modActions.ToBool(act("visible"))
            If act.Exists("dash_style") Then rbDash = CStr(act("dash_style"))
            modActionsTable.Do_set_row_borders CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CStr(act("side")), rbColor, rbWeight, rbVisible, rbDash
        Case "set_column_borders"
            Dim cbBdColor As String: cbBdColor = ""
            Dim cbBdWeight As Single: cbBdWeight = 0
            Dim cbBdVisible As Boolean: cbBdVisible = True
            Dim cbBdDash As String: cbBdDash = ""
            If act.Exists("color") Then cbBdColor = CStr(act("color"))
            If act.Exists("weight_pt") Then cbBdWeight = CSng(act("weight_pt"))
            If act.Exists("visible") Then cbBdVisible = modActions.ToBool(act("visible"))
            If act.Exists("dash_style") Then cbBdDash = CStr(act("dash_style"))
            modActionsTable.Do_set_column_borders CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("col")), CStr(act("side")), cbBdColor, cbBdWeight, cbBdVisible, cbBdDash
        Case "unmerge_cells"
            modActionsTable.Do_unmerge_cells CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col"))
        ' --- Cell paragraph actions ---------------------------------------
        Case "set_cell_paragraph_text"
            modActionsTable.Do_set_cell_paragraph_text CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CLng(act("paragraph_index")), CStr(act("value"))
        Case "set_cell_paragraph_font_size"
            modActionsTable.Do_set_cell_paragraph_font_size CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CLng(act("paragraph_index")), CLng(act("value"))
        Case "set_cell_paragraph_font_color"
            modActionsTable.Do_set_cell_paragraph_font_color CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CLng(act("paragraph_index")), CStr(act("value"))
        Case "set_cell_paragraph_bold"
            modActionsTable.Do_set_cell_paragraph_bold CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CLng(act("paragraph_index")), modActions.ToBool(act("value"))
        Case "set_cell_paragraph_italic"
            modActionsTable.Do_set_cell_paragraph_italic CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CLng(act("paragraph_index")), modActions.ToBool(act("value"))
        Case "set_cell_paragraph_alignment"
            modActionsTable.Do_set_cell_paragraph_alignment CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CLng(act("paragraph_index")), CStr(act("value"))
        Case "set_cell_bullet_style"
            modActionsTable.Do_set_cell_bullet_style CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CLng(act("paragraph_index")), CStr(act("value"))
        Case "add_cell_paragraph"
            modActionsTable.Do_add_cell_paragraph CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CLng(act("after_paragraph_index")), CStr(act("value"))
        Case "delete_cell_paragraph"
            modActionsTable.Do_delete_cell_paragraph CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CLng(act("paragraph_index"))
        Case "set_cell_indent_level"
            modActionsTable.Do_set_cell_indent_level CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CLng(act("paragraph_index")), CLng(act("value"))
        Case "append_cell_text"
            modActionsTable.Do_append_cell_text CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CStr(act("value"))
        Case "set_cell"
            modActionsTable.Do_set_cell_act act
        ' --- Shape fill / line / hyperlink / picture-fill ----------------
        Case "clear_fill"
            modActions.Do_clear_fill CLng(act("slide")), CLng(act("shape_id"))
        Case "clear_line"
            modActions.Do_clear_line CLng(act("slide")), CLng(act("shape_id"))
        Case "set_fill_visible"
            modActions.Do_set_fill_visible CLng(act("slide")), CLng(act("shape_id")), modActions.ToBool(act("value"))
        Case "set_line_visible"
            modActions.Do_set_line_visible CLng(act("slide")), CLng(act("shape_id")), modActions.ToBool(act("value"))
        Case "set_shape_hyperlink"
            modActions.Do_set_shape_hyperlink CLng(act("slide")), CLng(act("shape_id")), CStr(act("value"))
        Case "set_shape_picture_fill"
            modActions.Do_set_shape_picture_fill CLng(act("slide")), CLng(act("shape_id")), CStr(act("picture_path"))
        ' --- WAVE 3 ----------------------------------------------------
        Case "set_slide_transition"
            Dim stSpeed As String: stSpeed = ""
            Dim stOnClick As Boolean: stOnClick = True
            Dim stAdvSec As Double: stAdvSec = 0
            Dim stHasSpeed As Boolean: stHasSpeed = act.Exists("speed")
            Dim stHasOnClick As Boolean: stHasOnClick = act.Exists("advance_on_click")
            Dim stHasAdvSec As Boolean: stHasAdvSec = act.Exists("advance_after_seconds")
            If stHasSpeed Then stSpeed = CStr(act("speed"))
            If stHasOnClick Then stOnClick = modActions.ToBool(act("advance_on_click"))
            If stHasAdvSec Then stAdvSec = CDbl(act("advance_after_seconds"))
            modActionsSlide.Do_set_slide_transition CLng(act("slide")), CStr(act("effect")), _
                stSpeed, stOnClick, stAdvSec, stHasSpeed, stHasOnClick, stHasAdvSec
        Case "change_slide_layout"
            modActionsSlide.Do_change_slide_layout CLng(act("slide")), CLng(act("layout_index"))
        Case "add_section"
            modActionsSlide.Do_add_section CLng(act("before_slide")), CStr(act("name"))
        Case "delete_section"
            Dim secDelSlides As Boolean: secDelSlides = False
            If act.Exists("delete_slides") Then secDelSlides = modActions.ToBool(act("delete_slides"))
            modActionsSlide.Do_delete_section CLng(act("section_index")), secDelSlides
        Case "rename_section"
            modActionsSlide.Do_rename_section CLng(act("section_index")), CStr(act("name"))
        Case "move_section"
            modActionsSlide.Do_move_section CLng(act("section_index")), CLng(act("to_position"))
        Case "apply_picture_artistic_effect"
            Dim apeInt As Long: apeInt = 50
            Dim apeHasInt As Boolean: apeHasInt = act.Exists("intensity")
            If apeHasInt Then apeInt = CLng(act("intensity"))
            modActionsEffects.Do_apply_picture_artistic_effect CLng(act("slide")), CLng(act("shape_id")), _
                CStr(act("effect")), apeInt, apeHasInt
        Case "reset_picture"
            modActionsEffects.Do_reset_picture CLng(act("slide")), CLng(act("shape_id"))
        Case "set_shape_visible"
            modActionsEffects.Do_set_shape_visible CLng(act("slide")), CLng(act("shape_id")), _
                modActions.ToBool(act("value"))
        Case "reconnect_connector"
            Dim rcFromSite As Long: rcFromSite = 1
            Dim rcToSite As Long: rcToSite = 1
            Dim rcHasFrom As Boolean: rcHasFrom = act.Exists("from_connection_site")
            Dim rcHasTo As Boolean: rcHasTo = act.Exists("to_connection_site")
            If rcHasFrom Then rcFromSite = CLng(act("from_connection_site"))
            If rcHasTo Then rcToSite = CLng(act("to_connection_site"))
            modActionsEffects.Do_reconnect_connector CLng(act("slide")), CLng(act("shape_id")), _
                ResolveActShapeId(act, "from_shape_id"), ResolveActShapeId(act, "to_shape_id"), _
                rcFromSite, rcToSite, rcHasFrom, rcHasTo
        Case "set_run_kerning"
            modActionsRun.Do_set_run_kerning CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("paragraph_index")), CLng(act("run_index")), CDbl(act("value"))
        Case "set_run_baseline_offset"
            modActionsRun.Do_set_run_baseline_offset CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("paragraph_index")), CLng(act("run_index")), CDbl(act("value"))
        Case "set_bullet_start_number"
            modActionsText.Do_set_bullet_start_number CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("paragraph_index")), CLng(act("value"))
        Case "set_notes_font_size"
            modActions.Do_set_notes_font_size CLng(act("slide")), CLng(act("value"))
        Case "set_notes_font_color"
            modActions.Do_set_notes_font_color CLng(act("slide")), CStr(act("value"))
        Case "set_notes_font_bold"
            modActions.Do_set_notes_font_bold CLng(act("slide")), modActions.ToBool(act("value"))
        Case "set_notes_font_italic"
            modActions.Do_set_notes_font_italic CLng(act("slide")), modActions.ToBool(act("value"))
        Case "set_notes_font_name"
            modActions.Do_set_notes_font_name CLng(act("slide")), CStr(act("value"))
        Case "fit_cell_to_content"
            modActionsTable.Do_fit_cell_to_content CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col"))
        Case "set_data_label_text"
            modActionsChart.Do_set_data_label_text CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("series_index")), CLng(act("point_index")), CStr(act("value"))
        Case "apply_template"
            modActionsTemplate.Do_apply_template_act act
        Case "run_verification"
            ' Standalone mid-batch verification trigger. Writes sidecar JSON.
            Dim rvScope As String: rvScope = "deck"
            Dim rvMax As Long: rvMax = 100
            If act.Exists("scope") Then rvScope = CStr(act("scope"))
            If act.Exists("max_warnings") Then rvMax = CLng(act("max_warnings"))
            Dim rvWarns As Collection
            Set rvWarns = modVerify.RunVerificationLoop(rvScope, rvMax)
            On Error Resume Next
            Dim rvPath As String: rvPath = ActivePresentation.FullName & ".warnings.json"
            Dim fnum As Integer: fnum = FreeFile
            Open rvPath For Output As #fnum
            Print #fnum, "{""warnings"":" & modVerify.WarningsToJson(rvWarns) & "}"
            Close #fnum
            On Error GoTo 0
    End Select
End Sub

' Strip common LLM noise that surrounds JSON output:
' 1. Unicode BOM at the start
' 2. Smart / curly quotes -> ASCII double quotes
' 3. Markdown code fences ``` ```json ... ``` ```
' 4. Prose BEFORE the first { or [
' 5. Prose AFTER the last matching } or ]
' 6. JavaScript-style comments (// line and /* block */) outside strings
' 7. Trailing commas before } or ]
' Returns the cleaned JSON string. If no { or [ is found, returns the original
' input so the existing error path still surfaces a useful message.
' =============================================================================
' Pre-Apply error-fix helpers — when Parse shows validation errors (missing
' fields, bad enum values, etc.) the user can click "Fix Errors" on the form
' to get an LLM-ready prompt with the exact errors + canonical guidance.
' =============================================================================

' Build a paste-ready prompt that explains every invalid action in the batch
' plus the canonical signature/example for that action type. Returns "" if
' the batch parses cleanly with no validation errors.
Public Function BuildErrorFixPrompt(jsonText As String) As String
    BuildErrorFixPrompt = ""
    Dim cleaned As String: cleaned = SanitizeJsonInput(jsonText)
    Dim parsed As Object
    On Error Resume Next
    Set parsed = modJSON.ParseJson(cleaned)
    If Err.Number <> 0 Then
        BuildErrorFixPrompt = _
            "Decko could not parse the actions JSON you returned. " & _
            "PowerPoint says: " & Err.Description & vbCrLf & vbCrLf & _
            "FIX: emit a single valid JSON object of shape {""actions"":[...]} . " & _
            "Wrap the whole thing in one outer object; do not return a bare array. " & _
            "Strip any prose, markdown fences, or comments outside the JSON."
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    If Not parsed.Exists("actions") Then
        BuildErrorFixPrompt = "Missing top-level ""actions"" array. " & _
            "Wrap your actions in {""actions"":[...]}."
        Exit Function
    End If

    Dim actions As Object: Set actions = parsed("actions")
    Dim errors As Collection: Set errors = New Collection
    Dim i As Long
    For i = 1 To actions.Count
        Dim act As Object: Set act = actions(i)
        Dim reason As String: reason = PreviewValidate(act)
        If Len(reason) > 0 Then
            Dim ed As Object: Set ed = CreateObject("Scripting.Dictionary")
            ed("index") = i
            ed("type") = GetStr(act, "type")
            ed("reason") = reason
            ed("json") = modJSON.ConvertToJson(act)
            errors.Add ed
        End If
    Next i

    If errors.Count = 0 Then Exit Function   ' no errors — clipboard stays clean

    Dim sb As String
    sb = "Decko's parser found " & errors.Count & " action(s) in your batch with " & _
         "validation errors. For each one below, the ERROR field shows what's " & _
         "wrong, and the CORRECT SHAPE field shows the canonical signature + a " & _
         "working example. Return a single JSON object {""actions"":[...]} with " & _
         "all errors fixed. Keep the actions that were already valid; only rewrite " & _
         "the failing ones." & vbCrLf & vbCrLf

    For i = 1 To errors.Count
        Set ed = errors(i)
        sb = sb & "--- ACTION " & ed("index") & " (type: " & ed("type") & ") ---" & vbCrLf
        sb = sb & "YOU SENT: " & ed("json") & vbCrLf
        sb = sb & "ERROR: " & ed("reason") & vbCrLf
        ' If the error is unknown_type, suggest the closest known action
        ' names so the LLM doesn't invent more nonsense or give up.
        If LCase(Left(CStr(ed("reason")), 13)) = "unknown_type:" Then
            Dim suggestions As String
            suggestions = FindSimilarActions(CStr(ed("type")))
            If Len(suggestions) > 0 Then
                sb = sb & "DID YOU MEAN: " & suggestions & vbCrLf
            End If
        End If
        sb = sb & "CORRECT SHAPE:" & vbCrLf
        sb = sb & GetActionGuidance(CStr(ed("type"))) & vbCrLf & vbCrLf
    Next i
    BuildErrorFixPrompt = sb
End Function

' Score similarity between two action-type strings via word-stem overlap
' (split on '_'). Higher = more similar.
Private Function SimilarityScore(a As String, b As String) As Long
    If LCase(a) = LCase(b) Then SimilarityScore = 1000: Exit Function
    Dim aw() As String, bw() As String
    aw = Split(LCase(a), "_")
    bw = Split(LCase(b), "_")
    Dim score As Long: score = 0
    Dim i As Long, j As Long
    For i = 0 To UBound(aw)
        For j = 0 To UBound(bw)
            If aw(i) = bw(j) Then
                score = score + 10
            ElseIf Len(aw(i)) >= 4 And Len(bw(j)) >= 4 Then
                If InStr(aw(i), bw(j)) > 0 Or InStr(bw(j), aw(i)) > 0 Then _
                    score = score + 3
            End If
        Next j
    Next i
    ' Penalty for length mismatch (one is much longer than the other)
    Dim lenDelta As Long: lenDelta = Abs(Len(a) - Len(b))
    SimilarityScore = score - (lenDelta \ 4)
End Function

' Return top-3 known action types whose name is most similar to badType,
' comma-separated. Empty string if no match exceeds the noise floor.
Public Function FindSimilarActions(badType As String) As String
    Dim known As String: known = GetAllActionTypes()
    Dim names() As String: names = Split(known, ",")
    Dim n As Long: n = UBound(names) - LBound(names) + 1
    Dim top1 As String, top2 As String, top3 As String
    Dim s1 As Long, s2 As Long, s3 As Long
    s1 = 0: s2 = 0: s3 = 0
    Dim i As Long, score As Long
    For i = LBound(names) To UBound(names)
        Dim nm As String: nm = Trim(names(i))
        If Len(nm) = 0 Then GoTo NextI
        score = SimilarityScore(badType, nm)
        If score > s1 Then
            s3 = s2: top3 = top2
            s2 = s1: top2 = top1
            s1 = score: top1 = nm
        ElseIf score > s2 Then
            s3 = s2: top3 = top2
            s2 = score: top2 = nm
        ElseIf score > s3 Then
            s3 = score: top3 = nm
        End If
NextI:
    Next i
    ' Require at least one shared word stem (score >= 10) to suggest anything.
    Dim out As String: out = ""
    If s1 >= 10 Then out = top1
    If s2 >= 10 Then out = out & ", " & top2
    If s3 >= 10 Then out = out & ", " & top3
    FindSimilarActions = out
End Function

' Master list of every known action type. Used by FindSimilarActions for
' "did you mean" suggestions when the LLM invents an action name. Update
' when a new action is added to DispatchAction.
Public Function GetAllActionTypes() As String
    Dim s As String
    ' Core shape (modActions)
    s = "set_text,set_font_size,set_font_bold,set_font_italic,set_font_color,set_fill_color,"
    s = s & "move_shape,resize_shape,delete_shape,duplicate_shape,rotate_shape,flip_shape,"
    s = s & "set_shape_adjustment,z_order,copy_formatting,set_shape_name,set_pos,set_shape_alt_text,"
    s = s & "lock_aspect_ratio,clear_fill,clear_line,set_fill_visible,set_line_visible,"
    s = s & "set_shape_hyperlink,set_shape_picture_fill,set_shape_kind,set_shape_visible,"
    s = s & "add_shape,add_text_box,add_line,"
    ' Paragraph / run
    s = s & "set_paragraph_text,add_paragraph,delete_paragraph,set_bullet_style,set_indent_level,"
    s = s & "set_paragraph_font_size,set_paragraph_font_color,set_paragraph_alignment,"
    s = s & "set_paragraph_line_spacing,set_paragraph_bold,set_paragraph_italic,"
    s = s & "set_paragraph_underline,set_paragraph_font_name,set_paragraph_space_before,"
    s = s & "set_paragraph_space_after,clear_paragraph_formatting,set_bullet_start_number,"
    s = s & "set_run_bold,set_run_italic,set_run_underline,set_run_strikethrough,set_run_text,add_run,"
    s = s & "set_run_subscript,set_run_superscript,set_run_font_color,set_run_font_size,"
    s = s & "set_run_font_name,set_run_hyperlink,set_run_highlight,set_run_kerning,"
    s = s & "set_run_baseline_offset,"
    ' Text frame
    s = s & "set_text_vertical_align,set_text_autofit,set_text_margin,fit_to_content,"
    s = s & "enable_text_shrink_for_overflow,find_replace_text,find_replace_regex,"
    ' Layout
    s = s & "align_shapes,distribute_horizontal,distribute_vertical,tile_grid,smart_spacing,"
    s = s & "equalize_spacing,uniform_size,match_size,match_position,swap_positions,"
    s = s & "group_by_overlap,fit_to_slide_margins,move_shape_relative,nudge,snap_to_grid,"
    s = s & "align_to_slide_center,clear_slide,recolor_fill_match,recolor_font_match,"
    s = s & "delete_shapes_match,recolor_palette_deck_wide,recolor_deck,scan_palette,swap_font_deck_wide,"
    ' Connectors / groups
    s = s & "add_connector,reconnect_connector,group_shapes,ungroup,"
    GetAllActionTypes = s & GetAllActionTypes_Part2()
End Function

' Part 2 keeps each function under VBA's 24-continuation-line limit.
Private Function GetAllActionTypes_Part2() As String
    Dim s As String
    ' Tables
    s = "add_table,set_cell_text,add_table_row,delete_table_row,add_table_col,delete_table_col,"
    s = s & "swap_table_columns,swap_table_rows,merge_cells,unmerge_cells,"
    s = s & "set_table_col_width,set_table_row_height,set_cell_border,set_cell_text_align,"
    s = s & "set_cell_fill,apply_table_style,build_image_grid_table,set_cell_padding,"
    s = s & "clear_cell_text,set_table_style_options,populate_table_row,populate_table_column,"
    s = s & "populate_table_cells,set_cell_font_size,set_cell_font_color,set_cell_font_bold,"
    s = s & "set_cell_font_italic,set_cell_font_underline,set_cell_font_name,"
    s = s & "set_cell_text_orientation,set_row_fill,set_column_fill,set_row_font_size,"
    s = s & "set_column_font_size,set_row_font_color,set_column_font_color,set_row_font_bold,"
    s = s & "set_column_font_bold,clear_row_text,clear_column_text,set_table_font_size,"
    s = s & "set_table_font_name,set_table_font_color,auto_fit_table_text,set_table_borders,"
    s = s & "set_row_borders,set_column_borders,fit_cell_to_content,set_cell_paragraph_text,"
    s = s & "set_cell_paragraph_font_size,set_cell_paragraph_font_color,set_cell_paragraph_bold,"
    s = s & "set_cell_paragraph_italic,set_cell_paragraph_alignment,set_cell_bullet_style,"
    s = s & "add_cell_paragraph,delete_cell_paragraph,set_cell_indent_level,append_cell_text,set_cell,"
    GetAllActionTypes_Part2 = s & GetAllActionTypes_Part3()
End Function

Private Function GetAllActionTypes_Part3() As String
    Dim s As String
    ' Charts
    s = "add_chart,set_chart_type,set_chart_title,set_chart_axis_title,"
    s = s & "set_chart_legend_position,set_chart_legend,set_series_color,set_series_values,"
    s = s & "set_chart_categories,set_series_name,set_chart_axis,set_chart_gridlines,"
    s = s & "set_chart_format,set_chart_series,add_chart_trendline,set_chart_error_bars,"
    s = s & "set_chart_data_table,set_line_smoothing,delete_series,add_series,set_data_label_text,"
    ' Images / web
    s = s & "insert_picture,replace_picture,insert_icon,fetch_page_images,download_image,"
    s = s & "open_image_picker,build_image_picker_slide,bulk_insert_image,"
    ' Slides / deck
    s = s & "add_slide,delete_slide,duplicate_slide,move_slide,extract_slides,"
    s = s & "import_slides_from_deck,apply_layout_to_slides,change_slide_layout,apply_theme,"
    s = s & "set_theme_font,set_slide_size,bulk_insert_text_box,set_slide_background_color,"
    s = s & "insert_slide_number,set_slide_hidden,set_slide_name,set_slide_transition,"
    s = s & "add_section,delete_section,rename_section,move_section,"
    ' Notes
    s = s & "set_speaker_notes,append_speaker_notes,clear_speaker_notes,set_notes_font_size,"
    s = s & "set_notes_font_color,set_notes_font_bold,set_notes_font_italic,set_notes_font_name,"
    ' Effects
    s = s & "set_line_color,set_line_weight,set_line_style,set_shadow,set_glow,set_reflection,"
    s = s & "set_transparency,set_gradient_fill,set_3d_bevel,set_3d_rotation,set_soft_edge,"
    s = s & "apply_preset_effect,crop_picture,recolor_picture,set_brightness,set_contrast,"
    s = s & "apply_picture_artistic_effect,reset_picture,clear_shadow,clear_glow,"
    s = s & "clear_reflection,clear_all_effects,run_verification,apply_template"
    GetAllActionTypes_Part3 = s
End Function

' Canonical signature + example for the most-misused action types. For types
' not listed here, returns a pointer to ACTIONS_REFERENCE.md. Keep table tight
' (top ~35 types) so the prompt stays focused.
Public Function GetActionGuidance(actionType As String) As String
    Select Case LCase(actionType)
        Case "set_text"
            GetActionGuidance = _
                "  REQUIRED: slide(int), shape_id(int|ref_name), value(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_text"",""slide"":1,""shape_id"":3,""value"":""Q3 Revenue""}" & vbCrLf & _
                "  NOTE: destroys per-paragraph formatting. Use set_paragraph_text for bullet lists."
        Case "set_paragraph_text"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index (0-based int), value(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_paragraph_text"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":""Hello""}"
        Case "add_paragraph"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, after_paragraph_index(int; -1 prepends), value(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_paragraph"",""slide"":1,""shape_id"":3,""after_paragraph_index"":2,""value"":""New bullet""}"
        Case "delete_paragraph"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index" & vbCrLf & _
                "  EXAMPLE:  {""type"":""delete_paragraph"",""slide"":1,""shape_id"":3,""paragraph_index"":2}"
        Case "set_font_size"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(int>0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_font_size"",""slide"":1,""shape_id"":3,""value"":14}"
        Case "set_font_color", "set_fill_color"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(#RRGGBB hex string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":1,""shape_id"":3,""value"":""#15283C""}"
        Case "set_font_bold", "set_font_italic"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(bool)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":1,""shape_id"":3,""value"":true}"
        Case "set_paragraph_font_size"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, value(int>0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_paragraph_font_size"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":12}"
        Case "set_paragraph_font_color"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, value(#RRGGBB)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_paragraph_font_color"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":""#15283C""}"
        Case "set_run_bold", "set_run_italic", "set_run_underline", "set_run_strikethrough"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, run_index, value(bool)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":1,""shape_id"":3,""paragraph_index"":0,""run_index"":1,""value"":true}" & vbCrLf & _
                "  NOTE: paragraph_index AND run_index are both 0-based."
        Case "set_run_font_color"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, run_index, value(#RRGGBB)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_run_font_color"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""run_index"":1,""value"":""#15283C""}" & vbCrLf & _
                "  NOTE: paragraph_index AND run_index are both 0-based."
        Case "set_run_font_size"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, run_index, value(int>0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_run_font_size"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""run_index"":1,""value"":12}" & vbCrLf & _
                "  NOTE: paragraph_index AND run_index are both 0-based."
        Case "set_run_text"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, run_index, value(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_run_text"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""run_index"":1,""value"":""Revenue""}" & vbCrLf & _
                "  NOTE: paragraph_index AND run_index are both 0-based."
        Case "add_run"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, value(string)" & vbCrLf & _
                "  OPTIONAL: bold(bool), italic(bool), underline(bool), color(#RRGGBB), font_name(string), font_size(int)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_run"",""slide"":1,""shape_id"":3,""paragraph_index"":1,""value"":""18% YoY"",""bold"":true,""color"":""#C00000""}" & vbCrLf & _
                "  NOTE: appends a new run at the END of the paragraph; does not rebuild tr.Text so existing run formatting is preserved."
        Case "set_run_font_name"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, run_index, value(string, non-empty font name)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_run_font_name"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""run_index"":1,""value"":""Calibri""}"
        Case "add_shape"
            GetActionGuidance = _
                "  REQUIRED: slide, kind(string), pos({left,top,width,height})" & vbCrLf & _
                "  OPTIONAL: fill, stroke, text, font_size, font_color, h_align, v_align, ref_name" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_shape"",""slide"":1,""kind"":""rrect"",""pos"":{""left"":60,""top"":120,""width"":200,""height"":80},""fill"":""#15283C"",""text"":""Phase 1"",""font_color"":""#FFFFFF""}"
        Case "add_text_box"
            GetActionGuidance = _
                "  REQUIRED: slide, text(string — NOT 'value'), pos" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_text_box"",""slide"":1,""text"":""Label"",""pos"":{""left"":60,""top"":120,""width"":200,""height"":40}}"
        Case "add_line"
            GetActionGuidance = _
                "  REQUIRED: slide, x1, y1, x2, y2, color(#RRGGBB), weight_pt" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_line"",""slide"":1,""x1"":60,""y1"":100,""x2"":300,""y2"":100,""color"":""#15283C"",""weight_pt"":1.5}"
        Case "move_shape"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, left(num), top(num)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""move_shape"",""slide"":1,""shape_id"":3,""left"":100,""top"":120}"
        Case "resize_shape"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, width(num), height(num)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""resize_shape"",""slide"":1,""shape_id"":3,""width"":300,""height"":200}"
        Case "delete_shape"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id" & vbCrLf & _
                "  EXAMPLE:  {""type"":""delete_shape"",""slide"":1,""shape_id"":3}"
        Case "add_slide"
            GetActionGuidance = _
                "  REQUIRED: position(int, 1-based), layout_index(int, 0-based; 6 = blank)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_slide"",""position"":3,""layout_index"":6}"
        Case "delete_slide", "duplicate_slide"
            GetActionGuidance = _
                "  REQUIRED: slide" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":3}"
        Case "set_cell_text"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id (the table), row(1-based int), col(1-based int), value(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell_text"",""slide"":1,""shape_id"":4,""row"":2,""col"":3,""value"":""$1.2B""}"
        Case "populate_table_row"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row(1-based), values(array of strings; one per column)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""populate_table_row"",""slide"":1,""shape_id"":4,""row"":3,""values"":[""Q1"",""$1.2B"",""+12%""]}" & vbCrLf & _
                "  TIP: use this instead of N separate set_cell_text calls — avoids column-shift bugs."
        Case "add_table"
            GetActionGuidance = _
                "  REQUIRED: slide, rows(int), cols(int), pos({left,top,width,height})" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_table"",""slide"":1,""rows"":4,""cols"":3,""pos"":{""left"":60,""top"":120,""width"":600,""height"":300},""ref_name"":""tbl1""}"
        Case "add_chart"
            GetActionGuidance = _
                "  REQUIRED: slide, chart_type(string), pos, categories(array), series(array of {name, values, color?})" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_chart"",""slide"":1,""chart_type"":""columnclustered""," & _
                """pos"":{""left"":60,""top"":120,""width"":560,""height"":340}," & _
                """categories"":[""FY22"",""FY23"",""FY24""]," & _
                """series"":[{""name"":""Revenue ($M)"",""values"":[120,138,151],""color"":""#15283C""}]}" & vbCrLf & _
                "  NOTE: each series.values length MUST equal categories length."
        Case "find_replace_text"
            GetActionGuidance = _
                "  REQUIRED: scope(""deck"" or ""slide:N""), find(string), replace(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""find_replace_text"",""scope"":""deck"",""find"":""Acme"",""replace"":""NewCo""}"
        Case "set_paragraph_alignment"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, value(""left""|""center""|""right""|""justify"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_paragraph_alignment"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":""center""}"
        Case "set_bullet_style"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, value(""none""|""disc""|""square""|""dash""|""number""|""letter"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_bullet_style"",""slide"":1,""shape_id"":3,""paragraph_index"":1,""value"":""disc""}"
        Case "set_indent_level"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, value(int 0..4)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_indent_level"",""slide"":1,""shape_id"":3,""paragraph_index"":2,""value"":1}"
        Case "set_chart_series"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, series_index(1-based), props(object)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_chart_series"",""slide"":1,""shape_id"":2,""series_index"":1," & _
                """props"":{""fill"":""#15283C"",""show_labels"":true,""label_color"":""#FFFFFF""}}"
        Case "insert_picture"
            GetActionGuidance = _
                "  REQUIRED: slide, pos, path (or picture_path; both accepted)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""insert_picture"",""slide"":1,""path"":""C:\\imgs\\logo.png""," & _
                """pos"":{""left"":60,""top"":120,""width"":200,""height"":120}}"
        Case "insert_icon"
            GetActionGuidance = _
                "  REQUIRED: slide, icon(lowercase_underscore name), left, top, width, height (ALL four required, NO pos object)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""insert_icon"",""slide"":1,""icon"":""building_factory"",""left"":60,""top"":120,""width"":48,""height"":48,""color"":""#15283C""}"
        Case "add_connector"
            GetActionGuidance = _
                "  REQUIRED: slide, kind(""straight""|""elbow""|""curved""), from_shape_id (or from_shape_name), to_shape_id (or to_shape_name)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_connector"",""slide"":1,""kind"":""elbow"",""from_shape_name"":""box1"",""to_shape_name"":""box2"",""arrow_end"":""filled""}"
        Case "set_speaker_notes", "append_speaker_notes"
            GetActionGuidance = _
                "  REQUIRED: slide, value(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":3,""value"":""Mention Q3 EBITDA expansion.""}"
        Case "clear_speaker_notes"
            GetActionGuidance = _
                "  REQUIRED: slide" & vbCrLf & _
                "  EXAMPLE:  {""type"":""clear_speaker_notes"",""slide"":3}"
        ' ---- shape geometry / state ----
        Case "duplicate_shape"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, left(num), top(num)" & vbCrLf & _
                "  OPTIONAL: ref_name" & vbCrLf & _
                "  EXAMPLE:  {""type"":""duplicate_shape"",""slide"":1,""shape_id"":3,""left"":400,""top"":120}"
        Case "rotate_shape"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, degrees(num)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""rotate_shape"",""slide"":1,""shape_id"":3,""degrees"":45}"
        Case "flip_shape"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, axis(""h""|""v"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""flip_shape"",""slide"":1,""shape_id"":3,""axis"":""h""}"
        Case "set_shape_adjustment"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, index(1-based int), value(num, usually 0.0-1.0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_shape_adjustment"",""slide"":1,""shape_id"":3,""index"":1,""value"":0.25}"
        Case "z_order"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, order(""front""|""back""|""forward""|""backward"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""z_order"",""slide"":1,""shape_id"":3,""order"":""front""}"
        Case "copy_formatting"
            GetActionGuidance = _
                "  REQUIRED: slide, source_shape_id, target_shape_id" & vbCrLf & _
                "  EXAMPLE:  {""type"":""copy_formatting"",""slide"":1,""source_shape_id"":3,""target_shape_id"":5}"
        Case "set_shape_name"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(string, non-empty, unique on slide)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_shape_name"",""slide"":1,""shape_id"":3,""value"":""hero_card""}"
        Case "set_pos"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, AND at least one of: left, top, width, height (all num pt)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_pos"",""slide"":1,""shape_id"":3,""left"":100,""top"":120,""width"":300,""height"":200}"
        Case "set_shape_alt_text"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(string; """" clears)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_shape_alt_text"",""slide"":1,""shape_id"":3,""value"":""Q3 revenue chart""}"
        Case "lock_aspect_ratio", "set_shape_visible", "set_fill_visible", "set_line_visible"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(bool)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":1,""shape_id"":3,""value"":true}"
        Case "clear_fill", "clear_line", "clear_shadow", "clear_glow", "clear_reflection", "clear_all_effects"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":1,""shape_id"":3}"
        Case "set_shape_hyperlink"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(URL starting http://|https://|mailto:|#slide:N; or """" to clear)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_shape_hyperlink"",""slide"":1,""shape_id"":3,""value"":""#slide:5""}"
        Case "set_shape_picture_fill"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, picture_path(absolute local file path)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_shape_picture_fill"",""slide"":1,""shape_id"":3,""picture_path"":""C:\\imgs\\hero.jpg""}"
        Case "set_shape_kind"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, kind(string — see add_shape kind vocabulary)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_shape_kind"",""slide"":1,""shape_id"":3,""kind"":""rrect""}"
        ' ---- paragraph-level granular ----
        Case "set_paragraph_bold", "set_paragraph_italic", "set_paragraph_underline"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, value(bool)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":true}"
        Case "set_paragraph_font_name"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, value(string, non-empty)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_paragraph_font_name"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":""Calibri""}"
        Case "set_paragraph_space_before", "set_paragraph_space_after"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, value(num pt, >=0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":1,""shape_id"":3,""paragraph_index"":1,""value"":6}"
        Case "set_paragraph_line_spacing"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, value(num, multiple e.g. 1.0, 1.5)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_paragraph_line_spacing"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":1.15}"
        Case "clear_paragraph_formatting"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index" & vbCrLf & _
                "  EXAMPLE:  {""type"":""clear_paragraph_formatting"",""slide"":1,""shape_id"":3,""paragraph_index"":0}"
        Case "set_bullet_start_number"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, value(int >=1)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_bullet_start_number"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":5}"
        ' ---- run-level extras ----
        Case "set_run_subscript", "set_run_superscript"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, run_index, value(bool)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":1,""shape_id"":3,""paragraph_index"":0,""run_index"":1,""value"":true}"
        Case "set_run_hyperlink"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, run_index, value(URL: http://|https://|mailto:|#slide:N; or """" to clear)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_run_hyperlink"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""run_index"":1,""value"":""https://example.com""}"
        Case "set_run_highlight"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, run_index, value(#RRGGBB or """" to clear)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_run_highlight"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""run_index"":1,""value"":""#FFF59D""}"
        Case "set_run_kerning"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, run_index, value(num pt; 0=default, +=wider, -=tighter)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_run_kerning"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""run_index"":0,""value"":1.5}"
        Case "set_run_baseline_offset"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, run_index, value(num -1.0..1.0; fraction of font height)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_run_baseline_offset"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""run_index"":1,""value"":0.3}"
        Case "delete_paragraph"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index" & vbCrLf & _
                "  EXAMPLE:  {""type"":""delete_paragraph"",""slide"":1,""shape_id"":3,""paragraph_index"":2}"
        ' ---- text-frame behavior ----
        Case "set_text_vertical_align"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(""top""|""middle""|""bottom"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_text_vertical_align"",""slide"":1,""shape_id"":3,""value"":""middle""}"
        Case "set_text_autofit"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, mode(""none""|""shrink""|""resize"")  -- note: field is 'mode' NOT 'value'" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_text_autofit"",""slide"":1,""shape_id"":3,""mode"":""shrink""}"
        Case "set_text_margin"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, left(num>=0), right(num>=0), top(num>=0), bottom(num>=0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_text_margin"",""slide"":1,""shape_id"":3,""left"":4,""right"":4,""top"":2,""bottom"":2}"
        Case "fit_to_content"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id" & vbCrLf & _
                "  EXAMPLE:  {""type"":""fit_to_content"",""slide"":1,""shape_id"":3}"
        Case "enable_text_shrink_for_overflow"
            GetActionGuidance = _
                "  REQUIRED: scope(""deck""|""slide:N"")" & vbCrLf & _
                "  OPTIONAL: include_titles(bool)=false" & vbCrLf & _
                "  EXAMPLE:  {""type"":""enable_text_shrink_for_overflow"",""scope"":""slide:3""}"
        ' ---- find/replace ----
        Case "find_replace_regex"
            GetActionGuidance = _
                "  REQUIRED: scope(""deck""|""slide:N""), pattern(regex string), replacement(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""find_replace_regex"",""scope"":""deck"",""pattern"":""\\$\\d+M"",""replacement"":""TBD""}"
        ' ---- layout / alignment ----
        Case "align_shapes"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_ids(array of int/ref_name), anchor(""left""|""right""|""top""|""bottom""|""hcenter""|""vcenter"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""align_shapes"",""slide"":1,""shape_ids"":[3,4,5],""anchor"":""top""}"
        Case "distribute_horizontal", "distribute_vertical"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_ids(array of >=3 elements)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":1,""shape_ids"":[3,4,5,6]}"
        Case "tile_grid"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_ids(array), cols(int), gap_pt(num)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""tile_grid"",""slide"":1,""shape_ids"":[3,4,5,6],""cols"":2,""gap_pt"":12}"
        Case "smart_spacing", "equalize_spacing"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_ids(array), gap_pt(num, smart only), axis(""h""|""v"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":1,""shape_ids"":[3,4,5],""gap_pt"":10,""axis"":""h""}"
        Case "uniform_size"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_ids(array), width_pt(num>0), height_pt(num>0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""uniform_size"",""slide"":1,""shape_ids"":[3,4,5],""width_pt"":200,""height_pt"":80}"
        Case "match_size"
            GetActionGuidance = _
                "  REQUIRED: slide, ref_shape_id, target_shape_ids(array)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""match_size"",""slide"":1,""ref_shape_id"":3,""target_shape_ids"":[4,5,6]}"
        Case "match_position"
            GetActionGuidance = _
                "  REQUIRED: slide, ref_shape_id, target_shape_id, edge(""left""|""right""|""top""|""bottom""|""hcenter""|""vcenter"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""match_position"",""slide"":1,""ref_shape_id"":3,""target_shape_id"":4,""edge"":""left""}"
        Case "swap_positions"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_a_id, shape_b_id" & vbCrLf & _
                "  EXAMPLE:  {""type"":""swap_positions"",""slide"":1,""shape_a_id"":3,""shape_b_id"":4}"
        Case "group_by_overlap"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_ids(array)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""group_by_overlap"",""slide"":1,""shape_ids"":[3,4,5,6]}"
        Case "fit_to_slide_margins"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id" & vbCrLf & _
                "  OPTIONAL: margin_pt(num)=36" & vbCrLf & _
                "  EXAMPLE:  {""type"":""fit_to_slide_margins"",""slide"":1,""shape_id"":3,""margin_pt"":24}"
        Case "move_shape_relative"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, dx_pt(num), dy_pt(num)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""move_shape_relative"",""slide"":1,""shape_id"":3,""dx_pt"":10,""dy_pt"":-5}"
        Case "nudge"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, direction(""l""|""r""|""u""|""d""), amount_pt(num>=0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""nudge"",""slide"":1,""shape_id"":3,""direction"":""r"",""amount_pt"":5}"
        Case "snap_to_grid"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, grid_pt(num>0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""snap_to_grid"",""slide"":1,""shape_id"":3,""grid_pt"":12}"
        Case "align_to_slide_center"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, axis(""h""|""v""|""both"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""align_to_slide_center"",""slide"":1,""shape_id"":3,""axis"":""both""}"
        Case "clear_slide"
            GetActionGuidance = _
                "  REQUIRED: slide" & vbCrLf & _
                "  OPTIONAL: keep_shape_ids(array of int/ref_name)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""clear_slide"",""slide"":3,""keep_shape_ids"":[2]}"
        ' ---- recolor / match-delete ----
        Case "recolor_fill_match", "recolor_font_match"
            GetActionGuidance = _
                "  REQUIRED: scope(""deck""|""slide:N""), from(#RRGGBB), to(#RRGGBB)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""scope"":""deck"",""from"":""#FF0000"",""to"":""#15283C""}"
        Case "delete_shapes_match"
            GetActionGuidance = _
                "  REQUIRED: scope(""deck""|""slide:N""), AND at least one filter: kind(string), fill(#RRGGBB), text_contains(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""delete_shapes_match"",""scope"":""slide:3"",""kind"":""rectangle"",""fill"":""#CCCCCC""}"
        Case "recolor_palette_deck_wide"
            GetActionGuidance = _
                "  REQUIRED: from_hex(#RRGGBB), to_hex(#RRGGBB), target(""fill""|""font""|""both"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""recolor_palette_deck_wide"",""from_hex"":""#FF0000"",""to_hex"":""#15283C"",""target"":""both""}"
        Case "recolor_deck"
            GetActionGuidance = _
                "  Batch palette remap — N from->to pairs in one deck pass. Covers shape fill/border/font, table fill/border/font, chart series, slide backgrounds, groups." & vbCrLf & _
                "  REQUIRED: mappings(array of {from:#RRGGBB, to:#RRGGBB})" & vbCrLf & _
                "  OPTIONAL: scope(""all""|""fill""|""font""|""border""|""table_fill""|""table_font""|""table_border""|""chart"") default all" & vbCrLf & _
                "  EXAMPLE:  {""type"":""recolor_deck"",""mappings"":[{""from"":""#FF0000"",""to"":""#003087""},{""from"":""#FFFFFF"",""to"":""#F5F5F5""}]}"
        Case "scan_palette"
            GetActionGuidance = _
                "  Scan active deck for all explicit RGB colors. Writes role-tagged JSON to Windows clipboard AND to %TEMP%\decko_palette.json." & vbCrLf & _
                "  Use before recolor_deck to discover what colors to remap." & vbCrLf & _
                "  NO REQUIRED FIELDS." & vbCrLf & _
                "  OPTIONAL: scope(""deck"" default | ""slide:N"" for single slide)" & vbCrLf & _
                "  OUTPUT: JSON array [{""hex"":""#RRGGBB"",""count"":N,""roles"":[""fill""|""font""|""border""]}] sorted by count desc" & vbCrLf & _
                "  EXAMPLE:  {""type"":""scan_palette""}" & vbCrLf & _
                "  EXAMPLE:  {""type"":""scan_palette"",""scope"":""slide:1""}"
        ' ---- connectors / groups ----
        Case "group_shapes"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_ids(array of >=2 elements)" & vbCrLf & _
                "  OPTIONAL: ref_name" & vbCrLf & _
                "  EXAMPLE:  {""type"":""group_shapes"",""slide"":1,""shape_ids"":[3,4,5],""ref_name"":""logo_group""}"
        Case "ungroup"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id (the group shape)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""ungroup"",""slide"":1,""shape_id"":7}"
        Case "reconnect_connector"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id (the connector), from_shape_id, to_shape_id" & vbCrLf & _
                "  OPTIONAL: from_connection_site(int), to_connection_site(int)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""reconnect_connector"",""slide"":1,""shape_id"":7,""from_shape_id"":3,""to_shape_id"":5}"
        ' ---- tables — granular ----
        Case "add_table_row"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, after_row(int; 0 = before first)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_table_row"",""slide"":1,""shape_id"":4,""after_row"":2}"
        Case "delete_table_row"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row(1-based)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""delete_table_row"",""slide"":1,""shape_id"":4,""row"":3}"
        Case "add_table_col"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, after_col(int; 0 = before first)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_table_col"",""slide"":1,""shape_id"":4,""after_col"":1}"
        Case "delete_table_col"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, col(1-based)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""delete_table_col"",""slide"":1,""shape_id"":4,""col"":2}"
        Case "swap_table_columns"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, col_a(1-based), col_b(1-based)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""swap_table_columns"",""slide"":1,""shape_id"":4,""col_a"":1,""col_b"":3}"
        Case "swap_table_rows"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row_a(1-based), row_b(1-based)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""swap_table_rows"",""slide"":1,""shape_id"":4,""row_a"":2,""row_b"":4}"
        Case "merge_cells"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row_a, col_a, row_b, col_b (all 1-based)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""merge_cells"",""slide"":1,""shape_id"":4,""row_a"":1,""col_a"":1,""row_b"":1,""col_b"":3}"
        Case "unmerge_cells"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row(1-based), col(1-based) — any cell in the merged region" & vbCrLf & _
                "  EXAMPLE:  {""type"":""unmerge_cells"",""slide"":1,""shape_id"":4,""row"":1,""col"":1}"
        Case "set_table_col_width"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, col(1-based), width_pt(num>0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_table_col_width"",""slide"":1,""shape_id"":4,""col"":2,""width_pt"":180}"
        Case "set_table_row_height"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row(1-based), height_pt(num>0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_table_row_height"",""slide"":1,""shape_id"":4,""row"":1,""height_pt"":36}"
        Case "set_cell_border"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, side(""top""|""left""|""bottom""|""right""|""diag_down""|""diag_up""|""all"")" & vbCrLf & _
                "  OPTIONAL: color(#RRGGBB), weight_pt(num), visible(bool)=true, dash_style" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell_border"",""slide"":1,""shape_id"":4,""row"":2,""col"":3,""side"":""all"",""color"":""#15283C"",""weight_pt"":0.75}"
        Case "set_cell_text_align"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, AND at least one of: h_align(""left""|""center""|""right""), v_align(""top""|""middle""|""bottom"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell_text_align"",""slide"":1,""shape_id"":4,""row"":2,""col"":3,""h_align"":""center"",""v_align"":""middle""}"
        Case "set_cell_fill"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, color(#RRGGBB)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell_fill"",""slide"":1,""shape_id"":4,""row"":1,""col"":1,""color"":""#15283C""}"
        Case "apply_table_style"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, style_id(string — lowercase_underscore name like 'medium_style_2_accent1' OR Office GUID)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""apply_table_style"",""slide"":1,""shape_id"":4,""style_id"":""medium_style_2_accent1""}"
        Case "set_cell_padding"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, left, right, top, bottom (all num pt, >=0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell_padding"",""slide"":1,""shape_id"":4,""row"":1,""col"":1,""left"":4,""right"":4,""top"":2,""bottom"":2}"
        Case "clear_cell_text"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col" & vbCrLf & _
                "  EXAMPLE:  {""type"":""clear_cell_text"",""slide"":1,""shape_id"":4,""row"":2,""col"":3}"
        Case "set_table_style_options"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, AND at least one of: header_row, total_row, banded_rows, first_column, last_column, banded_columns (all bool)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_table_style_options"",""slide"":1,""shape_id"":4,""header_row"":true,""banded_rows"":true}"
        Case "populate_table_column"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, col(1-based), values(array of strings; one per row starting at row 1)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""populate_table_column"",""slide"":1,""shape_id"":4,""col"":1,""values"":[""Q1"",""Q2"",""Q3"",""Q4""]}"
        Case "populate_table_cells"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, start_row(1-based), start_col(1-based), values(2D array — outer=rows, inner=cells)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""populate_table_cells"",""slide"":1,""shape_id"":4,""start_row"":2,""start_col"":1,""values"":[[""Q1"",""$1.2B""],[""Q2"",""$1.4B""]]}"
        Case "set_cell_font_size"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, value(int>0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell_font_size"",""slide"":1,""shape_id"":4,""row"":1,""col"":1,""value"":12}"
        Case "set_cell_font_color"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, value(#RRGGBB)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell_font_color"",""slide"":1,""shape_id"":4,""row"":1,""col"":1,""value"":""#FFFFFF""}"
        Case "set_cell_font_bold", "set_cell_font_italic", "set_cell_font_underline"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, value(bool)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":1,""shape_id"":4,""row"":1,""col"":1,""value"":true}"
        Case "set_cell_font_name"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, value(string, non-empty)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell_font_name"",""slide"":1,""shape_id"":4,""row"":1,""col"":1,""value"":""Calibri""}"
        Case "set_cell_text_orientation"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, value(""horizontal""|""vertical_90""|""vertical_270""|""stacked"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell_text_orientation"",""slide"":1,""shape_id"":4,""row"":1,""col"":2,""value"":""vertical_90""}"
        Case "set_row_fill", "set_row_font_color"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row(1-based), value(#RRGGBB)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":1,""shape_id"":4,""row"":1,""value"":""#15283C""}"
        Case "set_column_fill", "set_column_font_color"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, col(1-based), value(#RRGGBB)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":1,""shape_id"":4,""col"":1,""value"":""#15283C""}"
        Case "set_row_font_size"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row(1-based), value(int>0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_row_font_size"",""slide"":1,""shape_id"":4,""row"":1,""value"":12}"
        Case "set_column_font_size"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, col(1-based), value(int>0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_column_font_size"",""slide"":1,""shape_id"":4,""col"":1,""value"":10}"
        Case "set_row_font_bold"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row(1-based), value(bool)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_row_font_bold"",""slide"":1,""shape_id"":4,""row"":1,""value"":true}"
        Case "set_column_font_bold"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, col(1-based), value(bool)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_column_font_bold"",""slide"":1,""shape_id"":4,""col"":1,""value"":true}"
        Case "clear_row_text"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row(1-based)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""clear_row_text"",""slide"":1,""shape_id"":4,""row"":3}"
        Case "clear_column_text"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, col(1-based)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""clear_column_text"",""slide"":1,""shape_id"":4,""col"":2}"
        Case "set_table_font_size"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(int>0)  -- applies to every cell" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_table_font_size"",""slide"":1,""shape_id"":4,""value"":10}"
        Case "set_table_font_name"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(string, non-empty)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_table_font_name"",""slide"":1,""shape_id"":4,""value"":""Calibri""}"
        Case "set_table_font_color"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(#RRGGBB)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_table_font_color"",""slide"":1,""shape_id"":4,""value"":""#15283C""}"
        Case "auto_fit_table_text"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id  -- enables shrink-to-fit on every cell" & vbCrLf & _
                "  EXAMPLE:  {""type"":""auto_fit_table_text"",""slide"":1,""shape_id"":4}"
        Case "set_table_borders"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, side(""top""|""left""|""bottom""|""right""|""diag_down""|""diag_up""|""all"")" & vbCrLf & _
                "  OPTIONAL: color, weight_pt, visible(bool)=true, dash_style" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_table_borders"",""slide"":1,""shape_id"":4,""side"":""all"",""color"":""#15283C"",""weight_pt"":0.5}"
        Case "set_row_borders"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row(1-based), side" & vbCrLf & _
                "  OPTIONAL: color, weight_pt, visible, dash_style" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_row_borders"",""slide"":1,""shape_id"":4,""row"":1,""side"":""bottom"",""color"":""#15283C"",""weight_pt"":1.5}"
        Case "set_column_borders"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, col(1-based), side" & vbCrLf & _
                "  OPTIONAL: color, weight_pt, visible, dash_style" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_column_borders"",""slide"":1,""shape_id"":4,""col"":1,""side"":""right""}"
        Case "fit_cell_to_content"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col" & vbCrLf & _
                "  EXAMPLE:  {""type"":""fit_cell_to_content"",""slide"":1,""shape_id"":4,""row"":2,""col"":3}"
        Case "build_image_grid_table"
            GetActionGuidance = _
                "  REQUIRED: slide, pos({left,top,width,height}), rows(array of row objects)" & vbCrLf & _
                "  Each row object: {name, image_path OR image_url, bullets:[strings]}" & vbCrLf & _
                "  EXAMPLE:  {""type"":""build_image_grid_table"",""slide"":1,""pos"":{""left"":60,""top"":120,""width"":800,""height"":400}," & _
                """rows"":[{""name"":""John Smith"",""image_path"":""C:/imgs/j.png"",""bullets"":[""Coverage MD"",""12 yrs""]}]}" & vbCrLf & _
                "  See ACTIONS_REFERENCE.md §3.12 for full schema (image_col, name_position, name_font, etc.)"
        ' ---- cell paragraph actions ----
        Case "set_cell_paragraph_text"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, paragraph_index(0-based), value(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell_paragraph_text"",""slide"":1,""shape_id"":4,""row"":1,""col"":1,""paragraph_index"":0,""value"":""Header""}"
        Case "set_cell_paragraph_font_size"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, paragraph_index, value(int>0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell_paragraph_font_size"",""slide"":1,""shape_id"":4,""row"":1,""col"":1,""paragraph_index"":0,""value"":12}"
        Case "set_cell_paragraph_font_color"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, paragraph_index, value(#RRGGBB)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell_paragraph_font_color"",""slide"":1,""shape_id"":4,""row"":1,""col"":1,""paragraph_index"":0,""value"":""#FFFFFF""}"
        Case "set_cell_paragraph_bold", "set_cell_paragraph_italic"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, paragraph_index, value(bool)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":1,""shape_id"":4,""row"":1,""col"":1,""paragraph_index"":0,""value"":true}"
        Case "set_cell_paragraph_alignment"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, paragraph_index, value(""left""|""center""|""right""|""justify"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell_paragraph_alignment"",""slide"":1,""shape_id"":4,""row"":1,""col"":1,""paragraph_index"":0,""value"":""center""}"
        Case "set_cell_bullet_style"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, paragraph_index, value(""none""|""disc""|""square""|""dash""|""number""|""letter"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell_bullet_style"",""slide"":1,""shape_id"":4,""row"":1,""col"":1,""paragraph_index"":0,""value"":""disc""}"
        Case "add_cell_paragraph"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, after_paragraph_index(int; -1 prepends), value(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_cell_paragraph"",""slide"":1,""shape_id"":4,""row"":1,""col"":1,""after_paragraph_index"":0,""value"":""Sub-bullet""}"
        Case "delete_cell_paragraph"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, paragraph_index" & vbCrLf & _
                "  EXAMPLE:  {""type"":""delete_cell_paragraph"",""slide"":1,""shape_id"":4,""row"":1,""col"":1,""paragraph_index"":1}"
        Case "set_cell_indent_level"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, paragraph_index, value(int 0-4)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell_indent_level"",""slide"":1,""shape_id"":4,""row"":1,""col"":1,""paragraph_index"":1,""value"":1}" & vbCrLf & _
                "  NOTE: PowerPoint COM cannot set cell paragraph level via VBA. This action raises an " & _
                "error. Use python-pptx post-save: para._p.get_or_add_pPr().set('lvl', str(n))"
        Case "append_cell_text"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, value(string — appended after newline to existing cell text)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""append_cell_text"",""slide"":1,""shape_id"":4,""row"":1,""col"":1,""value"":""+12% YoY""}"
        Case "set_cell"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, row, col, AND at least one of: text, font_size, font_color, font_bold, font_italic, font_underline, font_name, fill, h_align, v_align" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_cell"",""slide"":1,""shape_id"":4,""row"":1,""col"":1,""text"":""Revenue"",""font_size"":12,""font_bold"":true,""fill"":""#15283C"",""font_color"":""#FFFFFF"",""h_align"":""center""}"
        ' ---- charts — granular ----
        Case "set_chart_type"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(chart type string — see add_chart vocabulary)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_chart_type"",""slide"":1,""shape_id"":2,""value"":""barclustered""}"
        Case "set_chart_title"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(string)" & vbCrLf & _
                "  OPTIONAL: enabled(bool)=true, props({font_size,font_color,font_bold,font_italic,position})" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_chart_title"",""slide"":1,""shape_id"":2,""value"":""FY25 Revenue"",""props"":{""font_size"":18,""font_bold"":true,""font_color"":""#15283C""}}"
        Case "set_chart_axis_title"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, axis(""x""|""y""|""category""|""value""), value(string)" & vbCrLf & _
                "  OPTIONAL: props({font_size,font_color,font_bold,font_italic})" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_chart_axis_title"",""slide"":1,""shape_id"":2,""axis"":""y"",""value"":""Revenue ($M)""}"
        Case "set_chart_legend_position"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(""top""|""right""|""bottom""|""left""|""none"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_chart_legend_position"",""slide"":1,""shape_id"":2,""value"":""bottom""}"
        Case "set_chart_legend"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, props(object: visible|position|font_size|font_color)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_chart_legend"",""slide"":1,""shape_id"":2,""props"":{""position"":""right"",""font_size"":10}}"
        Case "set_series_color"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, series_index(1-based), value(#RRGGBB)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_series_color"",""slide"":1,""shape_id"":2,""series_index"":1,""value"":""#15283C""}"
        Case "set_series_values"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, series_index(1-based), values(array of numbers; length must match categories)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_series_values"",""slide"":1,""shape_id"":2,""series_index"":1,""values"":[120,138,151,170]}"
        Case "set_chart_categories"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, categories(array of strings)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_chart_categories"",""slide"":1,""shape_id"":2,""categories"":[""FY22"",""FY23"",""FY24"",""FY25""]}"
        Case "set_series_name"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, series_index(1-based), value(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_series_name"",""slide"":1,""shape_id"":2,""series_index"":1,""value"":""Revenue""}"
        Case "set_chart_axis"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, axis(""x""|""y""|""y2""|""x2"" or aliases), props(object — see ACTIONS_REFERENCE.md §3.11)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_chart_axis"",""slide"":1,""shape_id"":2,""axis"":""y"",""props"":{""min"":0,""max"":200,""major_unit"":50,""number_format"":""$#,##0""}}"
        Case "set_chart_gridlines"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, props(object: major|minor|major_color|major_weight|major_dash|minor_color|minor_weight|minor_dash)" & vbCrLf & _
                "  OPTIONAL: axis(""x""|""y""|""category""|""value""|""both"")=""y""" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_chart_gridlines"",""slide"":1,""shape_id"":2,""props"":{""major"":true,""major_color"":""#E0E0E0"",""major_dash"":""dot""}}"
        Case "set_chart_format"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, props(object — see ACTIONS_REFERENCE.md §3.11 for full key list)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_chart_format"",""slide"":1,""shape_id"":2,""props"":{""gap_width"":50,""overlap"":0}}"
        Case "set_chart_series"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, series_index(1-based), props(object)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_chart_series"",""slide"":1,""shape_id"":2,""series_index"":1,""props"":{""fill"":""#15283C"",""show_labels"":true,""label_color"":""#FFFFFF"",""label_format"":""$#,##0""}}" & vbCrLf & _
                "  NOTE: data-label number_format / label_format / show_labels DO NOT WORK on the 7 modern chart types (waterfall, pareto, funnel, histogram, boxwhisker, treemap, sunburst). " & _
                "Those charts must be formatted by hand: right-click label > Format Data Labels > Number > Custom > enter format code."
        Case "add_chart_trendline"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, series_index(1-based), props(object: kind|order|period|forward|backward|display_equation|display_r_squared|color|dash|weight)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_chart_trendline"",""slide"":1,""shape_id"":2,""series_index"":1,""props"":{""kind"":""linear"",""color"":""#FF0000"",""display_r_squared"":true}}"
        Case "set_chart_error_bars"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, series_index(1-based), props({direction,include,type,amount,end_style})" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_chart_error_bars"",""slide"":1,""shape_id"":2,""series_index"":1,""props"":{""direction"":""y"",""include"":""both"",""type"":""percent"",""amount"":5}}"
        Case "set_chart_data_table"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, visible(bool)" & vbCrLf & _
                "  OPTIONAL: props({show_legend_key,horizontal_border,vertical_border,outline_border,font_size,font_color})" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_chart_data_table"",""slide"":1,""shape_id"":2,""visible"":true}"
        Case "set_line_smoothing"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, series_index(1-based), value(bool)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_line_smoothing"",""slide"":1,""shape_id"":2,""series_index"":1,""value"":true}"
        Case "delete_series"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, series_index(1-based)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""delete_series"",""slide"":1,""shape_id"":2,""series_index"":3}"
        Case "add_series"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, name(string), values(array of numbers; length must match categories)" & vbCrLf & _
                "  OPTIONAL: color(#RRGGBB)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_series"",""slide"":1,""shape_id"":2,""name"":""Forecast"",""values"":[180,195,210,225],""color"":""#A6A6A6""}"
        Case "set_data_label_text"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, series_index(1-based), point_index(1-based), value(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_data_label_text"",""slide"":1,""shape_id"":2,""series_index"":1,""point_index"":3,""value"":""peak""}"
        ' ---- images / web ----
        Case "replace_picture"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, path(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""replace_picture"",""slide"":1,""shape_id"":3,""path"":""C:\\imgs\\new.png""}"
        Case "fetch_page_images"
            GetActionGuidance = _
                "  REQUIRED: url(string)" & vbCrLf & _
                "  OPTIONAL: dest_folder(string), ref_name(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""fetch_page_images"",""url"":""https://example.com""}"
        Case "download_image"
            GetActionGuidance = _
                "  REQUIRED: url(string), dest_path(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""download_image"",""url"":""https://example.com/img.jpg"",""dest_path"":""C:\\imgs\\img.jpg""}"
        Case "open_image_picker"
            GetActionGuidance = _
                "  OPTIONAL: folder(string)  -- no required fields" & vbCrLf & _
                "  EXAMPLE:  {""type"":""open_image_picker"",""folder"":""C:\\imgs""}"
        Case "build_image_picker_slide"
            GetActionGuidance = _
                "  OPTIONAL: folder, cols(int)=4, insert_at(int)=0, max_per_slide(int)=24" & vbCrLf & _
                "  EXAMPLE:  {""type"":""build_image_picker_slide"",""cols"":4,""max_per_slide"":24}"
        Case "bulk_insert_image"
            GetActionGuidance = _
                "  REQUIRED: slide_indices(array of ints), picture_path, left, top, width, height (all num pt)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""bulk_insert_image"",""slide_indices"":[1,2,3],""picture_path"":""C:\\imgs\\logo.png"",""left"":800,""top"":510,""width"":120,""height"":40}"
        ' ---- slides ----
        Case "move_slide"
            GetActionGuidance = _
                "  REQUIRED: from (or from_slide), to (or to_slide) — both 1-based" & vbCrLf & _
                "  EXAMPLE:  {""type"":""move_slide"",""from"":3,""to"":1}"
        Case "extract_slides"
            GetActionGuidance = _
                "  REQUIRED: slide_indices(array of ints), output_path(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""extract_slides"",""slide_indices"":[1,3,5],""output_path"":""C:\\extracted.pptx""}"
        Case "import_slides_from_deck"
            GetActionGuidance = _
                "  REQUIRED: source_path(string), slide_indices(array of ints), target_position(int)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""import_slides_from_deck"",""source_path"":""C:\\other.pptx"",""slide_indices"":[2,3],""target_position"":1}"
        Case "apply_layout_to_slides"
            GetActionGuidance = _
                "  REQUIRED: slide_indices(array), layout_index(int, 0-based)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""apply_layout_to_slides"",""slide_indices"":[2,3,4],""layout_index"":1}"
        Case "change_slide_layout"
            GetActionGuidance = _
                "  REQUIRED: slide, layout_index(int, 0-based)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""change_slide_layout"",""slide"":3,""layout_index"":1}"
        Case "apply_theme"
            GetActionGuidance = _
                "  REQUIRED: theme_path(string — .thmx or .potx)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""apply_theme"",""theme_path"":""C:\\themes\\brand.thmx""}"
        Case "set_theme_font"
            GetActionGuidance = _
                "  REQUIRED: at least one of: major(string, heading), minor(string, body)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_theme_font"",""major"":""Calibri"",""minor"":""Calibri""}"
        Case "swap_font_deck_wide"
            GetActionGuidance = _
                "  REQUIRED: from_name(string, non-empty), to_name(string, non-empty)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""swap_font_deck_wide"",""from_name"":""Arial"",""to_name"":""Calibri""}"
        Case "set_slide_size"
            GetActionGuidance = _
                "  REQUIRED: preset(""16:9""|""4:3"") OR (width_pt(num>0) AND height_pt(num>0)) — not both" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_slide_size"",""preset"":""16:9""}"
        Case "bulk_insert_text_box"
            GetActionGuidance = _
                "  REQUIRED: slide_indices(array of ints), text(string), left, top, width, height (all num pt)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""bulk_insert_text_box"",""slide_indices"":[1,2,3],""text"":""CONFIDENTIAL"",""left"":800,""top"":510,""width"":120,""height"":20}"
        Case "set_slide_background_color"
            GetActionGuidance = _
                "  REQUIRED: slide, color(#RRGGBB)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_slide_background_color"",""slide"":1,""color"":""#15283C""}"
        Case "insert_slide_number"
            GetActionGuidance = _
                "  REQUIRED: slide, pos({left,top,width,height})" & vbCrLf & _
                "  OPTIONAL: ref_name, font_color, font_size" & vbCrLf & _
                "  EXAMPLE:  {""type"":""insert_slide_number"",""slide"":1,""pos"":{""left"":900,""top"":520,""width"":40,""height"":20},""font_size"":10}"
        Case "set_slide_hidden"
            GetActionGuidance = _
                "  REQUIRED: slide, value(bool — true hides, false un-hides from slideshow)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_slide_hidden"",""slide"":5,""value"":true}"
        Case "set_slide_name"
            GetActionGuidance = _
                "  REQUIRED: slide, value(string, non-empty)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_slide_name"",""slide"":1,""value"":""Title slide""}"
        Case "set_slide_transition"
            GetActionGuidance = _
                "  REQUIRED: slide, effect(""none""|""fade""|""push""|""wipe""|""split""|""reveal""|""cut""|""dissolve""|""checkerboard""|""blinds""|""random_bars""|""box""|""comb""|""zoom""|""morph"")" & vbCrLf & _
                "  OPTIONAL: speed(""slow""|""medium""|""fast""), advance_on_click(bool), advance_after_seconds(num)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_slide_transition"",""slide"":1,""effect"":""fade"",""speed"":""medium""}"
        Case "add_section"
            GetActionGuidance = _
                "  REQUIRED: before_slide(int, 1-based), name(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_section"",""before_slide"":3,""name"":""Financials""}"
        Case "delete_section"
            GetActionGuidance = _
                "  REQUIRED: section_index(int, 1-based)" & vbCrLf & _
                "  OPTIONAL: delete_slides(bool)=false" & vbCrLf & _
                "  EXAMPLE:  {""type"":""delete_section"",""section_index"":2,""delete_slides"":false}"
        Case "rename_section"
            GetActionGuidance = _
                "  REQUIRED: section_index(int), name(string)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""rename_section"",""section_index"":1,""name"":""Intro""}"
        Case "move_section"
            GetActionGuidance = _
                "  REQUIRED: section_index(int), to_position(int)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""move_section"",""section_index"":2,""to_position"":1}"
        ' ---- effects ----
        Case "set_line_color"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(#RRGGBB)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_line_color"",""slide"":1,""shape_id"":3,""value"":""#15283C""}"
        Case "set_line_weight"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, weight_pt(num>0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_line_weight"",""slide"":1,""shape_id"":3,""weight_pt"":1.5}"
        Case "set_line_style"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, style(""solid""|""dash""|""dot""|""dashdot"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_line_style"",""slide"":1,""shape_id"":3,""style"":""dash""}"
        Case "set_shadow"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, offset_x(num), offset_y(num), blur(num), color(#RRGGBB), transparency(num 0..1)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_shadow"",""slide"":1,""shape_id"":3,""offset_x"":3,""offset_y"":3,""blur"":6,""color"":""#000000"",""transparency"":0.5}"
        Case "set_glow"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, color(#RRGGBB), radius(num), transparency(num 0..1)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_glow"",""slide"":1,""shape_id"":3,""color"":""#FFD700"",""radius"":8,""transparency"":0.3}"
        Case "set_reflection"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, size(num 0..1), transparency(num 0..1), distance(num pt)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_reflection"",""slide"":1,""shape_id"":3,""size"":0.5,""transparency"":0.5,""distance"":4}"
        Case "set_transparency"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(num 0..1)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_transparency"",""slide"":1,""shape_id"":3,""value"":0.3}"
        Case "set_gradient_fill"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, color1(#RRGGBB), color2(#RRGGBB), angle(num deg)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_gradient_fill"",""slide"":1,""shape_id"":3,""color1"":""#15283C"",""color2"":""#2A4F82"",""angle"":90}"
        Case "set_3d_bevel"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, bevel_type(""circle""|""slope""|""cross""|""angle""|""softround""), depth_pt(num)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_3d_bevel"",""slide"":1,""shape_id"":3,""bevel_type"":""circle"",""depth_pt"":6}"
        Case "set_3d_rotation"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, AND at least one of: x(deg), y(deg), z(deg)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_3d_rotation"",""slide"":1,""shape_id"":3,""x"":20,""y"":-30}"
        Case "set_soft_edge"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, radius_pt(num; 0 clears)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_soft_edge"",""slide"":1,""shape_id"":3,""radius_pt"":5}"
        Case "apply_preset_effect"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, preset_index(int 1..24)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""apply_preset_effect"",""slide"":1,""shape_id"":3,""preset_index"":12}"
        Case "crop_picture"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, left(num), right(num), top(num), bottom(num) — all pt" & vbCrLf & _
                "  EXAMPLE:  {""type"":""crop_picture"",""slide"":1,""shape_id"":3,""left"":10,""right"":10,""top"":0,""bottom"":0}"
        Case "recolor_picture"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, color_type(""grayscale""|""sepia""|""washout""|""bw""|""auto"")" & vbCrLf & _
                "  EXAMPLE:  {""type"":""recolor_picture"",""slide"":1,""shape_id"":3,""color_type"":""grayscale""}"
        Case "set_brightness", "set_contrast"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, value(num -1.0..1.0; picture only)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":1,""shape_id"":3,""value"":0.2}"
        Case "apply_picture_artistic_effect"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, effect(""none""|""marker""|""pencil_grayscale""|""pencil_sketch""|""line_drawing""|""chalk_sketch""|""paint_strokes""|""paint_brush""|""glow_diffused""|""blur""|""light_screen""|""watercolor""|""film_grain""|""mosaic_bubbles""|""glass""|""cement""|""texturizer""|""crisscross""|""pastels_smooth""|""plastic_wrap""|""cutout""|""photocopy""|""glow_edges"")" & vbCrLf & _
                "  OPTIONAL: intensity(int 0..100)=50" & vbCrLf & _
                "  EXAMPLE:  {""type"":""apply_picture_artistic_effect"",""slide"":1,""shape_id"":3,""effect"":""watercolor"",""intensity"":50}"
        Case "reset_picture"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id  -- undoes brightness/contrast/crop/recolor/artistic effect" & vbCrLf & _
                "  EXAMPLE:  {""type"":""reset_picture"",""slide"":1,""shape_id"":3}"
        ' ---- speaker notes formatting ----
        Case "set_notes_font_size"
            GetActionGuidance = _
                "  REQUIRED: slide, value(int>0)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_notes_font_size"",""slide"":3,""value"":11}"
        Case "set_notes_font_color"
            GetActionGuidance = _
                "  REQUIRED: slide, value(#RRGGBB)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_notes_font_color"",""slide"":3,""value"":""#333333""}"
        Case "set_notes_font_bold", "set_notes_font_italic"
            GetActionGuidance = _
                "  REQUIRED: slide, value(bool)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""" & actionType & """,""slide"":3,""value"":true}"
        Case "set_notes_font_name"
            GetActionGuidance = _
                "  REQUIRED: slide, value(string, non-empty)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""set_notes_font_name"",""slide"":3,""value"":""Calibri""}"
        ' ---- verification ----
        Case "run_verification"
            GetActionGuidance = _
                "  OPTIONAL: scope(""deck""|""slide:N"")=""deck"", max_warnings(int)=100" & vbCrLf & _
                "  EXAMPLE:  {""type"":""run_verification"",""scope"":""deck""}"
        Case "apply_template"
            GetActionGuidance = _
                "  REQUIRED: template(""title""|""section""|""bullets""|""two_col""|""comparison""|""kpi_dashboard""|""quote""), content(object of the template's slots); OPTIONAL slide(int) targets an existing blank slide else a new slide is appended" & vbCrLf & _
                "  EXAMPLE:  {""type"":""apply_template"",""template"":""title"",""content"":{""title"":""Q3 Review"",""subtitle"":""FY26""}}"
        Case Else
            ' Fallback for any action whose guidance entry hasn't been added.
            ' This should be rare since the table above covers all ~165 known
            ' types. If it fires, add a Case for the new type.
            GetActionGuidance = _
                "  No canonical guidance entry for ""'" & actionType & "'"" yet. See " & _
                "docs/ACTIONS_REFERENCE.md §3 for the exact signature." & vbCrLf & _
                "  General rules: every action has type+slide; existing-shape actions add shape_id; " & vbCrLf & _
                "  paragraph actions add paragraph_index (0-based); run actions add paragraph_index + run_index (both 0-based); " & vbCrLf & _
                "  table cell actions use row+col (1-based); chart series actions use series_index (1-based)."
    End Select
End Function

Public Function SanitizeJsonInput(raw As String) As String
    Dim s As String: s = raw

    ' Strip UTF-8 BOM (EF BB BF) or UTF-16 BOM (FEFF).
    If Len(s) >= 3 Then
        If AscW(Mid(s, 1, 1)) = &HFEFF Then s = Mid(s, 2)
    End If

    ' Normalize smart quotes to ASCII. LLM autocorrectors sometimes wrap JSON
    ' keys/values in curly double quotes; converting those rescues the parse.
    ' BUT: only when the input has no ASCII double quotes at all -- otherwise the
    ' structure is already valid ASCII JSON and any curly double quotes are
    ' *content* inside string values (e.g. legal text full of "quoted" terms);
    ' converting them would inject unescaped quotes and break the parse. Curly
    ' single quotes are always safe to normalize inside JSON double-quoted strings.
    If InStr(s, """") = 0 Then
        s = Replace(s, ChrW(&H201C), """")  ' left double
        s = Replace(s, ChrW(&H201D), """")  ' right double
    End If
    s = Replace(s, ChrW(&H2018), "'")   ' left single
    s = Replace(s, ChrW(&H2019), "'")   ' right single

    ' Strip Markdown code fences (e.g. ```json ... ``` or ``` ... ```).
    s = ReplaceCaseInsensitive(s, "```json", "")
    s = Replace(s, "```", "")

    ' Extract the JSON region. Prose, rejected drafts, or worked examples
    ' around the real payload can carry their own { } / [ ] (e.g. an LLM
    ' showing a draft before the final answer, or "replace {placeholder}"
    ' after it), so a naive first-{ ... last-} span splices unrelated text
    ' together and breaks the parse. Instead collect every top-level
    ' balanced bracket region (string- and comment-aware) and keep the
    ' longest -- the real batch is essentially always the largest region.
    Dim span As String: span = ExtractJsonSpan(s)
    If Len(span) = 0 Then
        SanitizeJsonInput = raw
        Exit Function
    End If
    s = span

    ' Strip JS-style comments. Preserve // and /* sequences when they appear
    ' inside string literals; only strip outside strings.
    s = StripJsonComments(s)

    ' Strip trailing commas before } or ]. Walk char by char, string-aware.
    s = StripTrailingCommas(s)

    SanitizeJsonInput = s
End Function

' Return the longest top-level balanced { } or [ ] region in s, scanning
' string-aware and comment-aware so brackets inside JSON strings or JS
' comments never affect nesting depth. Returns "" if none found (which
' preserves the documented "no { or [" -> return original contract).
Private Function ExtractJsonSpan(s As String) As String
    Dim n As Long: n = Len(s)
    Dim i As Long: i = 1
    Dim best As String: best = ""
    Dim ch As String
    Do While i <= n
        ch = Mid(s, i, 1)
        If ch = "{" Or ch = "[" Then
            Dim endPos As Long: endPos = ScanBalancedEnd(s, i)
            Dim cand As String
            If endPos > 0 Then
                cand = Mid(s, i, endPos - i + 1)
                If Len(cand) > Len(best) Then best = cand
                i = endPos + 1
            Else
                ' Never rebalances from here: the remainder is the only
                ' candidate at this position and nothing later can be longer.
                cand = Mid(s, i)
                If Len(cand) > Len(best) Then best = cand
                Exit Do
            End If
        Else
            i = i + 1
        End If
    Loop
    ExtractJsonSpan = best
End Function

' Starting at an opening { or [ at startPos, return the position of its
' matching close, or 0 if it never rebalances. String-aware (double-quoted,
' backslash escapes) and JS-comment-aware (// line, /* block */).
Private Function ScanBalancedEnd(s As String, ByVal startPos As Long) As Long
    Dim n As Long: n = Len(s)
    Dim depth As Long: depth = 0
    Dim inString As Boolean: inString = False
    Dim i As Long: i = startPos
    Dim ch As String, ch2 As String
    Do While i <= n
        ch = Mid(s, i, 1)
        If inString Then
            If ch = "\" And i < n Then
                i = i + 2
            ElseIf ch = """" Then
                inString = False
                i = i + 1
            Else
                i = i + 1
            End If
        Else
            If ch = """" Then
                inString = True
                i = i + 1
            ElseIf ch = "/" And i < n Then
                ch2 = Mid(s, i + 1, 1)
                If ch2 = "/" Then
                    Dim j As Long: j = i + 2
                    Do While j <= n
                        If Mid(s, j, 1) = vbLf Or Mid(s, j, 1) = vbCr Then Exit Do
                        j = j + 1
                    Loop
                    i = j
                ElseIf ch2 = "*" Then
                    Dim k As Long: k = i + 2
                    Do While k < n
                        If Mid(s, k, 1) = "*" And Mid(s, k + 1, 1) = "/" Then
                            k = k + 2
                            Exit Do
                        End If
                        k = k + 1
                    Loop
                    i = k
                Else
                    i = i + 1
                End If
            ElseIf ch = "{" Or ch = "[" Then
                depth = depth + 1
                i = i + 1
            ElseIf ch = "}" Or ch = "]" Then
                depth = depth - 1
                If depth = 0 Then
                    ScanBalancedEnd = i
                    Exit Function
                End If
                i = i + 1
            Else
                i = i + 1
            End If
        End If
    Loop
    ScanBalancedEnd = 0
End Function

' Walk JSON string-aware, removing commas that appear immediately before
' (ignoring whitespace) a closing } or ]. Trailing commas are valid in
' JavaScript but NOT in strict JSON, and LLMs frequently emit them.
Private Function StripTrailingCommas(s As String) As String
    Dim n As Long: n = Len(s)
    Dim out As String
    Dim i As Long: i = 1
    Dim inString As Boolean: inString = False
    Dim ch As String
    Do While i <= n
        ch = Mid(s, i, 1)
        If inString Then
            If ch = "\" And i < n Then
                out = out & ch & Mid(s, i + 1, 1)
                i = i + 2
            ElseIf ch = """" Then
                inString = False
                out = out & ch
                i = i + 1
            Else
                out = out & ch
                i = i + 1
            End If
        Else
            If ch = """" Then
                inString = True
                out = out & ch
                i = i + 1
            ElseIf ch = "," Then
                ' Look ahead past whitespace to see if next non-ws char is } or ].
                Dim j As Long: j = i + 1
                Do While j <= n
                    Dim nx As String: nx = Mid(s, j, 1)
                    If nx = " " Or nx = vbTab Or nx = vbCr Or nx = vbLf Then
                        j = j + 1
                    Else
                        Exit Do
                    End If
                Loop
                If j <= n And (Mid(s, j, 1) = "}" Or Mid(s, j, 1) = "]") Then
                    ' Drop the comma; preserve the whitespace that followed.
                    i = i + 1
                Else
                    out = out & ch
                    i = i + 1
                End If
            Else
                out = out & ch
                i = i + 1
            End If
        End If
    Loop
    StripTrailingCommas = out
End Function

Private Function ReplaceCaseInsensitive(haystack As String, needle As String, repl As String) As String
    Dim out As String: out = haystack
    Dim pos As Long: pos = InStr(1, out, needle, vbTextCompare)
    Do While pos > 0
        out = Left(out, pos - 1) & repl & Mid(out, pos + Len(needle))
        pos = InStr(pos + Len(repl), out, needle, vbTextCompare)
    Loop
    ReplaceCaseInsensitive = out
End Function

' Walk the string char by char, tracking string state, and skip JS comments
' that appear OUTSIDE string literals. JSON strings are double-quoted with
' backslash escapes.
Private Function StripJsonComments(s As String) As String
    Dim n As Long: n = Len(s)
    Dim out As String
    Dim i As Long: i = 1
    Dim inString As Boolean: inString = False
    Dim ch As String, ch2 As String
    Do While i <= n
        ch = Mid(s, i, 1)
        If inString Then
            If ch = "\" And i < n Then
                out = out & ch & Mid(s, i + 1, 1)
                i = i + 2
            ElseIf ch = """" Then
                inString = False
                out = out & ch
                i = i + 1
            Else
                out = out & ch
                i = i + 1
            End If
        Else
            If ch = """" Then
                inString = True
                out = out & ch
                i = i + 1
            ElseIf ch = "/" And i < n Then
                ch2 = Mid(s, i + 1, 1)
                If ch2 = "/" Then
                    ' Line comment: skip until newline (keep newline)
                    Dim j As Long: j = i + 2
                    Do While j <= n
                        If Mid(s, j, 1) = vbLf Or Mid(s, j, 1) = vbCr Then Exit Do
                        j = j + 1
                    Loop
                    i = j
                ElseIf ch2 = "*" Then
                    ' Block comment: skip until */
                    Dim k As Long: k = i + 2
                    Do While k < n
                        If Mid(s, k, 1) = "*" And Mid(s, k + 1, 1) = "/" Then
                            k = k + 2
                            Exit Do
                        End If
                        k = k + 1
                    Loop
                    i = k
                Else
                    out = out & ch
                    i = i + 1
                End If
            Else
                out = out & ch
                i = i + 1
            End If
        End If
    Loop
    StripJsonComments = out
End Function

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

' Group actions by (slide, shape, paragraph_index) when they target a run.
' Within each group, process descending run_index so a set_run_text on run 0
' doesn't shift the indices of runs 1, 2, ... that come later in the batch.
' Non-run actions retain their relative order; run-action groups land at the
' position of their first member in the original list.
Private Function ReorderForRunIndexSafety(src As Object) As Collection
    Dim out As New Collection
    Dim seenGroups As Object: Set seenGroups = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To src.Count
        Dim a As Object: Set a = src(i)
        If IsRunAction(GetStr(a, "type")) Then
            Dim key As String
            key = GetStr(a, "slide") & "|" & GetStr(a, "shape_id") & "|" & GetStr(a, "paragraph_index")
            If Not seenGroups.Exists(key) Then
                ' Collect all run actions in src that share the key, sort by run_index DESC, append to out.
                Dim group As New Collection
                Dim j As Long
                For j = i To src.Count
                    Dim aj As Object: Set aj = src(j)
                    If IsRunAction(GetStr(aj, "type")) _
                       And GetStr(aj, "slide") = GetStr(a, "slide") _
                       And GetStr(aj, "shape_id") = GetStr(a, "shape_id") _
                       And GetStr(aj, "paragraph_index") = GetStr(a, "paragraph_index") Then
                        InsertSortedDesc group, aj
                    End If
                Next j
                Dim k As Long
                For k = 1 To group.Count
                    out.Add group(k)
                Next k
                seenGroups.Add key, True
            End If
            ' Skip — group emitted already
        Else
            out.Add a
        End If
    Next i
    Set ReorderForRunIndexSafety = out
End Function

Private Function IsRunAction(t As String) As Boolean
    Select Case t
        Case "set_run_bold", "set_run_italic", "set_run_underline", _
             "set_run_subscript", "set_run_superscript", _
             "set_run_font_color", "set_run_font_size", "set_run_font_name", _
             "set_run_text", "set_run_hyperlink"
            IsRunAction = True
        Case Else
            IsRunAction = False
    End Select
End Function

Private Sub InsertSortedDesc(c As Collection, item As Object)
    Dim ri As Long: ri = CLng(item("run_index"))
    Dim i As Long
    For i = 1 To c.Count
        If CLng(c(i)("run_index")) < ri Then
            c.Add item, , i
            Exit Sub
        End If
    Next i
    c.Add item
End Sub

