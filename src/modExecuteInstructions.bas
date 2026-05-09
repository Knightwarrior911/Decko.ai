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

    ExecuteFromString = applied & " applied, " & skipped & " skipped. " & _
                        "Log: " & deckPath & ".action_log.jsonl"
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
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "type", "depth_pt"))
            If ValidateAction = "" Then
                Dim bv As String: bv = LCase(CStr(act("type")))
                If bv <> "circle" And bv <> "slope" And bv <> "cross" And bv <> "angle" And bv <> "softround" Then _
                    ValidateAction = "type: must be circle/slope/cross/angle/softround"
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
        Case Else
            ValidateAction = "unknown_type: " & t
    End Select
End Function

Private Function RequireFields(act As Object, fields As Variant) As String
    Dim i As Long
    For i = LBound(fields) To UBound(fields)
        Dim f As String: f = CStr(fields(i))
        If f = "shape_id" Then
            ' accept shape_id OR shape_name
            If Not act.Exists("shape_id") And Not act.Exists("shape_name") Then
                RequireFields = "missing_field: shape_id or shape_name"
                Exit Function
            End If
        ElseIf Not act.Exists(f) Then
            RequireFields = "missing_field: " & f
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
    ' Resolve shape_name → inject shape_id so all dispatch code works unchanged
    If Not act.Exists("shape_id") Then
        If act.Exists("shape_name") Then
            Dim sh As Shape
            Set sh = modActions.FindShapeByName(CLng(act("slide")), CStr(act("shape_name")))
            If sh Is Nothing Then
                ValidateShape = "shape_name '" & CStr(act("shape_name")) & "': not found"
                Exit Function
            End If
            act.Add "shape_id", CLng(sh.Id)
        Else
            ValidateShape = "shape_id or shape_name: required"
            Exit Function
        End If
    End If
    Dim shCheck As Shape
    Set shCheck = modActions.FindShape(CLng(act("slide")), CLng(act("shape_id")))
    If shCheck Is Nothing Then ValidateShape = "shape_not_found"
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
            If act.Exists("font_bold") Then asBold = CBool(act("font_bold"))
            If act.Exists("h_align") Then asAlign = CStr(act("h_align"))
            If act.Exists("v_align") Then asVAlign = CStr(act("v_align"))
            modActionsLayout.Do_add_shape CLng(act("slide")), CStr(act("kind")), _
                                          CSng(posDict("left")), CSng(posDict("top")), _
                                          CSng(posDict("width")), CSng(posDict("height")), _
                                          fh, shex, swt, asRef, asTxt, asFc, asFs, asBold, asAlign, asVAlign
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
            modActionsConnector.Do_add_connector CLng(act("slide")), _
                                                 CLng(act("from_shape_id")), _
                                                 CLng(act("to_shape_id")), _
                                                 CStr(act("kind")), ae, cc, cw, astart, asize, fp, tp, ds
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
        Case "set_run_bold"
            modActionsRun.Do_set_run_bold CLng(act("slide")), CLng(act("shape_id")), _
                                          CLng(act("paragraph_index")), CLng(act("run_index")), _
                                          CBool(act("value"))
        Case "set_run_italic"
            modActionsRun.Do_set_run_italic CLng(act("slide")), CLng(act("shape_id")), _
                                            CLng(act("paragraph_index")), CLng(act("run_index")), _
                                            CBool(act("value"))
        Case "set_run_underline"
            modActionsRun.Do_set_run_underline CLng(act("slide")), CLng(act("shape_id")), _
                                               CLng(act("paragraph_index")), CLng(act("run_index")), _
                                               CBool(act("value"))
        Case "set_run_subscript"
            modActionsRun.Do_set_run_subscript CLng(act("slide")), CLng(act("shape_id")), _
                                               CLng(act("paragraph_index")), CLng(act("run_index")), _
                                               CBool(act("value"))
        Case "set_run_superscript"
            modActionsRun.Do_set_run_superscript CLng(act("slide")), CLng(act("shape_id")), _
                                                 CLng(act("paragraph_index")), CLng(act("run_index")), _
                                                 CBool(act("value"))
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
            modActionsLayout.Do_match_size CLng(act("slide")), CLng(act("ref_shape_id")), _
                                           act("target_shape_ids")
        Case "uniform_size"
            modActionsLayout.Do_uniform_size CLng(act("slide")), act("shape_ids"), _
                                             CDbl(act("width_pt")), CDbl(act("height_pt"))
        Case "smart_spacing"
            modActionsLayout.Do_smart_spacing CLng(act("slide")), act("shape_ids"), _
                                              CDbl(act("gap_pt")), CStr(act("axis"))
        Case "equalize_spacing"
            modActionsLayout.Do_equalize_spacing CLng(act("slide")), act("shape_ids"), _
                                                 CStr(act("axis"))
        Case "match_position"
            modActionsLayout.Do_match_position CLng(act("slide")), CLng(act("ref_shape_id")), _
                                               CLng(act("target_shape_id")), CStr(act("edge"))
        Case "swap_positions"
            modActionsLayout.Do_swap_positions CLng(act("slide")), CLng(act("shape_a_id")), _
                                               CLng(act("shape_b_id"))
        Case "group_by_overlap"
            modActionsLayout.Do_group_by_overlap CLng(act("slide")), act("shape_ids")
        Case "find_replace_regex"
            modActionsDeck.Do_find_replace_regex CStr(act("scope")), CStr(act("pattern")), CStr(act("replacement"))
        Case "swap_font_deck_wide"
            modActionsDeck.Do_swap_font_deck_wide CStr(act("from_name")), CStr(act("to_name"))
        Case "recolor_palette_deck_wide"
            modActionsDeck.Do_recolor_palette_deck_wide CStr(act("from_hex")), CStr(act("to_hex")), CStr(act("target"))
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
                CStr(act("type")), CDbl(act("depth_pt"))
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
            If act.Exists("font_bold") Then tbBold = CBool(act("font_bold"))
            If act.Exists("font_italic") Then tbItalic = CBool(act("font_italic"))
            If act.Exists("h_align") Then tbAlign = CStr(act("h_align"))
            If act.Exists("fill") Then If Not IsNull(act("fill")) Then tbFill = CStr(act("fill"))
            If act.Exists("stroke") Then If Not IsNull(act("stroke")) Then tbStroke = CStr(act("stroke"))
            If act.Exists("stroke_weight_pt") Then tbSw = CSng(act("stroke_weight_pt"))
            modActionsLayout.Do_add_text_box CLng(act("slide")), CStr(act("text")), _
                CSng(tbPos("left")), CSng(tbPos("top")), CSng(tbPos("width")), CSng(tbPos("height")), _
                tbRef, tbFc, tbFs, tbBold, tbItalic, tbAlign, tbFill, tbStroke, tbSw
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
            If act.Exists("color") Then cbColor = CStr(act("color"))
            If act.Exists("weight_pt") Then cbWeight = CSng(act("weight_pt"))
            If act.Exists("visible") Then cbVisible = CBool(act("visible"))
            modActionsTable.Do_set_cell_border CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("row")), CLng(act("col")), CStr(act("side")), cbColor, cbWeight, cbVisible
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
                CLng(act("source_shape_id")), CLng(act("target_shape_id"))
        Case "set_run_strikethrough"
            modActionsRun.Do_set_run_strikethrough CLng(act("slide")), CLng(act("shape_id")), _
                CLng(act("paragraph_index")), CLng(act("run_index")), CBool(act("value"))
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
Public Function SanitizeJsonInput(raw As String) As String
    Dim s As String: s = raw

    ' Strip UTF-8 BOM (EF BB BF) or UTF-16 BOM (FEFF).
    If Len(s) >= 3 Then
        If AscW(Mid(s, 1, 1)) = &HFEFF Then s = Mid(s, 2)
    End If

    ' Normalize smart quotes to ASCII. LLM autocorrectors sometimes inject these.
    ' U+201C and U+201D are curly double quotes; U+2018 and U+2019 are curly singles.
    s = Replace(s, ChrW(&H201C), """")  ' left double
    s = Replace(s, ChrW(&H201D), """")  ' right double
    s = Replace(s, ChrW(&H2018), "'")   ' left single
    s = Replace(s, ChrW(&H2019), "'")   ' right single

    ' Strip Markdown code fences (e.g. ```json ... ``` or ``` ... ```).
    s = ReplaceCaseInsensitive(s, "```json", "")
    s = Replace(s, "```", "")

    ' Find first { or [ - that's where JSON starts.
    Dim openPos As Long
    Dim posBrace As Long: posBrace = InStr(s, "{")
    Dim posBracket As Long: posBracket = InStr(s, "[")
    If posBrace = 0 And posBracket = 0 Then
        SanitizeJsonInput = raw
        Exit Function
    End If
    If posBrace = 0 Then
        openPos = posBracket
    ElseIf posBracket = 0 Then
        openPos = posBrace
    Else
        openPos = IIf(posBrace < posBracket, posBrace, posBracket)
    End If
    s = Mid(s, openPos)

    ' Find last } or ] - that's where JSON ends.
    Dim closePos As Long
    Dim lastBrace As Long: lastBrace = InStrRev(s, "}")
    Dim lastBracket As Long: lastBracket = InStrRev(s, "]")
    closePos = IIf(lastBrace > lastBracket, lastBrace, lastBracket)
    If closePos > 0 Then s = Left(s, closePos)

    ' Strip JS-style comments. Preserve // and /* sequences when they appear
    ' inside string literals; only strip outside strings.
    s = StripJsonComments(s)

    ' Strip trailing commas before } or ]. Walk char by char, string-aware.
    s = StripTrailingCommas(s)

    SanitizeJsonInput = s
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
