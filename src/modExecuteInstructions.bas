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
