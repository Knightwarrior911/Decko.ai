Attribute VB_Name = "modActions"
Option Explicit

' --- Lookup helpers --------------------------------------------------------

Public Function FindShape(slideNum As Long, shapeId As Long) As Shape
    Set FindShape = Nothing
    Dim pres As Presentation
    Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then Exit Function

    Dim sl As Slide
    Set sl = pres.Slides(slideNum)
    Dim sh As Shape
    For Each sh In sl.Shapes
        If sh.Id = shapeId Then
            Set FindShape = sh
            Exit Function
        End If
    Next sh
End Function

Public Function FindShapeByName(slideNum As Long, refName As String) As Shape
    Set FindShapeByName = Nothing
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then Exit Function
    Dim sh As Shape
    For Each sh In pres.Slides(slideNum).Shapes
        If sh.Name = refName Then
            Set FindShapeByName = sh
            Exit Function
        End If
    Next sh
End Function

' Universal shape reference resolver. Accepts numeric Id (Long/string-of-digits)
' or string ref_name. Returns the numeric Id or raises a descriptive error.
' Use this in the dispatcher for any shape-id field (from_shape_id, ref_shape_id,
' source_shape_id, target_shape_id, shape_a_id, shape_b_id, etc.) where the LLM
' might emit either an integer or the ref_name from the original add_shape call.
Public Function ResolveShapeRef(ByVal slideNum As Long, ByVal raw As Variant, ByVal label As String) As Long
    If IsNumeric(raw) Then
        ResolveShapeRef = CLng(raw)
        Exit Function
    End If
    Dim asStr As String: asStr = CStr(raw)
    Dim sh As Shape: Set sh = FindShapeByName(slideNum, asStr)
    If sh Is Nothing Then
        Err.Raise vbObjectError + 2050, "ResolveShapeRef", _
                  label & " not found on slide " & slideNum & ": " & asStr
    End If
    ResolveShapeRef = sh.Id
End Function

' Debug helper: round-trip a JSON array through ResolveShapeRef per-element.
' Called from Python tests to localize where Variant-binding fails.
Public Function DebugResolveArray(slideNum As Long, jsonArrayText As String) As String
    Dim parsed As Object
    Set parsed = modJSON.ParseJson("{""a"":" & jsonArrayText & "}")
    Dim col As Object: Set col = parsed("a")
    Dim out As String: out = ""
    Dim i As Long
    For i = 1 To col.Count
        On Error Resume Next
        Dim v As Variant: v = col(i)
        Dim t As String: t = TypeName(v)
        Dim id As Long: id = ResolveShapeRef(slideNum, v, "test[" & (i - 1) & "]")
        If Err.Number <> 0 Then
            out = out & "[" & i & "] " & t & " ERR=" & Err.Description & "; "
            Err.Clear
        Else
            out = out & "[" & i & "] " & t & " ok=" & id & "; "
        End If
        On Error GoTo 0
    Next i
    DebugResolveArray = out
End Function

' --- Text/format actions ---------------------------------------------------

Public Sub Do_set_text(slideNum As Long, shapeId As Long, value As String)
    Dim sh As Shape
    Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2001, "Do_set_text", "shape not found"
    If Not sh.HasTextFrame Then Err.Raise vbObjectError + 2002, "Do_set_text", "shape has no text frame"
    sh.TextFrame.TextRange.Text = value
End Sub

Public Sub Do_set_font_size(slideNum As Long, shapeId As Long, value As Long)
    Dim sh As Shape: Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2001, "Do_set_font_size", "shape not found"
    If Not sh.HasTextFrame Then Err.Raise vbObjectError + 2002, "Do_set_font_size", "no text frame"
    sh.TextFrame.TextRange.Font.Size = value
End Sub

Public Sub Do_set_font_bold(slideNum As Long, shapeId As Long, value As Boolean)
    Dim sh As Shape: Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2001, "Do_set_font_bold", "shape not found"
    If Not sh.HasTextFrame Then Err.Raise vbObjectError + 2002, "Do_set_font_bold", "no text frame"
    sh.TextFrame.TextRange.Font.Bold = IIf(value, msoTrue, msoFalse)
End Sub

Public Sub Do_set_font_italic(slideNum As Long, shapeId As Long, value As Boolean)
    Dim sh As Shape: Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2001, "Do_set_font_italic", "shape not found"
    If Not sh.HasTextFrame Then Err.Raise vbObjectError + 2002, "Do_set_font_italic", "no text frame"
    sh.TextFrame.TextRange.Font.Italic = IIf(value, msoTrue, msoFalse)
End Sub

Public Sub Do_set_font_color(slideNum As Long, shapeId As Long, hexValue As String)
    Dim sh As Shape: Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2001, "Do_set_font_color", "shape not found"
    If Not sh.HasTextFrame Then Err.Raise vbObjectError + 2002, "Do_set_font_color", "no text frame"
    sh.TextFrame.TextRange.Font.Color.RGB = HexToRgb(hexValue)
End Sub

Public Function HexToRgb(ByVal hexValue As String) As Long
    Dim h As String: h = hexValue
    If Left(h, 1) = "#" Then h = Mid(h, 2)
    If Len(h) <> 6 Then Err.Raise vbObjectError + 2003, "HexToRgb", "expected #RRGGBB, got: " & hexValue
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Mid(h, 1, 2))
    g = CLng("&H" & Mid(h, 3, 2))
    b = CLng("&H" & Mid(h, 5, 2))
    HexToRgb = RGB(r, g, b)
End Function

' Universal color resolver. Accepts:
'   "#RRGGBB" / "RRGGBB" / "rgb(r,g,b)" / array [r,g,b] / Collection(r,g,b) / numeric Long.
Public Function ResolveColor(ByVal v As Variant) As Long
    If TypeName(v) = "Collection" Then
        Dim col As Object: Set col = v
        If col.Count <> 3 Then Err.Raise vbObjectError + 2004, "ResolveColor", "rgb collection needs 3 elements"
        ResolveColor = RGB(CLng(col(1)), CLng(col(2)), CLng(col(3)))
        Exit Function
    ElseIf IsArray(v) Then
        Dim lo As Long, hi As Long
        lo = LBound(v): hi = UBound(v)
        If hi - lo + 1 <> 3 Then Err.Raise vbObjectError + 2004, "ResolveColor", "rgb array needs 3 elements"
        ResolveColor = RGB(CLng(v(lo)), CLng(v(lo + 1)), CLng(v(lo + 2)))
        Exit Function
    ElseIf IsNumeric(v) Then
        ResolveColor = CLng(v)
        Exit Function
    End If
    Dim s As String: s = Trim(CStr(v))
    If LCase(Left(s, 4)) = "rgb(" And Right(s, 1) = ")" Then
        Dim inner As String: inner = Mid(s, 5, Len(s) - 5)
        Dim parts() As String: parts = Split(inner, ",")
        If UBound(parts) - LBound(parts) + 1 <> 3 Then Err.Raise vbObjectError + 2004, "ResolveColor", "rgb() needs 3 args"
        ResolveColor = RGB(CLng(Trim(parts(0))), CLng(Trim(parts(1))), CLng(Trim(parts(2))))
        Exit Function
    End If
    ResolveColor = HexToRgb(s)
End Function

' Locale-tolerant boolean coercion. Accepts True/False, 1/0, "true"/"false"/"yes"/"no"/"y"/"n"/"on"/"off".
Public Function ToBool(ByVal v As Variant) As Boolean
    If IsNull(v) Or IsEmpty(v) Then ToBool = False: Exit Function
    If TypeName(v) = "Boolean" Then ToBool = CBool(v): Exit Function
    If IsNumeric(v) Then ToBool = (CDbl(v) <> 0): Exit Function
    Dim s As String: s = LCase(Trim(CStr(v)))
    Select Case s
        Case "true", "yes", "y", "on", "1":  ToBool = True
        Case "false", "no", "n", "off", "0", "": ToBool = False
        Case Else: Err.Raise vbObjectError + 2005, "ToBool", "cannot coerce to bool: " & CStr(v)
    End Select
End Function

' Tolerant integer coercion. Accepts numeric, "42", "42.0", "42pt".
Public Function ToLong(ByVal v As Variant) As Long
    If IsNumeric(v) Then ToLong = CLng(CDbl(v)): Exit Function
    Dim s As String: s = Trim(CStr(v))
    ' Strip common unit suffixes (pt, px, em)
    Dim units As Variant: units = Array("pt", "px", "em", "rem", "%")
    Dim u As Variant
    For Each u In units
        If Len(s) > Len(u) And LCase(Right(s, Len(u))) = u Then s = Trim(Left(s, Len(s) - Len(u)))
    Next u
    If IsNumeric(s) Then ToLong = CLng(CDbl(s)): Exit Function
    Err.Raise vbObjectError + 2006, "ToLong", "cannot coerce to long: " & CStr(v)
End Function

' Tolerant single-precision float coercion (same suffix-stripping).
Public Function ToSng(ByVal v As Variant) As Single
    If IsNumeric(v) Then ToSng = CSng(CDbl(v)): Exit Function
    Dim s As String: s = Trim(CStr(v))
    Dim units As Variant: units = Array("pt", "px", "em", "rem", "%")
    Dim u As Variant
    For Each u In units
        If Len(s) > Len(u) And LCase(Right(s, Len(u))) = u Then s = Trim(Left(s, Len(s) - Len(u)))
    Next u
    If IsNumeric(s) Then ToSng = CSng(CDbl(s)): Exit Function
    Err.Raise vbObjectError + 2007, "ToSng", "cannot coerce to single: " & CStr(v)
End Function

' --- Fill color action -------------------------------------------------------

Public Sub Do_set_fill_color(slideNum As Long, shapeId As Long, hexValue As String)
    Dim sh As Shape: Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2001, "Do_set_fill_color", "shape not found"
    sh.Fill.Visible = msoTrue
    sh.Fill.Solid
    sh.Fill.ForeColor.RGB = HexToRgb(hexValue)
End Sub

' --- Geometry actions --------------------------------------------------------

Public Sub Do_move_shape(slideNum As Long, shapeId As Long, leftPt As Single, topPt As Single)
    Dim sh As Shape: Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2001, "Do_move_shape", "shape not found"
    sh.Left = leftPt
    sh.Top = topPt
End Sub

Public Sub Do_resize_shape(slideNum As Long, shapeId As Long, widthPt As Single, heightPt As Single)
    Dim sh As Shape: Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2001, "Do_resize_shape", "shape not found"
    sh.LockAspectRatio = msoFalse
    sh.Width = widthPt
    sh.Height = heightPt
End Sub

Public Sub Do_delete_shape(slideNum As Long, shapeId As Long)
    Dim sh As Shape: Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2001, "Do_delete_shape", "shape not found"
    sh.Delete
End Sub

' --- Slide ops ---------------------------------------------------------------

Public Sub Do_add_slide(position As Long, layoutIndex As Long)
    Dim pres As Presentation: Set pres = ActivePresentation
    If position < 1 Or position > pres.Slides.Count + 1 Then position = pres.Slides.Count + 1
    Dim layout As CustomLayout
    Set layout = pres.SlideMaster.CustomLayouts(layoutIndex + 1)  ' 1-based in VBA
    pres.Slides.AddSlide position, layout
End Sub

Public Sub Do_delete_slide(slideNum As Long)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 2001, "Do_delete_slide", "slide out of range"
    End If
    pres.Slides(slideNum).Delete
End Sub

Public Sub Do_duplicate_slide(slideNum As Long)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 2001, "Do_duplicate_slide", "slide out of range"
    End If
    pres.Slides(slideNum).Duplicate
End Sub

' --- Table ops ---------------------------------------------------------------

Public Sub Do_set_cell_text(slideNum As Long, shapeId As Long, _
                            rowNum As Long, colNum As Long, value As String)
    Dim sh As Shape: Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2001, "Do_set_cell_text", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 2004, "Do_set_cell_text", "shape is not a table"
    sh.Table.Cell(rowNum, colNum).Shape.TextFrame.TextRange.Text = value
End Sub

Public Sub Do_swap_table_columns(slideNum As Long, shapeId As Long, _
                                 colA As Long, colB As Long)
    Dim sh As Shape: Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2001, "Do_swap_table_columns", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 2004, "Do_swap_table_columns", "shape is not a table"

    Dim tbl As Table: Set tbl = sh.Table
    Dim rowIdx As Long
    For rowIdx = 1 To tbl.Rows.Count
        SwapCellContents tbl.Cell(rowIdx, colA), tbl.Cell(rowIdx, colB)
    Next rowIdx
End Sub

Public Sub Do_swap_table_rows(slideNum As Long, shapeId As Long, _
                              rowA As Long, rowB As Long)
    Dim sh As Shape: Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2001, "Do_swap_table_rows", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 2004, "Do_swap_table_rows", "shape is not a table"

    Dim tbl As Table: Set tbl = sh.Table
    Dim colIdx As Long
    For colIdx = 1 To tbl.Columns.Count
        SwapCellContents tbl.Cell(rowA, colIdx), tbl.Cell(rowB, colIdx)
    Next colIdx
End Sub

Private Sub SwapCellContents(cellA As Object, cellB As Object)
    ' Swap text only. V1 keeps it simple.
    Dim aText As String, bText As String
    aText = SafeCellText(cellA)
    bText = SafeCellText(cellB)
    SetCellText cellA, bText
    SetCellText cellB, aText
End Sub

Private Function SafeCellText(c As Object) As String
    On Error Resume Next
    SafeCellText = c.Shape.TextFrame.TextRange.Text
End Function

Private Sub SetCellText(c As Object, t As String)
    On Error Resume Next
    c.Shape.TextFrame.TextRange.Text = t
End Sub

Public Sub Do_set_speaker_notes(slideNum As Long, value As String)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 5001, "Do_set_speaker_notes", "slide_out_of_range"
    End If
    Dim sl As Slide: Set sl = pres.Slides(slideNum)
    Dim ph As Object
    Dim i As Long
    For i = 1 To sl.NotesPage.Shapes.Placeholders.Count
        Set ph = sl.NotesPage.Shapes.Placeholders(i)
        If ph.PlaceholderFormat.Type = ppPlaceholderBody Then
            ph.TextFrame.TextRange.Text = value
            Exit Sub
        End If
    Next i
End Sub

Public Sub Do_append_speaker_notes(slideNum As Long, value As String)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 5001, "Do_append_speaker_notes", "slide_out_of_range"
    End If
    Dim sl As Slide: Set sl = pres.Slides(slideNum)
    Dim ph As Object
    Dim i As Long
    For i = 1 To sl.NotesPage.Shapes.Placeholders.Count
        Set ph = sl.NotesPage.Shapes.Placeholders(i)
        If ph.PlaceholderFormat.Type = ppPlaceholderBody Then
            Dim cur As String: cur = ph.TextFrame.TextRange.Text
            If Len(cur) = 0 Then
                ph.TextFrame.TextRange.Text = value
            Else
                ph.TextFrame.TextRange.Text = cur & vbCrLf & value
            End If
            Exit Sub
        End If
    Next i
End Sub

' Rename a shape. The new name then works as a ref_name in later actions
' (via shape_name aliases). Useful when the LLM forgot to set ref_name on
' add_shape and now needs to address it semantically.
Public Sub Do_set_shape_name(slideNum As Long, shapeId As Long, newName As String)
    Dim sh As Shape: Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2010, "Do_set_shape_name", "shape not found"
    If Len(Trim(newName)) = 0 Then Err.Raise vbObjectError + 2010, "Do_set_shape_name", "name empty"
    ' Reject duplicates on same slide — FindShapeByName picks first match and
    ' would silently shadow another shape with the same name.
    Dim existing As Shape: Set existing = FindShapeByName(slideNum, newName)
    If Not existing Is Nothing Then
        If existing.Id <> shapeId Then
            Err.Raise vbObjectError + 2010, "Do_set_shape_name", _
                      "name already in use on slide " & slideNum & ": " & newName
        End If
    End If
    sh.Name = newName
End Sub

' Atomic move + resize. Any of left/top/width/height may be omitted (pass a
' value < 0 via the dispatcher's sentinel logic); only specified fields change.
' The dispatcher decides which fields are present.
Public Sub Do_set_pos(slideNum As Long, shapeId As Long, _
                      leftPt As Single, topPt As Single, _
                      widthPt As Single, heightPt As Single, _
                      hasLeft As Boolean, hasTop As Boolean, _
                      hasWidth As Boolean, hasHeight As Boolean)
    Dim sh As Shape: Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2011, "Do_set_pos", "shape not found"
    If hasWidth Or hasHeight Then sh.LockAspectRatio = msoFalse
    If hasLeft Then sh.Left = leftPt
    If hasTop Then sh.Top = topPt
    If hasWidth Then sh.Width = widthPt
    If hasHeight Then sh.Height = heightPt
End Sub

' Set the shape's alt-text / accessibility description. Surfaces in screen
' readers and the Accessibility Checker. Pass "" to clear.
Public Sub Do_set_shape_alt_text(slideNum As Long, shapeId As Long, altText As String)
    Dim sh As Shape: Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2012, "Do_set_shape_alt_text", "shape not found"
    On Error Resume Next
    sh.AlternativeText = altText
    ' Title (the short "name" surfaced to screen readers) lives on Title in
    ' newer PowerPoint versions; mirror altText into it so single-source works.
    sh.Title = altText
    On Error GoTo 0
End Sub

' Toggle aspect-ratio lock on a shape. When locked, later resize_shape / set_pos
' with both width+height will still enforce the original aspect ratio. Mostly
' useful for pictures.
Public Sub Do_lock_aspect_ratio(slideNum As Long, shapeId As Long, value As Boolean)
    Dim sh As Shape: Set sh = FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 2013, "Do_lock_aspect_ratio", "shape not found"
    sh.LockAspectRatio = IIf(value, msoTrue, msoFalse)
End Sub

