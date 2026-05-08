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
    If Len(h) <> 6 Then Err.Raise vbObjectError + 2003, "HexToRgb", "expected #RRGGBB"
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Mid(h, 1, 2))
    g = CLng("&H" & Mid(h, 3, 2))
    b = CLng("&H" & Mid(h, 5, 2))
    HexToRgb = RGB(r, g, b)
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
