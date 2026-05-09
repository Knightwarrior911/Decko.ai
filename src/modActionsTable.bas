Attribute VB_Name = "modActionsTable"
Option Explicit

Public Sub Do_add_table_row(slideNum As Long, shapeId As Long, afterRow As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_add_table_row", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_add_table_row", "shape is not a table"
    Dim tbl As Object: Set tbl = sh.Table
    Dim curRows As Long: curRows = tbl.Rows.Count
    If afterRow < 0 Or afterRow > curRows Then
        Err.Raise vbObjectError + 8003, "Do_add_table_row", "after_row out of range"
    End If
    If afterRow >= curRows Then
        ' Append at end — call without arg
        tbl.Rows.Add
    Else
        Dim insertAt As Long
        insertAt = afterRow + 1
        tbl.Rows.Add insertAt
    End If
End Sub

Public Sub Do_delete_table_row(slideNum As Long, shapeId As Long, rowNum As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_delete_table_row", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_delete_table_row", "shape is not a table"
    Dim tbl As Object: Set tbl = sh.Table
    If rowNum < 1 Or rowNum > tbl.Rows.Count Then
        Err.Raise vbObjectError + 8004, "Do_delete_table_row", "row out of range"
    End If
    tbl.Rows(rowNum).Delete
End Sub

Public Sub Do_add_table_col(slideNum As Long, shapeId As Long, afterCol As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_add_table_col", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_add_table_col", "shape is not a table"
    Dim tbl As Object: Set tbl = sh.Table
    Dim curCols As Long: curCols = tbl.Columns.Count
    If afterCol < 0 Or afterCol > curCols Then
        Err.Raise vbObjectError + 8005, "Do_add_table_col", "after_col out of range"
    End If
    If afterCol >= curCols Then
        tbl.Columns.Add
    Else
        Dim insertAt As Long
        insertAt = afterCol + 1
        tbl.Columns.Add insertAt
    End If
End Sub

Public Sub Do_delete_table_col(slideNum As Long, shapeId As Long, colNum As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_delete_table_col", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_delete_table_col", "shape is not a table"
    Dim tbl As Object: Set tbl = sh.Table
    If colNum < 1 Or colNum > tbl.Columns.Count Then
        Err.Raise vbObjectError + 8006, "Do_delete_table_col", "col out of range"
    End If
    tbl.Columns(colNum).Delete
End Sub

Public Sub Do_merge_cells(slideNum As Long, shapeId As Long, _
                          rowA As Long, colA As Long, _
                          rowB As Long, colB As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_merge_cells", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_merge_cells", "shape is not a table"
    Dim tbl As Object: Set tbl = sh.Table
    tbl.Cell(rowA, colA).Merge tbl.Cell(rowB, colB)
End Sub

Public Sub Do_add_table(slideNum As Long, rows As Long, cols As Long, _
                        leftPt As Single, topPt As Single, _
                        widthPt As Single, heightPt As Single, _
                        Optional refName As String = "")
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 8010, "Do_add_table", "slide_out_of_range"
    End If
    If rows < 1 Or cols < 1 Then
        Err.Raise vbObjectError + 8011, "Do_add_table", "rows and cols must be >= 1"
    End If
    Dim sh As Shape
    Set sh = pres.Slides(slideNum).Shapes.AddTable(rows, cols, leftPt, topPt, widthPt, heightPt)
    If Len(refName) > 0 Then sh.Name = refName
End Sub

Public Sub Do_set_table_col_width(slideNum As Long, shapeId As Long, _
                                   colNum As Long, widthPt As Single)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_set_table_col_width", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_set_table_col_width", "shape is not a table"
    Dim tbl As Object: Set tbl = sh.Table
    If colNum < 1 Or colNum > tbl.Columns.Count Then
        Err.Raise vbObjectError + 8012, "Do_set_table_col_width", "col out of range"
    End If
    tbl.Columns(colNum).Width = widthPt
End Sub

Public Sub Do_set_table_row_height(slideNum As Long, shapeId As Long, _
                                    rowNum As Long, heightPt As Single)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_set_table_row_height", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_set_table_row_height", "shape is not a table"
    Dim tbl As Object: Set tbl = sh.Table
    If rowNum < 1 Or rowNum > tbl.Rows.Count Then
        Err.Raise vbObjectError + 8013, "Do_set_table_row_height", "row out of range"
    End If
    tbl.Rows(rowNum).Height = heightPt
End Sub

Public Sub Do_set_cell_border(slideNum As Long, shapeId As Long, _
                               rowNum As Long, colNum As Long, _
                               side As String, hexColor As String, _
                               weightPt As Single, visible As Boolean)
    ' side: top/left/bottom/right/diag_down/diag_up/all
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_set_cell_border", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_set_cell_border", "shape is not a table"
    Dim tbl As Object: Set tbl = sh.Table
    Dim cell As Object: Set cell = tbl.Cell(rowNum, colNum)
    Dim sides() As Long
    Dim n As Long: n = ResolveBorderSides(side, sides)
    Dim i As Long
    For i = 0 To n - 1
        With cell.Borders(sides(i))
            If visible Then
                .Visible = msoTrue
                If Len(hexColor) > 0 Then .ForeColor.RGB = modActions.HexToRgb(hexColor)
                If weightPt > 0 Then .Weight = weightPt
            Else
                .Visible = msoFalse
            End If
        End With
    Next i
End Sub

Private Function ResolveBorderSides(side As String, ByRef out() As Long) As Long
    Select Case LCase(Trim(side))
        Case "top":       ReDim out(0 To 0): out(0) = 1: ResolveBorderSides = 1
        Case "left":      ReDim out(0 To 0): out(0) = 2: ResolveBorderSides = 1
        Case "bottom":    ReDim out(0 To 0): out(0) = 3: ResolveBorderSides = 1
        Case "right":     ReDim out(0 To 0): out(0) = 4: ResolveBorderSides = 1
        Case "diag_down": ReDim out(0 To 0): out(0) = 5: ResolveBorderSides = 1
        Case "diag_up":   ReDim out(0 To 0): out(0) = 6: ResolveBorderSides = 1
        Case "all"
            ReDim out(0 To 3)
            out(0) = 1: out(1) = 2: out(2) = 3: out(3) = 4
            ResolveBorderSides = 4
        Case Else: Err.Raise vbObjectError + 8014, "ResolveBorderSides", "unknown side: " & side
    End Select
End Function

Public Sub Do_set_cell_text_align(slideNum As Long, shapeId As Long, _
                                   rowNum As Long, colNum As Long, _
                                   hAlign As String, vAlign As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_set_cell_text_align", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_set_cell_text_align", "shape is not a table"
    Dim cell As Object: Set cell = sh.Table.Cell(rowNum, colNum)
    Select Case LCase(hAlign)
        Case "left":   cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
        Case "center": cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
        Case "right":  cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignRight
        Case "":       ' no change
        Case Else: Err.Raise vbObjectError + 8015, "Do_set_cell_text_align", "h_align: left/center/right"
    End Select
    Select Case LCase(vAlign)
        Case "top":    cell.Shape.TextFrame.VerticalAnchor = msoAnchorTop
        Case "middle": cell.Shape.TextFrame.VerticalAnchor = msoAnchorMiddle
        Case "bottom": cell.Shape.TextFrame.VerticalAnchor = msoAnchorBottom
        Case "":       ' no change
        Case Else: Err.Raise vbObjectError + 8016, "Do_set_cell_text_align", "v_align: top/middle/bottom"
    End Select
End Sub

Public Sub Do_apply_table_style(slideNum As Long, shapeId As Long, styleId As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_apply_table_style", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_apply_table_style", "shape is not a table"
    Dim guid As String: guid = ResolveTableStyleId(styleId)
    sh.Table.ApplyStyle guid, True
End Sub

Private Function ResolveTableStyleId(s As String) As String
    Dim k As String: k = LCase(Trim(s))
    ' Pass-through if it looks like a GUID
    If Left(k, 1) = "{" Then ResolveTableStyleId = s: Exit Function
    Select Case k
        Case "no_style_no_grid":         ResolveTableStyleId = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
        Case "no_style_with_grid":       ResolveTableStyleId = "{5940675A-B579-460E-94D1-54222C63F5DA}"
        Case "themed_style_1":           ResolveTableStyleId = "{2A488322-F2BA-4B5B-9748-0D474271808F}"
        Case "themed_style_1_accent1":   ResolveTableStyleId = "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}"
        Case "themed_style_1_accent2":   ResolveTableStyleId = "{3FFC9D24-7A4B-46FE-A0B8-A5B62BC9B89A}"
        Case "themed_style_2":           ResolveTableStyleId = "{D113A9D2-947B-4B7A-9F2A-77107C1A24F4}"
        Case "themed_style_2_accent1":   ResolveTableStyleId = "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}"
        Case "medium_style_2":           ResolveTableStyleId = "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}"
        Case "medium_style_2_accent1":   ResolveTableStyleId = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
        Case "medium_style_2_accent2":   ResolveTableStyleId = "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}"
        Case "dark_style_2":             ResolveTableStyleId = "{8EC20E35-A176-4012-BC5E-935CFFF8708E}"
        Case "dark_style_2_accent1":     ResolveTableStyleId = "{91EBBBCC-3105-4EDF-9809-1C9D6F9B5EB7}"
        Case "light_style_1":            ResolveTableStyleId = "{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}"
        Case "light_style_1_accent1":    ResolveTableStyleId = "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}"
        Case "light_style_2":            ResolveTableStyleId = "{7E9639D4-E3E2-4D34-9284-5A2195B3D0D8}"
        Case "light_style_2_accent1":    ResolveTableStyleId = "{69C7853C-536D-4A76-A0AE-DD22124D55A5}"
        Case Else: Err.Raise vbObjectError + 8020, "ResolveTableStyleId", "unknown style: " & s & " (use named or {GUID})"
    End Select
End Function

Public Sub Do_set_cell_fill(slideNum As Long, shapeId As Long, _
                             rowNum As Long, colNum As Long, hexColor As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_set_cell_fill", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_set_cell_fill", "shape is not a table"
    Dim cell As Object: Set cell = sh.Table.Cell(rowNum, colNum)
    cell.Shape.Fill.Visible = msoTrue
    cell.Shape.Fill.Solid
    cell.Shape.Fill.ForeColor.RGB = modActions.HexToRgb(hexColor)
End Sub
