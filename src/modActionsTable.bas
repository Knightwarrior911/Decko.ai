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

' Per-cell internal padding. All four sides; pass 0 for tight cells.
Public Sub Do_set_cell_padding(slideNum As Long, shapeId As Long, _
                                rowNum As Long, colNum As Long, _
                                leftPt As Double, rightPt As Double, _
                                topPt As Double, bottomPt As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8003, "Do_set_cell_padding", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8003, "Do_set_cell_padding", "shape is not a table"
    If leftPt < 0 Or rightPt < 0 Or topPt < 0 Or bottomPt < 0 Then _
        Err.Raise vbObjectError + 8003, "Do_set_cell_padding", "padding must be >= 0"
    Dim cell As Object: Set cell = sh.Table.Cell(rowNum, colNum)
    Dim cs As Shape: Set cs = cell.Shape
    If Not cs.HasTextFrame Then Exit Sub
    cs.TextFrame.MarginLeft   = leftPt
    cs.TextFrame.MarginRight  = rightPt
    cs.TextFrame.MarginTop    = topPt
    cs.TextFrame.MarginBottom = bottomPt
End Sub

' Empty a cell's text without removing the cell.
Public Sub Do_clear_cell_text(slideNum As Long, shapeId As Long, _
                                rowNum As Long, colNum As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8004, "Do_clear_cell_text", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8004, "Do_clear_cell_text", "shape is not a table"
    Dim cell As Object: Set cell = sh.Table.Cell(rowNum, colNum)
    On Error Resume Next
    cell.Shape.TextFrame.TextRange.Text = ""
    On Error GoTo 0
End Sub

' Toggle table style options independently of apply_table_style.
' Supports: header_row, total_row, banded_rows, first_column, last_column, banded_columns.
' All optional booleans — only the toggles you pass change.
Public Sub Do_set_table_style_options(slideNum As Long, shapeId As Long, _
                                       hasHeader As Boolean, headerVal As Boolean, _
                                       hasTotal As Boolean, totalVal As Boolean, _
                                       hasBandRows As Boolean, bandRowsVal As Boolean, _
                                       hasFirstCol As Boolean, firstColVal As Boolean, _
                                       hasLastCol As Boolean, lastColVal As Boolean, _
                                       hasBandCols As Boolean, bandColsVal As Boolean)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8005, "Do_set_table_style_options", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8005, "Do_set_table_style_options", "shape is not a table"
    Dim tbl As Table: Set tbl = sh.Table
    On Error Resume Next
    If hasHeader Then tbl.FirstRow = headerVal
    If hasTotal Then tbl.LastRow = totalVal
    If hasBandRows Then tbl.HorizBanding = bandRowsVal
    If hasFirstCol Then tbl.FirstCol = firstColVal
    If hasLastCol Then tbl.LastCol = lastColVal
    If hasBandCols Then tbl.VertBanding = bandColsVal
    On Error GoTo 0
End Sub

' Build a 2-column "image + bullets" table populated from a rows array.
' Per-row layout in the image_col cell: an image overlay plus a name caption
' anchored at the cell's top or bottom. The desc_col cell receives bullet
' paragraphs of the row's blurb. Image overlays are placed AFTER all cells
' are filled so the table's final cell rects are stable.
'
' Action shape (consumed via Do_build_image_grid_table_act):
'   {"type":"build_image_grid_table","slide":1,"ref_name":"tbl",
'    "pos":{"left":30,"top":60,"width":900,"height":500},
'    "image_col":1,"desc_col":2,
'    "name_position":"bottom","name_strip_pt":30,"image_pad_pt":6,
'    "header_row":false,
'    "col1_width_pt":300,"col2_width_pt":600,
'    "name_font":{"size":12,"bold":true,"color":"#15283C"},
'    "desc_font":{"size":10,"color":"#333333"},
'    "rows":[{"name":"Aerospace","image_path":"...","bullets":["..."]}]}
Public Sub Do_build_image_grid_table_act(act As Object)
    Dim stage As String: stage = "init"
    On Error GoTo bailout
    If Not act.Exists("rows") Then
        Err.Raise vbObjectError + 8030, "Do_build_image_grid_table", "rows array required"
    End If
    stage = "read_rows"
    Dim rows As Object: Set rows = act("rows")
    Dim n As Long: n = rows.Count
    If n = 0 Then
        Err.Raise vbObjectError + 8031, "Do_build_image_grid_table", "rows is empty"
    End If

    Dim slideNum As Long: slideNum = CLng(act("slide"))
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then slideNum = pres.Slides.Count

    Dim posD As Object: Set posD = act("pos")
    Dim totalLeft As Single: totalLeft = CSng(posD("left"))
    Dim totalTop As Single: totalTop = CSng(posD("top"))
    Dim totalW As Single: totalW = CSng(posD("width"))
    Dim totalH As Single: totalH = CSng(posD("height"))

    Dim imgCol As Long: imgCol = 1
    Dim descCol As Long: descCol = 2
    If act.Exists("image_col") Then imgCol = CLng(act("image_col"))
    If act.Exists("desc_col") Then descCol = CLng(act("desc_col"))

    Dim hasHeader As Boolean: hasHeader = False
    If act.Exists("header_row") Then hasHeader = modActions.ToBool(act("header_row"))

    Dim namePos As String: namePos = "bottom"
    If act.Exists("name_position") Then namePos = LCase(CStr(act("name_position")))

    Dim nameStrip As Single: nameStrip = 30
    If act.Exists("name_strip_pt") Then nameStrip = CSng(act("name_strip_pt"))
    Dim imgPad As Single: imgPad = 6
    If act.Exists("image_pad_pt") Then imgPad = CSng(act("image_pad_pt"))
    Dim imgFit As String: imgFit = "contain"
    If act.Exists("image_fit") Then imgFit = LCase(CStr(act("image_fit")))

    Dim totalRows As Long: totalRows = n + IIf(hasHeader, 1, 0)
    Dim cols As Long: cols = 2

    ' Create the table.
    stage = "add_table"
    Dim tShp As Shape
    Set tShp = pres.Slides(slideNum).Shapes.AddTable(totalRows, cols, _
        totalLeft, totalTop, totalW, totalH)
    If act.Exists("ref_name") Then
        If Len(CStr(act("ref_name"))) > 0 Then tShp.Name = CStr(act("ref_name"))
    End If

    ' Optional column widths.
    Dim w1 As Single, w2 As Single
    w1 = -1: w2 = -1
    If act.Exists("col1_width_pt") Then w1 = CSng(act("col1_width_pt"))
    If act.Exists("col2_width_pt") Then w2 = CSng(act("col2_width_pt"))
    If w1 < 0 And w2 < 0 Then
        w1 = totalW * 0.4
        w2 = totalW - w1
    ElseIf w1 < 0 Then
        w1 = totalW - w2
    ElseIf w2 < 0 Then
        w2 = totalW - w1
    End If
    On Error Resume Next
    tShp.Table.Columns(1).Width = w1
    tShp.Table.Columns(2).Width = w2
    Err.Clear
    On Error GoTo bailout

    ' Force uniform row heights so caller's totalH actually applies. PowerPoint
    ' otherwise distributes by content.
    Dim rh As Single: rh = totalH / totalRows
    Dim ri As Long
    On Error Resume Next
    For ri = 1 To totalRows
        tShp.Table.Rows(ri).Height = rh
    Next ri
    Err.Clear
    On Error GoTo bailout

    stage = "set_widths_done"
    ' Header row (optional) - bold, no bullets.
    Dim startDataRow As Long: startDataRow = 1
    If hasHeader Then
        startDataRow = 2
        Dim hdrImg As String: hdrImg = ""
        Dim hdrDesc As String: hdrDesc = ""
        If act.Exists("header_text_image") Then hdrImg = CStr(act("header_text_image"))
        If act.Exists("header_text_desc") Then hdrDesc = CStr(act("header_text_desc"))
        WriteHeaderCell tShp.Table.Cell(1, imgCol), hdrImg
        WriteHeaderCell tShp.Table.Cell(1, descCol), hdrDesc
    End If

    ' Caption + bullet font configs.
    Dim nameFontSize As Single: nameFontSize = 12
    Dim nameFontColor As String: nameFontColor = "#15283C"
    Dim nameFontBold As Boolean: nameFontBold = True
    If act.Exists("name_font") Then
        Dim nf As Object: Set nf = act("name_font")
        If nf.Exists("size") Then nameFontSize = CSng(nf("size"))
        If nf.Exists("color") Then nameFontColor = CStr(nf("color"))
        If nf.Exists("bold") Then nameFontBold = modActions.ToBool(nf("bold"))
    End If
    Dim descFontSize As Single: descFontSize = 10
    Dim descFontColor As String: descFontColor = "#333333"
    If act.Exists("desc_font") Then
        Dim df As Object: Set df = act("desc_font")
        If df.Exists("size") Then descFontSize = CSng(df("size"))
        If df.Exists("color") Then descFontColor = CStr(df("color"))
    End If

    ' Pass 1 - fill cell text (no images yet).
    stage = "pass1_start"
    Dim r As Long
    For r = 1 To n
        stage = "pass1_row=" & r
        Dim row As Object: Set row = rows(r)
        Dim tableRow As Long: tableRow = r + (startDataRow - 1)

        Dim nameTxt As String: nameTxt = ""
        If row.Exists("name") Then nameTxt = CStr(row("name"))
        stage = "pass1_row=" & r & "_name"
        WriteNameCell tShp.Table.Cell(tableRow, imgCol), nameTxt, _
                      namePos, nameFontSize, nameFontColor, nameFontBold

        Dim bullets As Object
        If row.Exists("bullets") Then Set bullets = row("bullets") Else Set bullets = Nothing
        stage = "pass1_row=" & r & "_bullets"
        WriteBulletCell tShp.Table.Cell(tableRow, descCol), bullets, _
                        descFontSize, descFontColor
    Next r

    ' Pass 2 - download (if URL) and overlay images on imageCol cells.
    stage = "pass2_start"
    For r = 1 To n
        stage = "pass2_row=" & r
        Dim row2 As Object: Set row2 = rows(r)
        Dim tableRow2 As Long: tableRow2 = r + (startDataRow - 1)

        Dim imgPath As String: imgPath = ""
        If row2.Exists("image_path") Then imgPath = CStr(row2("image_path"))
        If Len(imgPath) = 0 And row2.Exists("image_url") Then
            Dim url As String: url = CStr(row2("image_url"))
            If Len(url) > 0 Then imgPath = DownloadInlineImage(url)
        End If
        If Len(imgPath) > 0 And FileExistsLocal(imgPath) Then
            stage = "pass2_row=" & r & "_overlay"
            PlaceImageOverImageCell pres.Slides(slideNum), _
                                     tShp.Table.Cell(tableRow2, imgCol), _
                                     imgPath, namePos, nameStrip, imgPad, imgFit
        End If
    Next r
    Exit Sub
bailout:
    Err.Raise vbObjectError + 8033, "Do_build_image_grid_table", _
        "stage=" & stage & " err=" & Err.Description
End Sub

Private Sub WriteHeaderCell(cell As Object, txt As String)
    cell.Shape.TextFrame.TextRange.Text = txt
    cell.Shape.TextFrame.TextRange.Font.Bold = msoTrue
    cell.Shape.TextFrame.VerticalAnchor = msoAnchorMiddle
    cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
End Sub

Private Sub WriteNameCell(cell As Object, txt As String, namePos As String, _
                           sz As Single, hexColor As String, bold As Boolean)
    Dim tr As Object: Set tr = cell.Shape.TextFrame.TextRange
    tr.Text = txt
    tr.Font.Size = sz
    tr.Font.Bold = IIf(bold, msoTrue, msoFalse)
    On Error Resume Next
    tr.Font.Color.RGB = modActions.HexToRgb(hexColor)
    On Error GoTo 0
    tr.ParagraphFormat.Alignment = ppAlignCenter
    Select Case namePos
        Case "top":    cell.Shape.TextFrame.VerticalAnchor = msoAnchorTop
        Case Else:     cell.Shape.TextFrame.VerticalAnchor = msoAnchorBottom
    End Select
End Sub

Private Sub WriteBulletCell(cell As Object, bullets As Object, _
                             sz As Single, hexColor As String)
    Dim tr As Object: Set tr = cell.Shape.TextFrame.TextRange
    cell.Shape.TextFrame.VerticalAnchor = msoAnchorMiddle

    If bullets Is Nothing Then
        tr.Text = ""
        Exit Sub
    End If
    Dim m As Long: m = bullets.Count
    If m = 0 Then
        tr.Text = ""
        Exit Sub
    End If

    Dim joined As String: joined = ""
    Dim i As Long
    For i = 1 To m
        If i > 1 Then joined = joined & vbCr
        joined = joined & CStr(bullets(i))
    Next i
    tr.Text = joined
    tr.Font.Size = sz
    On Error Resume Next
    tr.Font.Color.RGB = modActions.HexToRgb(hexColor)
    On Error GoTo 0
    tr.ParagraphFormat.Alignment = ppAlignLeft

    For i = 1 To m
        On Error Resume Next
        With tr.Paragraphs(i).ParagraphFormat.Bullet
            .Type = 1            ' ppBulletUnnumbered
            .Character = 8226    ' bullet •
        End With
        On Error GoTo 0
    Next i
End Sub

Private Sub PlaceImageOverImageCell(sl As Slide, cell As Object, imgPath As String, _
                                     namePos As String, nameStrip As Single, _
                                     pad As Single, fit As String)
    Dim cl As Single: cl = cell.Shape.Left
    Dim ct As Single: ct = cell.Shape.Top
    Dim cw As Single: cw = cell.Shape.Width
    Dim ch As Single: ch = cell.Shape.Height

    Dim picL As Single, picT As Single, picW As Single, picH As Single
    picL = cl + pad
    picW = cw - 2 * pad
    picH = ch - nameStrip - 2 * pad
    If picH < 10 Then picH = 10
    If picW < 10 Then picW = 10
    If namePos = "top" Then
        picT = ct + nameStrip + pad
    Else
        picT = ct + pad
    End If

    imgPath = Replace(imgPath, "/", "\")   ' AddPicture requires backslashes
    If fit = "stretch" Then
        On Error Resume Next
        Dim pic As Shape
        Set pic = sl.Shapes.AddPicture(FileName:=imgPath, _
            LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
            Left:=picL, Top:=picT, Width:=picW, Height:=picH)
        If Err.Number = 0 And Not pic Is Nothing Then
            pic.LockAspectRatio = msoFalse
            pic.Width = picW
            pic.Height = picH
        End If
        Err.Clear
        On Error GoTo 0
    Else
        ' default = contain (preserve aspect, letterbox)
        modActionsImage.AddPictureContain sl, imgPath, picL, picT, picW, picH
    End If
End Sub

Private Function FileExistsLocal(p As String) As Boolean
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    FileExistsLocal = fso.FileExists(p)
End Function

Private Function DownloadInlineImage(url As String) As String
    Dim deckDir As String
    deckDir = ActivePresentation.Path
    Dim folder As String: folder = deckDir & "\assets\inline"
    On Error Resume Next
    If Dir(folder, vbDirectory) = "" Then
        Dim parent As String: parent = deckDir & "\assets"
        If Dir(parent, vbDirectory) = "" Then MkDir parent
        MkDir folder
    End If
    On Error GoTo 0
    Dim ts As String: ts = Format(Now, "yyyymmddhhnnss")
    Dim ext As String: ext = "jpg"
    Dim u As String: u = LCase(url)
    Dim q As Long: q = InStr(u, "?")
    If q > 0 Then u = Left(u, q - 1)
    Dim dot As Long: dot = InStrRev(u, ".")
    If dot > 0 Then
        Dim e As String: e = Mid(u, dot + 1)
        Select Case e
            Case "jpg", "jpeg", "png", "gif", "webp", "bmp"
                ext = e
        End Select
    End If
    Dim destPath As String
    destPath = folder & "\inline_" & ts & "_" & CLng(Rnd * 1000000) & "." & ext
    On Error Resume Next
    modActionsWeb.Do_download_image url, destPath
    If Err.Number <> 0 Then
        DownloadInlineImage = ""
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    DownloadInlineImage = destPath
End Function