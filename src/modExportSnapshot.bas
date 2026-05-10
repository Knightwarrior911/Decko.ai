Attribute VB_Name = "modExportSnapshot"
Option Explicit

' Build a JSON snapshot of ActivePresentation.
' V0: text shapes only — slides, slide_number, shapes[].{shape_id, shape_name, type, text}
' Later tasks add pos, font, fill, table, picture, theme.
Public Function BuildSnapshotJson() As String
    Dim pres As Presentation
    Set pres = ActivePresentation

    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root.Add "deck", BuildDeckDict(pres)
    root.Add "slides", BuildSlidesCollection(pres)

    BuildSnapshotJson = modJSON.ConvertToJson(root)
End Function

Private Function BuildDeckDict(pres As Presentation) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "path", pres.FullName
    d.Add "slide_width_pt", pres.PageSetup.SlideWidth
    d.Add "slide_height_pt", pres.PageSetup.SlideHeight
    d.Add "theme", BuildThemeDict(pres)
    Set BuildDeckDict = d
End Function

Private Function BuildSlidesCollection(pres As Presentation) As Collection
    Dim col As New Collection
    Dim i As Long
    For i = 1 To pres.Slides.Count
        col.Add BuildSlideDict(pres.Slides(i))
    Next i
    Set BuildSlidesCollection = col
End Function

Private Function BuildSlideDict(sl As Slide) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "slide_number", sl.SlideIndex
    d.Add "layout_name", sl.CustomLayout.Name
    d.Add "occupied_rects", BuildOccupiedRects(sl)
    d.Add "speaker_notes", BuildSpeakerNotes(sl)
    d.Add "shapes", BuildShapesCollection(sl)
    Set BuildSlideDict = d
End Function

Private Function BuildShapesCollection(sl As Slide) As Collection
    Dim col As New Collection
    Dim sh As Shape
    For Each sh In sl.Shapes
        col.Add BuildShapeDict(sh)
    Next sh
    Set BuildShapesCollection = col
End Function

Private Function BuildShapeDict(sh As Shape) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "shape_id", sh.Id
    d.Add "shape_name", sh.Name
    d.Add "type", ClassifyShapeType(sh)
    d.Add "pos", BuildPosDict(sh)

    If sh.HasTextFrame Then
        If sh.TextFrame.HasText Then
            d.Add "text", sh.TextFrame.TextRange.Text
            d.Add "font", BuildFontDict(sh.TextFrame.TextRange.Font)
            d.Add "paragraphs", BuildParagraphsCollection(sh.TextFrame.TextRange)
        End If
        Dim tfDict As Object: Set tfDict = CreateObject("Scripting.Dictionary")
        tfDict.Add "vertical_align", VerticalAnchorName(sh.TextFrame.VerticalAnchor)
        tfDict.Add "word_wrap", CBool(sh.TextFrame.WordWrap = msoTrue)
        tfDict.Add "auto_size", AutoSizeName(sh.TextFrame.AutoSize)
        Dim mDict As Object: Set mDict = CreateObject("Scripting.Dictionary")
        mDict.Add "left", CDbl(sh.TextFrame.MarginLeft)
        mDict.Add "right", CDbl(sh.TextFrame.MarginRight)
        mDict.Add "top", CDbl(sh.TextFrame.MarginTop)
        mDict.Add "bottom", CDbl(sh.TextFrame.MarginBottom)
        tfDict.Add "margin", mDict
        d.Add "text_frame", tfDict
    End If

    d.Add "fill", BuildFillHex(sh)

    If sh.HasTable Then
        d.Add "table", BuildTableDict(sh.Table)
        d.Add "table_extra", BuildTableExtra(sh.Table)
    End If

    If sh.HasChart Then
        d.Add "chart", BuildChartDict(sh.Chart)
    End If

    If sh.Type = msoPicture Then
        d.Add "picture", BuildPictureDict(sh)
    End If

    If sh.Type = msoGroup Then
        d.Add "group_children", BuildGroupChildren(sh)
    End If

    Set BuildShapeDict = d
End Function

Private Function IsConnectorShape(sh As Shape) As Boolean
    On Error Resume Next
    Dim c As Long: c = sh.Connector
    If Err.Number = 0 Then
        IsConnectorShape = (c = msoTrue)
    Else
        IsConnectorShape = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function ClassifyShapeType(sh As Shape) As String
    If sh.Type = msoPlaceholder Then
        Select Case sh.PlaceholderFormat.Type
            Case ppPlaceholderTitle, ppPlaceholderCenterTitle
                ClassifyShapeType = "title"
            Case ppPlaceholderBody, ppPlaceholderObject, ppPlaceholderSubtitle
                ClassifyShapeType = "body"
            Case Else
                ClassifyShapeType = "other"
        End Select
    ElseIf sh.HasChart Then
        ClassifyShapeType = "chart"
    ElseIf sh.HasTable Then
        ClassifyShapeType = "table"
    ElseIf sh.Type = msoPicture Then
        ClassifyShapeType = "picture"
    ElseIf IsConnectorShape(sh) Then
        ClassifyShapeType = "connector"
    ElseIf sh.Type = msoLine Then
        ClassifyShapeType = "line"
    ElseIf sh.Type = msoGroup Then
        ClassifyShapeType = "group"
    ElseIf sh.HasTextFrame Then
        ClassifyShapeType = "textbox"
    Else
        ClassifyShapeType = "other"
    End If
End Function

Private Function BuildPosDict(sh As Shape) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "left", CDbl(sh.Left)
    d.Add "top", CDbl(sh.Top)
    d.Add "width", CDbl(sh.Width)
    d.Add "height", CDbl(sh.Height)
    Set BuildPosDict = d
End Function

Private Function BuildFontDict(fnt As Font) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "name", fnt.Name
    ' fnt.Size returns Single; can be -1 if mixed across runs
    If fnt.Size > 0 Then
        d.Add "size", CDbl(fnt.Size)
    Else
        d.Add "size", Null
    End If
    d.Add "bold", (fnt.Bold = msoTrue)
    d.Add "italic", (fnt.Italic = msoTrue)
    d.Add "color", RgbToHex(fnt.Color.RGB)
    d.Add "underline", (fnt.Underline = msoTrue)
    d.Add "subscript", (fnt.BaselineOffset < -0.001)
    d.Add "superscript", (fnt.BaselineOffset > 0.001)
    Set BuildFontDict = d
End Function

Public Function RgbToHex(ByVal rgbVal As Long) As String
    Dim r As Long, g As Long, b As Long
    r = rgbVal And &HFF
    g = (rgbVal \ &H100) And &HFF
    b = (rgbVal \ &H10000) And &HFF
    RgbToHex = "#" & UCase(Right("00" & Hex(r), 2) & Right("00" & Hex(g), 2) & Right("00" & Hex(b), 2))
End Function

Private Function BuildFillHex(sh As Shape) As Variant
    On Error Resume Next
    Dim v As Variant
    v = Null
    If sh.Fill.Visible = msoTrue Then
        If sh.Fill.Type = msoFillSolid Then
            v = RgbToHex(sh.Fill.ForeColor.RGB)
        End If
    End If
    BuildFillHex = v
End Function

Private Function BuildTableDict(tbl As Table) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "rows", tbl.Rows.Count
    d.Add "cols", tbl.Columns.Count

    Dim cells As Collection
    Dim rowCol As Collection
    Dim r As Long, c As Long
    Set cells = New Collection
    For r = 1 To tbl.Rows.Count
        Set rowCol = New Collection
        For c = 1 To tbl.Columns.Count
            rowCol.Add BuildCellDict(tbl.Cell(r, c))
        Next c
        cells.Add rowCol
    Next r
    d.Add "cells", cells
    Set BuildTableDict = d
End Function

Private Function BuildCellDict(cell As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim shp As Shape
    Set shp = cell.Shape
    Dim txt As String
    txt = ""
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            txt = shp.TextFrame.TextRange.Text
        End If
    End If
    d.Add "text", txt
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            d.Add "font", BuildFontDict(shp.TextFrame.TextRange.Font)
        End If
    End If
    d.Add "fill", BuildFillHex(shp)
    Set BuildCellDict = d
End Function

Private Function BuildPictureDict(sh As Shape) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim sfn As String
    sfn = ""
    On Error Resume Next
    Dim shObj As Object
    Set shObj = sh
    sfn = CStr(shObj.PictureFormat.SourceFullName)
    If Len(sfn) = 0 Then sfn = CStr(shObj.PictureFormat.SourceFileName)
    Err.Clear
    On Error GoTo 0
    If Len(sfn) > 0 Then d.Add "filename", sfn
    Set BuildPictureDict = d
End Function

Private Function BuildThemeDict(pres As Presentation) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim scheme As Object
    Set scheme = pres.SlideMaster.Theme.ThemeColorScheme

    d.Add "dk1", RgbToHex(scheme.Colors(msoThemeDark1).RGB)
    d.Add "lt1", RgbToHex(scheme.Colors(msoThemeLight1).RGB)
    d.Add "dk2", RgbToHex(scheme.Colors(msoThemeDark2).RGB)
    d.Add "lt2", RgbToHex(scheme.Colors(msoThemeLight2).RGB)
    d.Add "accent1", RgbToHex(scheme.Colors(msoThemeAccent1).RGB)
    d.Add "accent2", RgbToHex(scheme.Colors(msoThemeAccent2).RGB)
    d.Add "accent3", RgbToHex(scheme.Colors(msoThemeAccent3).RGB)
    d.Add "accent4", RgbToHex(scheme.Colors(msoThemeAccent4).RGB)
    d.Add "accent5", RgbToHex(scheme.Colors(msoThemeAccent5).RGB)
    d.Add "accent6", RgbToHex(scheme.Colors(msoThemeAccent6).RGB)
    d.Add "hlink", RgbToHex(scheme.Colors(msoThemeHyperlink).RGB)
    d.Add "folHlink", RgbToHex(scheme.Colors(msoThemeFollowedHyperlink).RGB)

    Set BuildThemeDict = d
End Function

Private Function BuildOccupiedRects(sl As Slide) As Collection
    Dim col As New Collection
    Dim sh As Shape
    For Each sh In sl.Shapes
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d.Add "shape_id", sh.Id
        d.Add "left", CDbl(sh.Left)
        d.Add "top", CDbl(sh.Top)
        d.Add "right", CDbl(sh.Left + sh.Width)
        d.Add "bottom", CDbl(sh.Top + sh.Height)
        col.Add d
    Next sh
    Set BuildOccupiedRects = col
End Function

Private Function BuildSpeakerNotes(sl As Slide) As String
    On Error Resume Next
    Dim notesText As String
    notesText = ""
    Dim notesPg As Object
    Set notesPg = sl.NotesPage
    Dim ph As Object
    Dim i As Long
    For i = 1 To notesPg.Shapes.Placeholders.Count
        Set ph = notesPg.Shapes.Placeholders(i)
        If ph.HasTextFrame Then
            If ph.PlaceholderFormat.Type = ppPlaceholderBody Then
                notesText = ph.TextFrame.TextRange.Text
                Exit For
            End If
        End If
    Next i
    BuildSpeakerNotes = notesText
End Function

Private Function BuildParagraphsCollection(tr As TextRange) As Collection
    Dim col As New Collection
    Dim n As Long
    n = tr.Paragraphs().Count
    Dim i As Long
    For i = 1 To n
        col.Add BuildParagraphDict(tr.Paragraphs(i), i - 1)
    Next i
    Set BuildParagraphsCollection = col
End Function

Private Function BuildParagraphDict(para As TextRange, zeroIdx As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "index", zeroIdx
    d.Add "text", para.Text
    d.Add "bullet_style", BulletStyleName(para.ParagraphFormat.Bullet.Type, para.ParagraphFormat.Bullet.Style)
    d.Add "indent_level", CLng(para.IndentLevel) - 1
    d.Add "runs", BuildRunsCollection(para)
    d.Add "alignment", AlignmentName(para.ParagraphFormat.Alignment)
    ' line_spacing always emitted as a multiple
    Dim lsMul As Double
    If para.ParagraphFormat.LineRuleWithin = msoTrue Then
        lsMul = CDbl(para.ParagraphFormat.SpaceWithin)
    Else
        Dim baseSize As Double
        baseSize = 18#
        On Error Resume Next
        baseSize = CDbl(para.Runs(1).Font.Size)
        On Error GoTo 0
        If baseSize > 0 Then
            lsMul = CDbl(para.ParagraphFormat.SpaceWithin) / baseSize
        Else
            lsMul = 1#
        End If
    End If
    d.Add "line_spacing", lsMul
    d.Add "space_before", CDbl(para.ParagraphFormat.SpaceBefore)
    d.Add "space_after", CDbl(para.ParagraphFormat.SpaceAfter)
    Set BuildParagraphDict = d
End Function

Private Function BulletStyleName(btype As Long, bstyle As Long) As String
    Select Case btype
        Case 0: BulletStyleName = "none"
        Case 2: BulletStyleName = "number"
        Case 3: BulletStyleName = "image"
        Case Else: BulletStyleName = "disc"
    End Select
End Function

Private Function BuildRunsCollection(para As TextRange) As Collection
    Dim col As New Collection
    Dim n As Long
    n = para.Runs().Count
    Dim i As Long
    For i = 1 To n
        Dim run As TextRange
        Set run = para.Runs(i)
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d.Add "text", run.Text
        d.Add "font", BuildFontDict(run.Font)
        d.Add "hyperlink", RunHyperlink(run)
        col.Add d
    Next i
    Set BuildRunsCollection = col
End Function

Private Function BuildGroupChildren(sh As Shape) As Collection
    Dim col As New Collection
    Dim child As Shape
    For Each child In sh.GroupItems
        col.Add BuildShapeDict(child)
    Next child
    Set BuildGroupChildren = col
End Function

Private Function BuildChartDict(ch As Chart) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "is_native", True
    d.Add "type", ChartTypeName(ch.ChartType)

    On Error Resume Next
    If ch.HasTitle Then
        d.Add "title", ch.ChartTitle.Text
    Else
        d.Add "title", Null
    End If
    Err.Clear

    Dim ax As Object
    Set ax = CreateObject("Scripting.Dictionary")
    Dim xt As String, yt As String
    xt = "": yt = ""
    On Error Resume Next
    If ch.HasAxis(1) Then
        If ch.Axes(1).HasTitle Then xt = ch.Axes(1).AxisTitle.Text
    End If
    If ch.HasAxis(2) Then
        If ch.Axes(2).HasTitle Then yt = ch.Axes(2).AxisTitle.Text
    End If
    Err.Clear
    On Error GoTo 0
    ax.Add "x", xt
    ax.Add "y", yt
    d.Add "axis_titles", ax

    On Error Resume Next
    Dim leg As String: leg = "none"
    If ch.HasLegend Then leg = LegendPositionName(ch.Legend.Position)
    Err.Clear
    On Error GoTo 0
    d.Add "legend_position", leg

    d.Add "series", BuildSeriesCollection(ch)

    Set BuildChartDict = d
End Function

Private Function BuildSeriesCollection(ch As Chart) As Collection
    Dim col As New Collection
    On Error Resume Next
    Dim n As Long: n = ch.SeriesCollection.Count
    Dim i As Long
    For i = 1 To n
        Dim s As Object
        Set s = CreateObject("Scripting.Dictionary")
        s.Add "name", ch.SeriesCollection(i).Name
        Dim cats As Variant
        cats = ch.SeriesCollection(i).XValues
        s.Add "categories", VariantArrayToCollection(cats)
        Dim vals As Variant
        vals = ch.SeriesCollection(i).Values
        s.Add "values", VariantArrayToCollection(vals)
        col.Add s
    Next i
    Err.Clear
    On Error GoTo 0
    Set BuildSeriesCollection = col
End Function

Private Function VariantArrayToCollection(arr As Variant) As Variant
    On Error Resume Next
    Dim col As New Collection
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        col.Add arr(i)
    Next i
    If Err.Number <> 0 Then
        Err.Clear
        VariantArrayToCollection = Null
        Exit Function
    End If
    Set VariantArrayToCollection = col
End Function

Private Function ChartTypeName(t As Long) As String
    Select Case t
        Case 51: ChartTypeName = "columnClustered"
        Case 52: ChartTypeName = "columnStacked"
        Case 4:  ChartTypeName = "line"
        Case 5:  ChartTypeName = "pie"
        Case 57: ChartTypeName = "barClustered"
        Case 1:  ChartTypeName = "area"
        Case -4169: ChartTypeName = "scatter"
        Case Else: ChartTypeName = "type_" & t
    End Select
End Function

Private Function LegendPositionName(p As Long) As String
    Select Case p
        Case -4131: LegendPositionName = "left"
        Case -4152: LegendPositionName = "right"
        Case -4160: LegendPositionName = "top"
        Case -4107: LegendPositionName = "bottom"
        Case 2:     LegendPositionName = "corner"
        Case Else:  LegendPositionName = "right"
    End Select
End Function

Private Function BuildTableExtra(tbl As Table) As Object
    ' PowerPoint COM Cell object does not expose Merged/RowSpan/ColSpan.
    ' We detect merges by comparing cell.Shape.Width/Height against the
    ' individual column/row widths stored in tbl.Columns(c).Width.
    ' A merged cell has Width > its single column width.
    ' "Leader" cells are those where the previous cell in the same axis
    ' has a DIFFERENT (smaller) width — i.e. the span starts here.
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim merges As New Collection

    Dim nRows As Long: nRows = tbl.Rows.Count
    Dim nCols As Long: nCols = tbl.Columns.Count

    ' Cache individual column widths and row heights
    Dim colW() As Double
    Dim rowH() As Double
    ReDim colW(1 To nCols)
    ReDim rowH(1 To nRows)
    Dim ci As Long, ri As Long
    For ci = 1 To nCols
        colW(ci) = tbl.Columns(ci).Width
    Next ci
    For ri = 1 To nRows
        rowH(ri) = tbl.Rows(ri).Height
    Next ri

    Dim r As Long, c As Long
    For r = 1 To nRows
        For c = 1 To nCols
            Dim cellObj As Object
            Set cellObj = tbl.Cell(r, c)
            Dim cw As Double: cw = cellObj.Shape.Width
            Dim ch As Double: ch = cellObj.Shape.Height

            ' Compute spans by ratio (round to nearest integer)
            Dim cs As Long: cs = CLng(cw / colW(c) + 0.4999)
            Dim rs As Long: rs = CLng(ch / rowH(r) + 0.4999)
            If cs < 1 Then cs = 1
            If rs < 1 Then rs = 1

            If cs > 1 Or rs > 1 Then
                ' Only record if this is the leader cell.
                ' Leader: no previous cell in same span direction has same dimensions.
                Dim isLeader As Boolean: isLeader = True
                If cs > 1 And c > 1 Then
                    Dim prevCell As Object
                    Set prevCell = tbl.Cell(r, c - 1)
                    If Abs(prevCell.Shape.Width - cw) < 1 Then
                        isLeader = False
                    End If
                End If
                If rs > 1 And r > 1 And isLeader Then
                    Dim aboveCell As Object
                    Set aboveCell = tbl.Cell(r - 1, c)
                    If Abs(aboveCell.Shape.Height - ch) < 1 Then
                        isLeader = False
                    End If
                End If
                If isLeader Then
                    Dim m As Object
                    Set m = CreateObject("Scripting.Dictionary")
                    m.Add "row", r
                    m.Add "col", c
                    m.Add "row_span", rs
                    m.Add "col_span", cs
                    merges.Add m
                End If
            End If
        Next c
    Next r

    d.Add "merged_cells", merges
    Set BuildTableExtra = d
End Function

Private Function VerticalAnchorName(v As Long) As String
    Select Case v
        Case 1: VerticalAnchorName = "top"           ' msoAnchorTop
        Case 2: VerticalAnchorName = "top"           ' msoAnchorTopBaseline
        Case 3: VerticalAnchorName = "middle"        ' msoAnchorMiddle
        Case 4: VerticalAnchorName = "bottom"        ' msoAnchorBottom
        Case 5: VerticalAnchorName = "bottom"        ' msoAnchorBottomBaseline
        Case Else: VerticalAnchorName = "top"
    End Select
End Function

Private Function RunHyperlink(r As TextRange) As Variant
    Dim addr As String
    Dim act As Long
    act = 0
    On Error Resume Next
    act = r.ActionSettings(1).Action  ' 1 = ppMouseClick
    addr = r.ActionSettings(1).Hyperlink.Address
    On Error GoTo 0
    ' Treat as "no link" when Action is ppActionNone (0) or address is empty.
    If act = 0 Or Len(addr) = 0 Then
        RunHyperlink = Null
    Else
        RunHyperlink = addr
    End If
End Function

Private Function AlignmentName(v As Long) As String
    Select Case v
        Case 1: AlignmentName = "left"     ' ppAlignLeft
        Case 2: AlignmentName = "center"   ' ppAlignCenter
        Case 3: AlignmentName = "right"    ' ppAlignRight
        Case 4: AlignmentName = "justify"  ' ppAlignJustify
        Case Else: AlignmentName = "left"
    End Select
End Function

Private Function AutoSizeName(v As Long) As String
    Select Case v
        Case 0: AutoSizeName = "none"           ' ppAutoSizeNone
        Case 1: AutoSizeName = "shape_to_text"  ' ppAutoSizeShapeToFitText
        Case 2: AutoSizeName = "text_to_shape"  ' ppAutoSizeTextToFitShape
        Case Else: AutoSizeName = "none"
    End Select
End Function
