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
    End If

    d.Add "fill", BuildFillHex(sh)

    If sh.HasTable Then
        d.Add "table", BuildTableDict(sh.Table)
    End If

    If sh.Type = msoPicture Then
        d.Add "picture", BuildPictureDict(sh)
    End If

    Set BuildShapeDict = d
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
    ElseIf sh.HasTextFrame Then
        ClassifyShapeType = "textbox"
    ElseIf sh.HasTable Then
        ClassifyShapeType = "table"
    ElseIf sh.Type = msoPicture Then
        ClassifyShapeType = "picture"
    ElseIf sh.HasChart Then
        ClassifyShapeType = "chart"
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
    On Error Resume Next
    d.Add "filename", sh.PictureFormat.SourceFileName
    On Error GoTo 0
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
        col.Add d
    Next i
    Set BuildRunsCollection = col
End Function
