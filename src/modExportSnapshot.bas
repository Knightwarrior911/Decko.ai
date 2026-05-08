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
        End If
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
