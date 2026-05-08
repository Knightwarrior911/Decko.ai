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
