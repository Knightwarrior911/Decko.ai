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
