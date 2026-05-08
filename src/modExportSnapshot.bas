Attribute VB_Name = "modExportSnapshot"
Option Explicit

' Build a JSON snapshot of ActivePresentation.
' V0: text shapes only - slides, slide_number, shapes[].{shape_id, shape_name, type, text}
' Later tasks add pos, font, fill, table, picture, theme.
Public Function BuildSnapshotJson() As String
    On Error GoTo ErrHandler
    Dim pres As Presentation
    Set pres = ActivePresentation

    Dim parts() As String
    Dim nSlides As Long
    nSlides = pres.Slides.Count

    Dim slideParts() As String
    ReDim slideParts(1 To nSlides)
    Dim i As Long
    For i = 1 To nSlides
        slideParts(i) = BuildSlideJson(pres.Slides(i))
    Next i

    Dim deckJson As String
    deckJson = "{""path"":" & JsonStr(pres.FullName) & _
               ",""slide_width_pt"":" & CStr(pres.PageSetup.SlideWidth) & _
               ",""slide_height_pt"":" & CStr(pres.PageSetup.SlideHeight) & "}"

    Dim slidesJson As String
    slidesJson = "[" & Join(slideParts, ",") & "]"

    BuildSnapshotJson = "{""deck"":" & deckJson & ",""slides"":" & slidesJson & "}"
    Exit Function
ErrHandler:
    BuildSnapshotJson = "{""error"":""" & JsonEscStr(Err.Description) & """}"
End Function

Private Function BuildSlideJson(sl As Slide) As String
    Dim nShapes As Long
    nShapes = sl.Shapes.Count
    Dim shapeParts() As String
    ReDim shapeParts(1 To nShapes)
    Dim j As Long
    For j = 1 To nShapes
        shapeParts(j) = BuildShapeJson(sl.Shapes(j))
    Next j
    Dim shapesJson As String
    shapesJson = "[" & Join(shapeParts, ",") & "]"
    BuildSlideJson = "{""slide_number"":" & CStr(sl.SlideIndex) & _
                     ",""layout_name"":" & JsonStr(sl.CustomLayout.Name) & _
                     ",""shapes"":" & shapesJson & "}"
End Function

Private Function BuildShapeJson(sh As Shape) As String
    Dim typeStr As String
    typeStr = ClassifyShapeType(sh)
    Dim result As String
    result = "{""shape_id"":" & CStr(sh.Id) & _
             ",""shape_name"":" & JsonStr(sh.Name) & _
             ",""type"":" & JsonStr(typeStr)
    If sh.HasTextFrame Then
        If sh.TextFrame.HasText Then
            result = result & ",""text"":" & JsonStr(sh.TextFrame.TextRange.Text)
        End If
    End If
    result = result & "}"
    BuildShapeJson = result
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

' JSON string: wrap in quotes and escape special chars
Private Function JsonStr(s As String) As String
    JsonStr = """" & JsonEscStr(s) & """"
End Function

Private Function JsonEscStr(s As String) As String
    Dim result As String
    result = s
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, Chr(13) & Chr(10), "\n")
    result = Replace(result, Chr(10), "\n")
    result = Replace(result, Chr(13), "\r")
    result = Replace(result, Chr(9), "\t")
    JsonEscStr = result
End Function
