Attribute VB_Name = "modActionsSlide"
Option Explicit

Public Sub Do_set_slide_background_color(slideNum As Long, hexColor As String)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 7050, "Do_set_slide_background_color", "slide_out_of_range"
    End If
    Dim sl As Slide: Set sl = pres.Slides(slideNum)
    sl.FollowMasterBackground = msoFalse
    sl.Background.Fill.Visible = msoTrue
    sl.Background.Fill.Solid
    sl.Background.Fill.ForeColor.RGB = modActions.HexToRgb(hexColor)
End Sub

Public Sub Do_insert_slide_number(slideNum As Long, _
                                   leftPt As Single, topPt As Single, _
                                   widthPt As Single, heightPt As Single, _
                                   Optional refName As String = "", _
                                   Optional fontColor As String = "", _
                                   Optional fontSize As Long = 0)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 7051, "Do_insert_slide_number", "slide_out_of_range"
    End If
    Dim sh As Shape
    Set sh = pres.Slides(slideNum).Shapes.AddTextbox( _
        msoTextOrientationHorizontal, leftPt, topPt, widthPt, heightPt)
    If Len(refName) > 0 Then sh.Name = refName
    sh.TextFrame.AutoSize = ppAutoSizeNone
    sh.TextFrame.TextRange.InsertSlideNumber
    With sh.TextFrame.TextRange.Font
        If Len(fontColor) > 0 Then .Color.RGB = modActions.HexToRgb(fontColor)
        If fontSize > 0 Then .Size = fontSize
    End With
End Sub

Public Sub Do_move_slide(fromIdx As Long, toIdx As Long)
    Dim pres As Presentation: Set pres = ActivePresentation
    If fromIdx < 1 Or fromIdx > pres.Slides.Count Then
        Err.Raise vbObjectError + 7001, "Do_move_slide", "from out of range"
    End If
    If toIdx < 1 Or toIdx > pres.Slides.Count Then
        Err.Raise vbObjectError + 7002, "Do_move_slide", "to out of range"
    End If
    pres.Slides(fromIdx).MoveTo toIdx
End Sub

Public Sub Do_extract_slides(slideIndices As Variant, outputPath As String)
    Dim ids() As Long
    Dim cnt As Long: cnt = modActionsLayout.NormalizeIdsArray(slideIndices, ids)
    If cnt < 1 Then Err.Raise vbObjectError + 7003, "Do_extract_slides", "no slides specified"

    Dim src As Presentation: Set src = ActivePresentation
    Dim outPres As Presentation
    Set outPres = Application.Presentations.Add(WithWindow:=msoFalse)
    Dim i As Long
    For i = 0 To cnt - 1
        Dim n As Long: n = ids(i)
        If n < 1 Or n > src.Slides.Count Then
            Err.Raise vbObjectError + 7004, "Do_extract_slides", "slide index out of range: " & n
        End If
        src.Slides(n).Copy
        outPres.Slides.Paste
    Next i

    On Error Resume Next
    Do While outPres.Slides.Count > cnt
        outPres.Slides(1).Delete
    Loop
    Err.Clear
    On Error GoTo 0

    outPres.SaveAs outputPath
    outPres.Close
End Sub

Public Sub Do_import_slides_from_deck(sourcePath As String, slideIndices As Variant, _
                                      targetPosition As Long)
    Dim ids() As Long
    Dim cnt As Long: cnt = modActionsLayout.NormalizeIdsArray(slideIndices, ids)
    If cnt < 1 Then Err.Raise vbObjectError + 7005, "Do_import_slides_from_deck", "no slide indices"
    If Not FileExistsLocal(sourcePath) Then
        Err.Raise vbObjectError + 7006, "Do_import_slides_from_deck", "source_not_found: " & sourcePath
    End If

    Dim pres As Presentation: Set pres = ActivePresentation
    If targetPosition < 1 Then targetPosition = 1
    If targetPosition > pres.Slides.Count + 1 Then targetPosition = pres.Slides.Count + 1

    Dim i As Long
    Dim insertedSoFar As Long: insertedSoFar = 0
    i = 0
    Do While i <= cnt - 1
        Dim startIdx As Long: startIdx = ids(i)
        Dim endIdx As Long: endIdx = startIdx
        Do While i + 1 <= cnt - 1
            If ids(i + 1) = ids(i) + 1 Then
                i = i + 1
                endIdx = ids(i)
            Else
                Exit Do
            End If
        Loop
        pres.Slides.InsertFromFile sourcePath, _
                                   targetPosition - 1 + insertedSoFar, _
                                   startIdx, endIdx
        insertedSoFar = insertedSoFar + (endIdx - startIdx + 1)
        i = i + 1
    Loop
End Sub

Private Function FileExistsLocal(p As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExistsLocal = fso.FileExists(p)
End Function
