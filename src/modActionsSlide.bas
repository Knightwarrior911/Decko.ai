Attribute VB_Name = "modActionsSlide"
Option Explicit

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
