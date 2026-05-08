Attribute VB_Name = "modActionsText"
Option Explicit

' Helper used by every text-action Sub.
Public Function FindParagraph(slideNum As Long, shapeId As Long, zeroIdx As Long) As TextRange
    Set FindParagraph = Nothing
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Exit Function
    If Not sh.HasTextFrame Then Exit Function
    Dim n As Long: n = sh.TextFrame.TextRange.Paragraphs().Count
    If zeroIdx < 0 Or zeroIdx >= n Then Exit Function
    Set FindParagraph = sh.TextFrame.TextRange.Paragraphs(zeroIdx + 1)
End Function

Public Sub Do_set_paragraph_text(slideNum As Long, shapeId As Long, _
                                 paragraphIndex As Long, value As String)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 3001, "Do_set_paragraph_text", "paragraph not found"
    p.Text = value
End Sub

Public Sub Do_add_paragraph(slideNum As Long, shapeId As Long, _
                            afterParagraphIndex As Long, value As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 3001, "Do_add_paragraph", "shape not found"
    If Not sh.HasTextFrame Then Err.Raise vbObjectError + 3002, "Do_add_paragraph", "no text frame"

    Dim tr As TextRange: Set tr = sh.TextFrame.TextRange
    Dim n As Long: n = tr.Paragraphs().Count

    If afterParagraphIndex < 0 Then
        tr.Text = value & Chr(13) & tr.Text
        Exit Sub
    End If

    If afterParagraphIndex >= n Then
        tr.Text = tr.Text & Chr(13) & value
        Exit Sub
    End If

    ' Build new text by splitting at paragraph boundary
    ' Strip trailing Chr(13) from each paragraph's .Text before rebuilding
    Dim parts() As String
    Dim i As Long
    ReDim parts(n - 1)
    For i = 1 To n
        Dim pText As String
        pText = tr.Paragraphs(i).Text
        If Right(pText, 1) = Chr(13) Then pText = Left(pText, Len(pText) - 1)
        parts(i - 1) = pText
    Next i

    Dim newText As String
    newText = ""
    For i = 0 To afterParagraphIndex
        If i > 0 Then newText = newText & Chr(13)
        newText = newText & parts(i)
    Next i
    newText = newText & Chr(13) & value
    For i = afterParagraphIndex + 1 To n - 1
        newText = newText & Chr(13) & parts(i)
    Next i

    tr.Text = newText
End Sub

Public Sub Do_delete_paragraph(slideNum As Long, shapeId As Long, paragraphIndex As Long)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 3001, "Do_delete_paragraph", "paragraph not found"
    p.Delete
End Sub
