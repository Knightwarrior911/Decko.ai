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
    ' Strip trailing paragraph terminator(s) from incoming value. The terminator
    ' is implicit in the paragraph-level TextRange; including a trailing \r
    ' causes PowerPoint to insert a NEW paragraph after this one, effectively
    ' duplicating the list. LLM-generated values frequently include trailing
    ' \r because the snapshot shows it, but it should NOT be in the value.
    p.Text = StripTrailingPara(value)
End Sub

' Strip trailing paragraph-terminator characters (Chr(13) and Chr(10)) from
' a value that is meant to live INSIDE one paragraph. Repeated stripping in
' case the LLM emits "\r\r" or "\r\n".
Public Function StripTrailingPara(s As String) As String
    Dim out As String: out = s
    Do While Len(out) > 0
        Dim last As String: last = Right(out, 1)
        If last = Chr(13) Or last = Chr(10) Then
            out = Left(out, Len(out) - 1)
        Else
            Exit Do
        End If
    Loop
    StripTrailingPara = out
End Function

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

Public Sub Do_set_bullet_style(slideNum As Long, shapeId As Long, _
                               paragraphIndex As Long, value As String)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 3001, "Do_set_bullet_style", "paragraph not found"
    Dim b As Object: Set b = p.ParagraphFormat.Bullet
    Select Case LCase(value)
        Case "none"
            b.Type = 0
        Case "number"
            b.Type = 2
        Case "letter"
            b.Type = 2
            b.Style = 16
        Case "disc", "bullet"
            b.Type = 1
            b.Character = 8226
        Case "square"
            b.Type = 1
            b.Character = 9632
        Case "dash"
            b.Type = 1
            b.Character = 8211
        Case Else
            Err.Raise vbObjectError + 3003, "Do_set_bullet_style", "unknown bullet style: " & value
    End Select
End Sub

Public Sub Do_set_indent_level(slideNum As Long, shapeId As Long, _
                               paragraphIndex As Long, value As Long)
    If value < 0 Or value > 4 Then Err.Raise vbObjectError + 3004, "Do_set_indent_level", "indent_level must be 0..4"
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 3001, "Do_set_indent_level", "paragraph not found"
    p.IndentLevel = value + 1
End Sub

Public Sub Do_set_paragraph_font_size(slideNum As Long, shapeId As Long, _
                                      paragraphIndex As Long, value As Long)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 3001, "Do_set_paragraph_font_size", "paragraph not found"
    p.Font.Size = value
End Sub

Public Sub Do_set_paragraph_font_color(slideNum As Long, shapeId As Long, _
                                       paragraphIndex As Long, hexValue As String)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 3001, "Do_set_paragraph_font_color", "paragraph not found"
    p.Font.Color.RGB = modActions.HexToRgb(hexValue)
End Sub

Public Sub Do_find_replace_text(scope As String, findText As String, replaceText As String)
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim slideNumFilter As Long: slideNumFilter = 0
    If LCase(Left(scope, 6)) = "slide:" Then
        slideNumFilter = CLng(Mid(scope, 7))
        If slideNumFilter < 1 Or slideNumFilter > pres.Slides.Count Then
            Err.Raise vbObjectError + 3005, "Do_find_replace_text", "slide_out_of_range"
        End If
    ElseIf LCase(scope) <> "deck" Then
        Err.Raise vbObjectError + 3006, "Do_find_replace_text", "scope must be 'deck' or 'slide:N'"
    End If

    Dim i As Long
    For i = 1 To pres.Slides.Count
        If slideNumFilter = 0 Or slideNumFilter = i Then
            Dim sh As Shape
            For Each sh In pres.Slides(i).Shapes
                ReplaceInShape sh, findText, replaceText
            Next sh
        End If
    Next i
End Sub

Private Sub ReplaceInShape(sh As Shape, findText As String, replaceText As String)
    On Error Resume Next
    If sh.HasTextFrame Then
        If sh.TextFrame.HasText Then
            ' Use TextRange.Find loop to preserve run/paragraph formatting:
            ' mutating a Find-returned range's .Text changes only the matched
            ' chars, keeping surrounding runs intact.
            ReplaceInTextRange sh.TextFrame.TextRange, findText, replaceText
        End If
    End If
    If sh.HasTable Then
        Dim r As Long, c As Long
        For r = 1 To sh.Table.Rows.Count
            For c = 1 To sh.Table.Columns.Count
                Dim cellSh As Shape: Set cellSh = sh.Table.Cell(r, c).Shape
                If cellSh.HasTextFrame Then
                    If cellSh.TextFrame.HasText Then
                        ReplaceInTextRange cellSh.TextFrame.TextRange, findText, replaceText
                    End If
                End If
            Next c
        Next r
    End If
    If sh.Type = msoGroup Then
        Dim child As Shape
        For Each child In sh.GroupItems
            ReplaceInShape child, findText, replaceText
        Next child
    End If
End Sub

' Format-preserving literal replace inside a TextRange. Uses TextRange.Find
' which returns a TextRange of the match; assigning .Text on that returned
' range mutates only the matched chars and inherits formatting from the
' first matched char, leaving surrounding runs untouched.
'
' Paragraph-terminator hygiene: if BOTH find and replace end in \r (or \n),
' strip those trailing terminators from BOTH before searching. PowerPoint's
' Find returns a span that does not reliably include the terminator, and
' assigning a replacement string that includes \r causes a new paragraph
' break to be inserted on top of the existing terminator. Net symptom:
' "Name\r" -> "Other\r" produces "Name" + blank paragraph + "Other".
' Stripping eliminates the duplicate \r without changing the visible result.
Public Sub ReplaceInTextRange(tr As TextRange, findText As String, replaceText As String)
    If Len(findText) = 0 Then Exit Sub
    Dim f As String: f = findText
    Dim r As String: r = replaceText
    Do While Len(f) > 0 And Len(r) > 0 _
        And (Right(f, 1) = Chr(13) Or Right(f, 1) = Chr(10)) _
        And (Right(r, 1) = Chr(13) Or Right(r, 1) = Chr(10))
        f = Left(f, Len(f) - 1)
        r = Left(r, Len(r) - 1)
    Loop
    If Len(f) = 0 Then Exit Sub
    Dim guard As Long: guard = 0
    Dim found As TextRange
    Dim startPos As Long: startPos = 1
    Do
        Set found = Nothing
        On Error Resume Next
        Set found = tr.Find(FindWhat:=f, After:=startPos - 1)
        On Error GoTo 0
        If found Is Nothing Then Exit Do
        Dim matchStart As Long: matchStart = found.Start
        found.Text = r
        startPos = matchStart + Len(r)
        guard = guard + 1
        If guard > 10000 Then Exit Do  ' safety
    Loop
End Sub

Public Sub Do_set_paragraph_alignment(slideNum As Long, shapeId As Long, _
                                      paragraphIndex As Long, align As String)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 6001, "Do_set_paragraph_alignment", "paragraph not found"
    Select Case LCase(align)
        Case "left":    p.ParagraphFormat.Alignment = ppAlignLeft
        Case "center":  p.ParagraphFormat.Alignment = ppAlignCenter
        Case "right":   p.ParagraphFormat.Alignment = ppAlignRight
        Case "justify": p.ParagraphFormat.Alignment = ppAlignJustify
        Case Else: Err.Raise vbObjectError + 6001, "Do_set_paragraph_alignment", "bad align: " & align
    End Select
End Sub

Public Sub Do_set_paragraph_line_spacing(slideNum As Long, shapeId As Long, _
                                         paragraphIndex As Long, multiple As Double)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 6002, "Do_set_paragraph_line_spacing", "paragraph not found"
    If multiple <= 0 Then Err.Raise vbObjectError + 6002, "Do_set_paragraph_line_spacing", "multiple must be > 0"
    p.ParagraphFormat.LineRuleWithin = msoTrue
    p.ParagraphFormat.SpaceWithin = multiple
End Sub

Public Sub Do_set_text_vertical_align(slideNum As Long, shapeId As Long, anchor As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 6003, "Do_set_text_vertical_align", "shape not found"
    If Not sh.HasTextFrame Then Err.Raise vbObjectError + 6003, "Do_set_text_vertical_align", "no text frame"
    Select Case LCase(anchor)
        Case "top":    sh.TextFrame.VerticalAnchor = msoAnchorTop
        Case "middle": sh.TextFrame.VerticalAnchor = msoAnchorMiddle
        Case "bottom": sh.TextFrame.VerticalAnchor = msoAnchorBottom
        Case Else: Err.Raise vbObjectError + 6003, "Do_set_text_vertical_align", "bad anchor: " & anchor
    End Select
End Sub

Public Sub Do_set_text_margin(slideNum As Long, shapeId As Long, _
                              leftPt As Double, rightPt As Double, _
                              topPt As Double, bottomPt As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 6004, "Do_set_text_margin", "shape not found"
    If Not sh.HasTextFrame Then Err.Raise vbObjectError + 6004, "Do_set_text_margin", "no text frame"
    If leftPt < 0 Or rightPt < 0 Or topPt < 0 Or bottomPt < 0 Then _
        Err.Raise vbObjectError + 6004, "Do_set_text_margin", "margin must be >= 0"
    sh.TextFrame.MarginLeft   = leftPt
    sh.TextFrame.MarginRight  = rightPt
    sh.TextFrame.MarginTop    = topPt
    sh.TextFrame.MarginBottom = bottomPt
End Sub
