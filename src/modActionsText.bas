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

    ' Empty text frame: nothing to preserve, just set the text.
    If n = 0 Or Len(tr.Text) = 0 Then
        tr.Text = value
        Exit Sub
    End If

    ' Prior implementation rebuilt the entire text frame via tr.Text = newText,
    ' which discarded per-run formatting (bold/italic/size/color) on every
    ' existing paragraph. TextRange.InsertBefore / InsertAfter mutate only the
    ' boundary and leave surrounding runs untouched.
    If afterParagraphIndex < 0 Then
        tr.Paragraphs(1).InsertBefore value & Chr(13)
        Exit Sub
    End If

    If afterParagraphIndex >= n Then
        ' Append after the last paragraph. If the existing text already ends
        ' with a terminator, avoid doubling it (would create a blank line).
        If Right(tr.Text, 1) = Chr(13) Then
            tr.InsertAfter value
        Else
            tr.InsertAfter Chr(13) & value
        End If
        Exit Sub
    End If

    ' Insert between paragraphs. tr.Paragraphs(K).InsertAfter places the new
    ' text at K's end (just past K's terminator) = start of K+1, so the result
    ' is paraK\rvalue\rparaK+1, with every other paragraph's runs preserved.
    tr.Paragraphs(afterParagraphIndex + 1).InsertAfter value & Chr(13)
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

Public Sub Do_find_replace_text(scope As String, findText As String, replaceText As String, _
                                 Optional caseSensitive As Boolean = False, _
                                 Optional wholeWord As Boolean = False, _
                                 Optional includeNotes As Boolean = False)
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
                ReplaceInShape sh, findText, replaceText, caseSensitive, wholeWord
            Next sh
            If includeNotes Then
                Dim notesSh As Shape
                For Each notesSh In pres.Slides(i).NotesPage.Shapes
                    If notesSh.HasTextFrame Then
                        If notesSh.TextFrame.HasText Then
                            ReplaceInTextRange notesSh.TextFrame.TextRange, _
                                findText, replaceText, caseSensitive, wholeWord
                        End If
                    End If
                Next notesSh
            End If
        End If
    Next i
End Sub

Private Sub ReplaceInShape(sh As Shape, findText As String, replaceText As String, _
                            Optional caseSensitive As Boolean = False, _
                            Optional wholeWord As Boolean = False)
    On Error Resume Next
    If sh.HasTextFrame Then
        If sh.TextFrame.HasText Then
            ReplaceInTextRange sh.TextFrame.TextRange, findText, replaceText, caseSensitive, wholeWord
        End If
    End If
    If sh.HasTable Then
        Dim r As Long, c As Long
        For r = 1 To sh.Table.Rows.Count
            For c = 1 To sh.Table.Columns.Count
                Dim cellSh As Shape: Set cellSh = sh.Table.Cell(r, c).Shape
                If cellSh.HasTextFrame Then
                    If cellSh.TextFrame.HasText Then
                        ReplaceInTextRange cellSh.TextFrame.TextRange, findText, replaceText, caseSensitive, wholeWord
                    End If
                End If
            Next c
        Next r
    End If
    If sh.Type = msoGroup Then
        Dim child As Shape
        For Each child In sh.GroupItems
            ReplaceInShape child, findText, replaceText, caseSensitive, wholeWord
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
Public Sub ReplaceInTextRange(tr As TextRange, findText As String, replaceText As String, _
                                Optional caseSensitive As Boolean = False, _
                                Optional wholeWord As Boolean = False)
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
        Set found = tr.Find(FindWhat:=f, After:=startPos - 1, _
                            MatchCase:=caseSensitive, WholeWords:=wholeWord)
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

' Set TextFrame auto-fit mode. Modes:
'   "none"   - no auto-sizing; text overflows out of the shape (default)
'   "shrink" - shrink text font to fit when overflow (banker-deck style)
'   "resize" - resize the shape to fit the text
' Uses TextFrame2 (Office 2007+) which exposes msoAutoSize enum:
'   0 = msoAutoSizeNone, 1 = msoAutoSizeShapeToFitText, 2 = msoAutoSizeTextToFitShape
Public Sub Do_set_text_autofit(slideNum As Long, shapeId As Long, mode As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 6005, "Do_set_text_autofit", "shape not found"
    If Not sh.HasTextFrame Then Err.Raise vbObjectError + 6005, "Do_set_text_autofit", "no text frame"
    Dim m As Long
    Select Case LCase(mode)
        Case "none":   m = 0
        Case "shrink": m = 2
        Case "resize": m = 1
        Case Else: Err.Raise vbObjectError + 6005, "Do_set_text_autofit", "mode must be none/shrink/resize"
    End Select
    On Error Resume Next
    sh.TextFrame2.AutoSize = m
    On Error GoTo 0
End Sub

' Sweep all text shapes in scope and enable shrink-on-overflow. Call this
' after a batch of text-mutating actions to ensure no shape's text spills
' outside its frame. Skips title shapes by default (titles should not shrink).
'   scope = "deck" or "slide:N"
'   include_titles = "true" to include title shapes; default skips them
Public Sub Do_enable_text_shrink_for_overflow(scope As String, includeTitles As String)
    Dim startSlide As Long, endSlide As Long
    If LCase(scope) = "deck" Then
        startSlide = 1
        endSlide = ActivePresentation.Slides.Count
    ElseIf LCase(Left(scope, 6)) = "slide:" Then
        startSlide = CLng(Mid(scope, 7))
        endSlide = startSlide
    Else
        Err.Raise vbObjectError + 6006, "Do_enable_text_shrink_for_overflow", _
            "scope must be 'deck' or 'slide:N'"
    End If
    Dim incl As Boolean: incl = (LCase(includeTitles) = "true")
    Dim s As Long
    For s = startSlide To endSlide
        Dim sl As Slide: Set sl = ActivePresentation.Slides(s)
        Dim sh As Shape
        For Each sh In sl.Shapes
            EnableShrinkOnShape sh, incl
        Next sh
    Next s
End Sub

Private Sub EnableShrinkOnShape(sh As Shape, includeTitles As Boolean)
    On Error Resume Next
    If Not includeTitles Then
        If sh.Type = msoPlaceholder Then
            If sh.PlaceholderFormat.Type = ppPlaceholderTitle Or _
               sh.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
                Exit Sub
            End If
        End If
    End If
    If sh.HasTextFrame Then
        ' Only shrink if text actually overflows — avoids crushing badge circles
        ' that fit fine but get recalculated smaller by TextToFitShape AutoSize
        If sh.TextFrame2.AutoSize = 0 Then
            Dim bh As Single: bh = sh.TextFrame.TextRange.BoundHeight
            Dim bw As Single: bw = sh.TextFrame.TextRange.BoundWidth
            If bh > sh.Height + 2 Or bw > sh.Width + 2 Then
                sh.TextFrame2.AutoSize = 2  ' msoAutoSizeTextToFitShape
            End If
        End If
    End If
    If sh.Type = msoGroup Then
        Dim child As Shape
        For Each child In sh.GroupItems
            EnableShrinkOnShape child, includeTitles
        Next child
    End If
    On Error GoTo 0
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

' --- Paragraph-level granular formatting --------------------------------------
' These mirror the run-level toggles but apply to an entire paragraph at once,
' which is more convenient when the paragraph has only one run anyway and
' avoids hunting for run_index.

Public Sub Do_set_paragraph_bold(slideNum As Long, shapeId As Long, _
                                  paragraphIndex As Long, value As Boolean)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 6010, "Do_set_paragraph_bold", "paragraph not found"
    p.Font.Bold = IIf(value, msoTrue, msoFalse)
End Sub

Public Sub Do_set_paragraph_italic(slideNum As Long, shapeId As Long, _
                                    paragraphIndex As Long, value As Boolean)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 6011, "Do_set_paragraph_italic", "paragraph not found"
    p.Font.Italic = IIf(value, msoTrue, msoFalse)
End Sub

Public Sub Do_set_paragraph_underline(slideNum As Long, shapeId As Long, _
                                       paragraphIndex As Long, value As Boolean)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 6012, "Do_set_paragraph_underline", "paragraph not found"
    p.Font.Underline = IIf(value, msoTrue, msoFalse)
End Sub

Public Sub Do_set_paragraph_font_name(slideNum As Long, shapeId As Long, _
                                       paragraphIndex As Long, fontName As String)
    If Len(Trim(fontName)) = 0 Then Err.Raise vbObjectError + 6013, "Do_set_paragraph_font_name", "font name empty"
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 6013, "Do_set_paragraph_font_name", "paragraph not found"
    p.Font.Name = fontName
End Sub

' Space BEFORE a paragraph in points. Useful for spacing bullets without
' affecting their line-height. Negative values are clamped to 0.
Public Sub Do_set_paragraph_space_before(slideNum As Long, shapeId As Long, _
                                          paragraphIndex As Long, ptValue As Double)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 6014, "Do_set_paragraph_space_before", "paragraph not found"
    If ptValue < 0 Then ptValue = 0
    p.ParagraphFormat.SpaceBefore = ptValue
End Sub

Public Sub Do_set_paragraph_space_after(slideNum As Long, shapeId As Long, _
                                         paragraphIndex As Long, ptValue As Double)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 6015, "Do_set_paragraph_space_after", "paragraph not found"
    If ptValue < 0 Then ptValue = 0
    p.ParagraphFormat.SpaceAfter = ptValue
End Sub

' Numbered-list start number. For paragraphs with bullet type "number" or "letter",
' sets the starting count (default 1).
Public Sub Do_set_bullet_start_number(slideNum As Long, shapeId As Long, _
                                       paragraphIndex As Long, value As Long)
    If value < 1 Then Err.Raise vbObjectError + 6020, "Do_set_bullet_start_number", "value must be >= 1"
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 6020, "Do_set_bullet_start_number", "paragraph not found"
    On Error Resume Next
    p.ParagraphFormat.Bullet.StartValue = value
    On Error GoTo 0
End Sub

' Reset a paragraph's character formatting to defaults. Useful before applying
' a new style to an existing paragraph (e.g. demoting a heading back to a bullet
' before re-styling). Resets bold/italic/underline/strikethrough/baseline offset
' but PRESERVES font size and color (use set_paragraph_font_size/color to wipe those).
Public Sub Do_clear_paragraph_formatting(slideNum As Long, shapeId As Long, _
                                          paragraphIndex As Long)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 6016, "Do_clear_paragraph_formatting", "paragraph not found"
    On Error Resume Next
    p.Font.Bold = msoFalse
    p.Font.Italic = msoFalse
    p.Font.Underline = msoFalse
    p.Font.BaselineOffset = 0
    ' Strikethrough lives on Font2; clear via TextFrame2 path
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If Not sh Is Nothing Then
        Dim tr2 As Object: Set tr2 = sh.TextFrame2.TextRange
        If paragraphIndex + 1 <= tr2.Paragraphs.Count Then
            tr2.Paragraphs(paragraphIndex + 1).Font.Strikethrough = msoFalse
        End If
    End If
    On Error GoTo 0
End Sub
