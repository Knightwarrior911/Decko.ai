Attribute VB_Name = "modActionsRun"
Option Explicit

' Resolve a (slide, shape, paragraph_idx, run_idx) tuple to a TextRange that
' represents exactly one run. Returns Nothing on any out-of-range input.
Public Function FindRun(slideNum As Long, shapeId As Long, _
                        paragraphIndex As Long, runIndex As Long) As TextRange
    Set FindRun = Nothing
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Exit Function
    If Not sh.HasTextFrame Then Exit Function
    Dim tr As TextRange: Set tr = sh.TextFrame.TextRange
    If paragraphIndex < 0 Or paragraphIndex >= tr.Paragraphs().Count Then Exit Function
    Dim p As TextRange: Set p = tr.Paragraphs(paragraphIndex + 1)
    If runIndex < 0 Or runIndex >= p.Runs.Count Then Exit Function
    Set FindRun = p.Runs(runIndex + 1)
End Function

Public Sub Do_set_run_bold(slideNum As Long, shapeId As Long, _
                           paragraphIndex As Long, runIndex As Long, value As Boolean)
    Dim r As TextRange: Set r = FindRun(slideNum, shapeId, paragraphIndex, runIndex)
    If r Is Nothing Then Err.Raise vbObjectError + 5001, "Do_set_run_bold", "run not found"
    r.Font.Bold = IIf(value, msoTrue, msoFalse)
End Sub

Public Sub Do_set_run_italic(slideNum As Long, shapeId As Long, _
                             paragraphIndex As Long, runIndex As Long, value As Boolean)
    Dim r As TextRange: Set r = FindRun(slideNum, shapeId, paragraphIndex, runIndex)
    If r Is Nothing Then Err.Raise vbObjectError + 5002, "Do_set_run_italic", "run not found"
    r.Font.Italic = IIf(value, msoTrue, msoFalse)
End Sub

Public Sub Do_set_run_underline(slideNum As Long, shapeId As Long, _
                                paragraphIndex As Long, runIndex As Long, value As Boolean)
    Dim r As TextRange: Set r = FindRun(slideNum, shapeId, paragraphIndex, runIndex)
    If r Is Nothing Then Err.Raise vbObjectError + 5003, "Do_set_run_underline", "run not found"
    r.Font.Underline = IIf(value, msoTrue, msoFalse)
End Sub

Public Sub Do_set_run_strikethrough(slideNum As Long, shapeId As Long, _
                                     paragraphIndex As Long, runIndex As Long, value As Boolean)
    ' Strikethrough requires Font2 (TextRange2 path) on PowerPoint
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 5004, "Do_set_run_strikethrough", "shape not found"
    If Not sh.HasTextFrame Then Err.Raise vbObjectError + 5004, "Do_set_run_strikethrough", "no text frame"
    Dim tr2 As Object: Set tr2 = sh.TextFrame2.TextRange
    If paragraphIndex + 1 > tr2.Paragraphs.Count Then
        Err.Raise vbObjectError + 5004, "Do_set_run_strikethrough", "paragraph_index out of range"
    End If
    Dim p2 As Object: Set p2 = tr2.Paragraphs(paragraphIndex + 1)
    If runIndex + 1 > p2.Runs.Count Then
        Err.Raise vbObjectError + 5004, "Do_set_run_strikethrough", "run_index out of range"
    End If
    Dim run2 As Object: Set run2 = p2.Runs(runIndex + 1)
    run2.Font.Strikethrough = IIf(value, msoTrue, msoFalse)
End Sub

Public Sub Do_set_run_subscript(slideNum As Long, shapeId As Long, _
                                paragraphIndex As Long, runIndex As Long, value As Boolean)
    Dim r As TextRange: Set r = FindRun(slideNum, shapeId, paragraphIndex, runIndex)
    If r Is Nothing Then Err.Raise vbObjectError + 5005, "Do_set_run_subscript", "run not found"
    If value Then
        r.Font.BaselineOffset = -0.25
    Else
        r.Font.BaselineOffset = 0
    End If
End Sub

Public Sub Do_set_run_superscript(slideNum As Long, shapeId As Long, _
                                  paragraphIndex As Long, runIndex As Long, value As Boolean)
    Dim r As TextRange: Set r = FindRun(slideNum, shapeId, paragraphIndex, runIndex)
    If r Is Nothing Then Err.Raise vbObjectError + 5006, "Do_set_run_superscript", "run not found"
    If value Then
        r.Font.BaselineOffset = 0.3
    Else
        r.Font.BaselineOffset = 0
    End If
End Sub

Public Sub Do_set_run_font_color(slideNum As Long, shapeId As Long, _
                                 paragraphIndex As Long, runIndex As Long, hexValue As String)
    Dim r As TextRange: Set r = FindRun(slideNum, shapeId, paragraphIndex, runIndex)
    If r Is Nothing Then Err.Raise vbObjectError + 5007, "Do_set_run_font_color", "run not found"
    r.Font.Color.RGB = modActions.HexToRgb(hexValue)
End Sub

Public Sub Do_set_run_font_size(slideNum As Long, shapeId As Long, _
                                paragraphIndex As Long, runIndex As Long, ptValue As Long)
    Dim r As TextRange: Set r = FindRun(slideNum, shapeId, paragraphIndex, runIndex)
    If r Is Nothing Then Err.Raise vbObjectError + 5008, "Do_set_run_font_size", "run not found"
    If ptValue <= 0 Then Err.Raise vbObjectError + 5008, "Do_set_run_font_size", "size must be > 0"
    r.Font.Size = ptValue
End Sub

Public Sub Do_set_run_font_name(slideNum As Long, shapeId As Long, _
                                paragraphIndex As Long, runIndex As Long, fontName As String)
    Dim r As TextRange: Set r = FindRun(slideNum, shapeId, paragraphIndex, runIndex)
    If r Is Nothing Then Err.Raise vbObjectError + 5009, "Do_set_run_font_name", "run not found"
    If Len(Trim(fontName)) = 0 Then Err.Raise vbObjectError + 5009, "Do_set_run_font_name", "font name empty"
    r.Font.Name = fontName
End Sub

Public Sub Do_set_run_text(slideNum As Long, shapeId As Long, _
                           paragraphIndex As Long, runIndex As Long, value As String)
    Dim r As TextRange: Set r = FindRun(slideNum, shapeId, paragraphIndex, runIndex)
    If r Is Nothing Then Err.Raise vbObjectError + 5010, "Do_set_run_text", "run not found"
    ' Strip trailing paragraph terminators - a run lives inside one paragraph,
    ' so a trailing \r in the value would insert a new paragraph break.
    r.Text = modActionsText.StripTrailingPara(value)
End Sub

Public Sub Do_set_run_hyperlink(slideNum As Long, shapeId As Long, _
                                paragraphIndex As Long, runIndex As Long, url As String)
    Dim r As TextRange: Set r = FindRun(slideNum, shapeId, paragraphIndex, runIndex)
    If r Is Nothing Then Err.Raise vbObjectError + 5011, "Do_set_run_hyperlink", "run not found"
    ' 1 = ppMouseClick. To CLEAR a hyperlink, setting Address = "" is unreliable;
    ' the canonical way is to set Action = ppActionNone (0) which removes the
    ' attached action entirely. To SET a hyperlink, set Address (Action implicitly
    ' becomes ppActionHyperlink).
    If Len(url) = 0 Then
        r.ActionSettings(1).Action = 0   ' ppActionNone
        r.ActionSettings(1).Hyperlink.Address = ""
    Else
        r.ActionSettings(1).Hyperlink.Address = url
    End If
End Sub

' Character spacing / kerning for a run. value in points (0 = default,
' positive = wider letters, negative = tighter). Lives on Font2.
Public Sub Do_set_run_kerning(slideNum As Long, shapeId As Long, _
                               paragraphIndex As Long, runIndex As Long, ptValue As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 5013, "Do_set_run_kerning", "shape not found"
    If Not sh.HasTextFrame Then Err.Raise vbObjectError + 5013, "Do_set_run_kerning", "no text frame"
    Dim tr2 As Object: Set tr2 = sh.TextFrame2.TextRange
    If paragraphIndex + 1 > tr2.Paragraphs.Count Then
        Err.Raise vbObjectError + 5013, "Do_set_run_kerning", "paragraph_index out of range"
    End If
    Dim p2 As Object: Set p2 = tr2.Paragraphs(paragraphIndex + 1)
    If runIndex + 1 > p2.Runs.Count Then
        Err.Raise vbObjectError + 5013, "Do_set_run_kerning", "run_index out of range"
    End If
    On Error Resume Next
    p2.Runs(runIndex + 1).Font.Spacing = ptValue
    On Error GoTo 0
End Sub

' Custom baseline offset for a run. value: -1.0 to +1.0 as fraction of font
' height (set_run_superscript is fixed +0.3; this lets you fine-tune).
Public Sub Do_set_run_baseline_offset(slideNum As Long, shapeId As Long, _
                                       paragraphIndex As Long, runIndex As Long, value As Double)
    Dim r As TextRange: Set r = FindRun(slideNum, shapeId, paragraphIndex, runIndex)
    If r Is Nothing Then Err.Raise vbObjectError + 5014, "Do_set_run_baseline_offset", "run not found"
    If value < -1 Or value > 1 Then Err.Raise vbObjectError + 5014, "Do_set_run_baseline_offset", "value must be -1..1"
    r.Font.BaselineOffset = value
End Sub

' Text highlight (background fill behind characters). Lives on Font2/TextFrame2.
' Pass "" to clear the highlight. Caller-side range is one run, mirroring all
' other run-level actions.
Public Sub Do_set_run_highlight(slideNum As Long, shapeId As Long, _
                                paragraphIndex As Long, runIndex As Long, hexValue As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 5012, "Do_set_run_highlight", "shape not found"
    If Not sh.HasTextFrame Then Err.Raise vbObjectError + 5012, "Do_set_run_highlight", "no text frame"
    Dim tr2 As Object: Set tr2 = sh.TextFrame2.TextRange
    If paragraphIndex + 1 > tr2.Paragraphs.Count Then
        Err.Raise vbObjectError + 5012, "Do_set_run_highlight", "paragraph_index out of range"
    End If
    Dim p2 As Object: Set p2 = tr2.Paragraphs(paragraphIndex + 1)
    If runIndex + 1 > p2.Runs.Count Then
        Err.Raise vbObjectError + 5012, "Do_set_run_highlight", "run_index out of range"
    End If
    Dim run2 As Object: Set run2 = p2.Runs(runIndex + 1)
    On Error Resume Next
    If Len(hexValue) = 0 Then
        run2.Font.Highlight.RGB = RGB(255, 255, 255)
        run2.Font.Highlight.Visible = msoFalse
    Else
        run2.Font.Highlight.RGB = modActions.HexToRgb(hexValue)
        run2.Font.Highlight.Visible = msoTrue
    End If
    On Error GoTo 0
End Sub
