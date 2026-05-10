Attribute VB_Name = "modActionsDeck"
Option Explicit

' Deck-wide actions: regex find/replace, font swap, palette recolor, theme,
' slide size, theme font, bulk-insert across slides, layout application.

Public Sub Do_find_replace_regex(scope As String, pattern As String, replacement As String)
    If Len(pattern) = 0 Then Err.Raise vbObjectError + 7001, "Do_find_replace_regex", "pattern empty"
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    re.pattern = pattern

    Dim startSlide As Long, endSlide As Long
    ParseScopeRange scope, startSlide, endSlide

    Dim s As Long
    For s = startSlide To endSlide
        Dim sl As Slide: Set sl = ActivePresentation.Slides(s)
        Dim sh As Shape
        For Each sh In sl.Shapes
            ApplyRegexToShape sh, re, replacement
        Next sh
    Next s
End Sub

Private Sub ApplyRegexToShape(sh As Shape, re As Object, replacement As String)
    On Error Resume Next
    If sh.HasTextFrame Then
        If sh.TextFrame.HasText Then
            ' Per-paragraph regex match: process matches in REVERSE order so
            ' length-changing replacements don't shift earlier offsets. Each
            ' replacement is written via tr.Characters(start, len).Text =
            ' replacement, which preserves the formatting of the matched
            ' span's first character and leaves surrounding runs intact.
            Dim p As Long
            For p = 1 To sh.TextFrame.TextRange.Paragraphs.Count
                ReplaceRegexInParagraph sh.TextFrame.TextRange.Paragraphs(p), re, replacement
            Next p
        End If
    End If
    If sh.Type = msoGroup Then
        Dim child As Shape
        For Each child In sh.GroupItems
            ApplyRegexToShape child, re, replacement
        Next child
    End If
    On Error GoTo 0
End Sub

Private Sub ReplaceRegexInParagraph(para As TextRange, re As Object, replacement As String)
    On Error Resume Next
    Dim text As String: text = para.Text
    ' Strip trailing paragraph mark for matching.
    Dim hasCR As Boolean: hasCR = (Len(text) > 0 And Right(text, 1) = Chr(13))
    Dim matchSrc As String: matchSrc = text
    If hasCR Then matchSrc = Left(text, Len(text) - 1)
    Dim matches As Object: Set matches = re.Execute(matchSrc)
    If matches Is Nothing Then Exit Sub
    If matches.Count = 0 Then Exit Sub
    ' Walk matches in reverse to keep earlier offsets stable.
    Dim i As Long
    For i = matches.Count - 1 To 0 Step -1
        Dim m As Object: Set m = matches(i)
        Dim startOneBased As Long: startOneBased = m.FirstIndex + 1
        Dim matchLen As Long: matchLen = m.Length
        Dim repl As String
        repl = re.Replace(m.Value, replacement)
        ' Strip trailing paragraph terminators from replacement - they would
        ' insert spurious paragraph breaks inside this paragraph.
        repl = modActionsText.StripTrailingPara(repl)
        para.Characters(startOneBased, matchLen).Text = repl
    Next i
    On Error GoTo 0
End Sub

Private Sub ParseScopeRange(scope As String, ByRef startSlide As Long, ByRef endSlide As Long)
    If LCase(scope) = "deck" Then
        startSlide = 1
        endSlide = ActivePresentation.Slides.Count
    ElseIf LCase(Left(scope, 6)) = "slide:" Then
        Dim n As Long: n = CLng(Mid(scope, 7))
        startSlide = n
        endSlide = n
    Else
        Err.Raise vbObjectError + 7001, "ParseScopeRange", "scope must be 'deck' or 'slide:N'"
    End If
End Sub

Public Sub Do_swap_font_deck_wide(fromName As String, toName As String)
    If Len(fromName) = 0 Or Len(toName) = 0 Then _
        Err.Raise vbObjectError + 7002, "Do_swap_font_deck_wide", "font names must be non-empty"
    Dim sl As Slide
    For Each sl In ActivePresentation.Slides
        Dim sh As Shape
        For Each sh In sl.Shapes
            SwapFontInShape sh, fromName, toName
        Next sh
    Next sl
End Sub

Private Sub SwapFontInShape(sh As Shape, fromName As String, toName As String)
    On Error Resume Next
    If sh.HasTextFrame Then
        Dim para As TextRange
        Dim p As Long
        For p = 1 To sh.TextFrame.TextRange.Paragraphs.Count
            Set para = sh.TextFrame.TextRange.Paragraphs(p)
            Dim r As Long
            For r = 1 To para.Runs.Count
                If para.Runs(r).Font.Name = fromName Then
                    para.Runs(r).Font.Name = toName
                End If
            Next r
        Next p
    End If
    If sh.Type = msoGroup Then
        Dim child As Shape
        For Each child In sh.GroupItems
            SwapFontInShape child, fromName, toName
        Next child
    End If
    On Error GoTo 0
End Sub

Public Sub Do_recolor_palette_deck_wide(fromHex As String, toHex As String, target As String)
    Dim t As String: t = LCase(target)
    If t <> "fill" And t <> "font" And t <> "both" Then _
        Err.Raise vbObjectError + 7003, "Do_recolor_palette_deck_wide", "target must be fill/font/both"
    Dim fromRgb As Long: fromRgb = modActions.HexToRgb(fromHex)
    Dim toRgb As Long: toRgb = modActions.HexToRgb(toHex)
    Dim sl As Slide
    For Each sl In ActivePresentation.Slides
        Dim sh As Shape
        For Each sh In sl.Shapes
            RecolorInShape sh, fromRgb, toRgb, t
        Next sh
    Next sl
End Sub

Private Sub RecolorInShape(sh As Shape, fromRgb As Long, toRgb As Long, target As String)
    On Error Resume Next
    If target = "fill" Or target = "both" Then
        If sh.Fill.Visible = msoTrue And sh.Fill.Type = msoFillSolid Then
            If sh.Fill.ForeColor.RGB = fromRgb Then
                sh.Fill.ForeColor.RGB = toRgb
            End If
        End If
    End If
    If target = "font" Or target = "both" Then
        If sh.HasTextFrame Then
            Dim p As Long
            For p = 1 To sh.TextFrame.TextRange.Paragraphs.Count
                Dim para As TextRange: Set para = sh.TextFrame.TextRange.Paragraphs(p)
                Dim r As Long
                For r = 1 To para.Runs.Count
                    If para.Runs(r).Font.Color.RGB = fromRgb Then
                        para.Runs(r).Font.Color.RGB = toRgb
                    End If
                Next r
            Next p
        End If
    End If
    If sh.Type = msoGroup Then
        Dim child As Shape
        For Each child In sh.GroupItems
            RecolorInShape child, fromRgb, toRgb, target
        Next child
    End If
    On Error GoTo 0
End Sub

Public Sub Do_apply_theme(themePath As String)
    If Len(themePath) = 0 Then Err.Raise vbObjectError + 7004, "Do_apply_theme", "theme_path empty"
    On Error Resume Next
    ActivePresentation.ApplyTemplate themePath
    If Err.Number <> 0 Then
        Dim msg As String: msg = Err.Description
        Err.Clear
        Err.Raise vbObjectError + 7004, "Do_apply_theme", "ApplyTemplate failed: " & msg
    End If
    On Error GoTo 0
End Sub

Public Sub Do_set_slide_size_dims(widthPt As Double, heightPt As Double)
    If widthPt <= 0 Or heightPt <= 0 Then _
        Err.Raise vbObjectError + 7005, "Do_set_slide_size_dims", "dims must be > 0"
    ActivePresentation.PageSetup.SlideWidth = widthPt
    ActivePresentation.PageSetup.SlideHeight = heightPt
End Sub

Public Sub Do_set_slide_size_preset(preset As String)
    Select Case LCase(preset)
        Case "16:9"
            ActivePresentation.PageSetup.SlideWidth = 960
            ActivePresentation.PageSetup.SlideHeight = 540
        Case "4:3"
            ActivePresentation.PageSetup.SlideWidth = 720
            ActivePresentation.PageSetup.SlideHeight = 540
        Case Else
            Err.Raise vbObjectError + 7005, "Do_set_slide_size_preset", "preset must be 16:9 or 4:3"
    End Select
End Sub

Public Sub Do_set_theme_font(majorName As String, minorName As String)
    If Len(majorName) = 0 And Len(minorName) = 0 Then _
        Err.Raise vbObjectError + 7006, "Do_set_theme_font", "at least one of major/minor must be non-empty"
    ' MajorFont/MinorFont are ThemeFonts collections; index 1 = msoThemeLatin.
    With ActivePresentation.SlideMaster.Theme.ThemeFontScheme
        If Len(majorName) > 0 Then .MajorFont(1).Name = majorName
        If Len(minorName) > 0 Then .MinorFont(1).Name = minorName
    End With
End Sub

Public Sub Do_bulk_insert_image(slideIndices As Variant, picturePath As String, _
                                 leftPt As Double, topPt As Double, _
                                 widthPt As Double, heightPt As Double)
    If Len(picturePath) = 0 Then Err.Raise vbObjectError + 7007, "Do_bulk_insert_image", "picture_path empty"
    Dim ids() As Long
    Dim n As Long: n = modActionsLayout.NormalizeIdsArray(slideIndices, ids)
    If n < 1 Then Err.Raise vbObjectError + 7007, "Do_bulk_insert_image", "no slide indices"
    Dim total As Long: total = ActivePresentation.Slides.Count
    Dim i As Long
    For i = 0 To n - 1
        If ids(i) >= 1 And ids(i) <= total Then
            Dim pic As Shape
            Set pic = ActivePresentation.Slides(ids(i)).Shapes.AddPicture( _
                FileName:=picturePath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                Left:=leftPt, Top:=topPt, Width:=widthPt, Height:=heightPt)
            pic.LockAspectRatio = msoFalse
            pic.Width = widthPt
            pic.Height = heightPt
        End If
    Next i
End Sub

Public Sub Do_bulk_insert_text_box(slideIndices As Variant, text As String, _
                                    leftPt As Double, topPt As Double, _
                                    widthPt As Double, heightPt As Double)
    Dim ids() As Long
    Dim n As Long: n = modActionsLayout.NormalizeIdsArray(slideIndices, ids)
    If n < 1 Then Err.Raise vbObjectError + 7008, "Do_bulk_insert_text_box", "no slide indices"
    Dim total As Long: total = ActivePresentation.Slides.Count
    Dim i As Long
    For i = 0 To n - 1
        If ids(i) >= 1 And ids(i) <= total Then
            Dim sh As Shape
            Set sh = ActivePresentation.Slides(ids(i)).Shapes.AddTextbox( _
                Orientation:=msoTextOrientationHorizontal, _
                Left:=leftPt, Top:=topPt, Width:=widthPt, Height:=heightPt)
            sh.TextFrame.AutoSize = ppAutoSizeNone
            sh.TextFrame.TextRange.Text = text
        End If
    Next i
End Sub

Public Sub Do_apply_layout_to_slides(slideIndices As Variant, layoutIndex As Long)
    If layoutIndex < 0 Then Err.Raise vbObjectError + 7009, "Do_apply_layout_to_slides", "layout_index must be >= 0"
    Dim ids() As Long
    Dim n As Long: n = modActionsLayout.NormalizeIdsArray(slideIndices, ids)
    If n < 1 Then Err.Raise vbObjectError + 7009, "Do_apply_layout_to_slides", "no slide indices"
    Dim master As Master: Set master = ActivePresentation.SlideMaster
    Dim layoutCount As Long: layoutCount = master.CustomLayouts.Count
    If layoutIndex >= layoutCount Then _
        Err.Raise vbObjectError + 7009, "Do_apply_layout_to_slides", "layout_index out of range (max " & (layoutCount - 1) & ")"
    Dim total As Long: total = ActivePresentation.Slides.Count
    Dim i As Long
    For i = 0 To n - 1
        If ids(i) >= 1 And ids(i) <= total Then
            ActivePresentation.Slides(ids(i)).CustomLayout = master.CustomLayouts(layoutIndex + 1)
        End If
    Next i
End Sub
