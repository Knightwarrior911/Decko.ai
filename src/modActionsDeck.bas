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

' Batch palette remap — all N from→to pairs in one deck pass.
' mappingsObj: Collection of Dicts each with "from"/"to" hex strings.
' scope: fill|font|border|table_fill|table_font|table_border|chart|all (default "all")
Public Sub Do_recolor_deck(mappingsObj As Object, scope As String)
    Dim colorMap As Object: Set colorMap = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To mappingsObj.Count
        Dim entry As Object: Set entry = mappingsObj(i)
        Dim fRgb As Long: fRgb = modActions.HexToRgb(CStr(entry("from")))
        Dim tRgb As Long: tRgb = modActions.HexToRgb(CStr(entry("to")))
        colorMap(fRgb) = tRgb
    Next i
    If colorMap.Count = 0 Then Exit Sub

    Dim sc As String: sc = LCase(Trim(scope))
    If Len(sc) = 0 Then sc = "all"

    Dim sl As Slide
    For Each sl In ActivePresentation.Slides
        ' Slide background (solid fills only)
        If sc = "all" Or sc = "fill" Then
            On Error Resume Next
            If sl.Background.Fill.Type = msoFillSolid Then
                Dim bgRgb As Long: bgRgb = sl.Background.Fill.ForeColor.RGB
                If Err.Number = 0 Then
                    If colorMap.Exists(bgRgb) Then sl.Background.Fill.ForeColor.RGB = colorMap(bgRgb)
                End If
            End If
            Err.Clear
            On Error GoTo 0
        End If
        ' Shapes
        Dim sh As Shape
        For Each sh In sl.Shapes
            RecolorShapeFull sh, colorMap, sc
        Next sh
    Next sl
End Sub

Private Sub RecolorShapeFull(sh As Shape, colorMap As Object, sc As String)
    On Error Resume Next

    Dim doFill    As Boolean: doFill    = (sc = "all" Or sc = "fill")
    Dim doBorder  As Boolean: doBorder  = (sc = "all" Or sc = "border")
    Dim doFont    As Boolean: doFont    = (sc = "all" Or sc = "font")
    Dim doTFill   As Boolean: doTFill   = (sc = "all" Or sc = "table_fill")
    Dim doTFont   As Boolean: doTFont   = (sc = "all" Or sc = "table_font")
    Dim doTBorder As Boolean: doTBorder = (sc = "all" Or sc = "table_border")
    Dim doChart   As Boolean: doChart   = (sc = "all" Or sc = "chart")
    Dim rgb0      As Long

    ' ── shape fill ──────────────────────────────────────────────────────────
    If doFill And Not sh.HasTable And Not sh.HasChart Then
        If sh.Fill.Type = msoFillSolid Then
            Err.Clear: rgb0 = sh.Fill.ForeColor.RGB
            If Err.Number = 0 And colorMap.Exists(rgb0) Then sh.Fill.ForeColor.RGB = colorMap(rgb0)
        End If
        Err.Clear
    End If

    ' ── shape line/border ────────────────────────────────────────────────────
    If doBorder And Not sh.HasTable Then
        Err.Clear: rgb0 = sh.Line.ForeColor.RGB
        If Err.Number = 0 And colorMap.Exists(rgb0) Then sh.Line.ForeColor.RGB = colorMap(rgb0)
        Err.Clear
    End If

    ' ── shape text runs ──────────────────────────────────────────────────────
    If doFont And sh.HasTextFrame Then
        Dim p2 As Long
        For p2 = 1 To sh.TextFrame.TextRange.Paragraphs.Count
            Dim par2 As TextRange: Set par2 = sh.TextFrame.TextRange.Paragraphs(p2)
            Dim r2 As Long
            For r2 = 1 To par2.Runs.Count
                Err.Clear: rgb0 = par2.Runs(r2).Font.Color.RGB
                If Err.Number = 0 And colorMap.Exists(rgb0) Then par2.Runs(r2).Font.Color.RGB = colorMap(rgb0)
                Err.Clear
            Next r2
        Next p2
    End If

    ' ── table ────────────────────────────────────────────────────────────────
    If sh.HasTable Then
        Dim tbl2 As Object: Set tbl2 = sh.Table
        Dim rw As Long, cl As Long
        For rw = 1 To tbl2.Rows.Count
            For cl = 1 To tbl2.Columns.Count
                Dim tcell As Object: Set tcell = tbl2.Cell(rw, cl)

                If doTFill Then
                    If tcell.Shape.Fill.Type = msoFillSolid Then
                        Err.Clear: rgb0 = tcell.Shape.Fill.ForeColor.RGB
                        If Err.Number = 0 And colorMap.Exists(rgb0) Then tcell.Shape.Fill.ForeColor.RGB = colorMap(rgb0)
                    End If
                    Err.Clear
                End If

                If doTFont And tcell.Shape.HasTextFrame Then
                    Dim cp2 As Long
                    For cp2 = 1 To tcell.Shape.TextFrame.TextRange.Paragraphs.Count
                        Dim cpara2 As TextRange
                        Set cpara2 = tcell.Shape.TextFrame.TextRange.Paragraphs(cp2)
                        Dim cr2 As Long
                        For cr2 = 1 To cpara2.Runs.Count
                            Err.Clear: rgb0 = cpara2.Runs(cr2).Font.Color.RGB
                            If Err.Number = 0 And colorMap.Exists(rgb0) Then cpara2.Runs(cr2).Font.Color.RGB = colorMap(rgb0)
                            Err.Clear
                        Next cr2
                    Next cp2
                End If

                If doTBorder Then
                    Dim bd As Long
                    For bd = 1 To 6  ' ppBorderTop=1 .. ppBorderDiagonalUp=6
                        Err.Clear: rgb0 = tcell.Borders(bd).ForeColor.RGB
                        If Err.Number = 0 And colorMap.Exists(rgb0) Then tcell.Borders(bd).ForeColor.RGB = colorMap(rgb0)
                        Err.Clear
                    Next bd
                End If
            Next cl
        Next rw
    End If

    ' ── chart series fill + line ─────────────────────────────────────────────
    If doChart And sh.HasChart Then
        Dim chObj As Object: Set chObj = sh.Chart
        Dim si As Long
        For si = 1 To chObj.SeriesCollection.Count
            Dim serObj As Object
            Err.Clear: Set serObj = chObj.SeriesCollection(si)
            If Err.Number = 0 Then
                Err.Clear: rgb0 = serObj.Format.Fill.ForeColor.RGB
                If Err.Number = 0 And colorMap.Exists(rgb0) Then serObj.Format.Fill.ForeColor.RGB = colorMap(rgb0)
                Err.Clear
                rgb0 = serObj.Format.Line.ForeColor.RGB
                If Err.Number = 0 And colorMap.Exists(rgb0) Then serObj.Format.Line.ForeColor.RGB = colorMap(rgb0)
                Err.Clear
            End If
        Next si
        Err.Clear
    End If

    ' ── group recursion ──────────────────────────────────────────────────────
    If sh.Type = msoGroup Then
        Dim gchild As Shape
        For Each gchild In sh.GroupItems
            RecolorShapeFull gchild, colorMap, sc
        Next gchild
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

' Scan active presentation for explicit RGB colors; write role-tagged JSON to clipboard.
' scope: "deck" (default) or "slide:N"
' Output: JSON array sorted by count desc, each entry {"hex":"#RRGGBB","count":N,"roles":[...]}
' Roles: "fill" = solid shape/table/background fill; "font" = text run; "border" = shape line / table cell border
Public Sub Do_scan_palette(scope As String)
    If Len(scope) = 0 Then scope = "deck"
    Dim startSlide As Long, endSlide As Long
    ParseScopeRange scope, startSlide, endSlide

    Dim countDict  As Object: Set countDict  = CreateObject("Scripting.Dictionary")
    Dim fillDict   As Object: Set fillDict   = CreateObject("Scripting.Dictionary")
    Dim fontDict   As Object: Set fontDict   = CreateObject("Scripting.Dictionary")
    Dim borderDict As Object: Set borderDict = CreateObject("Scripting.Dictionary")

    Dim s As Long
    For s = startSlide To endSlide
        Dim sl As Slide: Set sl = ActivePresentation.Slides(s)
        ' Slide background fill
        On Error Resume Next
        Dim bgType As Long: bgType = sl.Background.Fill.Type
        If Err.Number = 0 And bgType = msoFillSolid Then
            Dim bgRgb As Long: bgRgb = sl.Background.Fill.ForeColor.RGB
            If Err.Number = 0 Then _
                RecordPaletteColor RgbLongToHex(bgRgb), countDict, fillDict, fontDict, borderDict, True, False, False
        End If
        Err.Clear: On Error GoTo 0
        ' Shapes
        Dim sh As Shape
        For Each sh In sl.Shapes
            ScanShapeForPalette sh, countDict, fillDict, fontDict, borderDict
        Next sh
    Next s

    ' Collect keys → array
    Dim nKeys As Long: nKeys = countDict.Count
    If nKeys = 0 Then
        Dim emptyDO As Object: Set emptyDO = CreateObject("MSForms.DataObject")
        emptyDO.SetText "[]": emptyDO.PutInClipboard
        Exit Sub
    End If

    Dim keys() As String: ReDim keys(1 To nKeys)
    Dim ki As Long: ki = 1
    Dim kv As Variant
    For Each kv In countDict.Keys
        keys(ki) = CStr(kv): ki = ki + 1
    Next kv

    ' Bubble sort by count desc
    Dim ii As Long, jj As Long, tmp As String
    For ii = 1 To nKeys - 1
        For jj = 1 To nKeys - ii
            If countDict(keys(jj)) < countDict(keys(jj + 1)) Then
                tmp = keys(jj): keys(jj) = keys(jj + 1): keys(jj + 1) = tmp
            End If
        Next jj
    Next ii

    ' Threshold: include count<2 only if fewer than 5 would survive otherwise
    Dim passing As Long: passing = 0
    For ii = 1 To nKeys
        If countDict(keys(ii)) >= 2 Then passing = passing + 1
    Next ii
    Dim includeAll As Boolean: includeAll = (passing < 5)

    ' Build JSON lines
    Dim lines() As String: ReDim lines(1 To nKeys)
    Dim included As Long: included = 0
    For ii = 1 To nKeys
        Dim cnt As Long: cnt = countDict(keys(ii))
        If includeAll Or cnt >= 2 Then
            included = included + 1
            Dim roles As String: roles = ""
            If fillDict.Exists(keys(ii))   Then roles = roles & """fill"","
            If fontDict.Exists(keys(ii))   Then roles = roles & """font"","
            If borderDict.Exists(keys(ii)) Then roles = roles & """border"","
            If Len(roles) > 0 Then roles = Left(roles, Len(roles) - 1)
            lines(included) = "  {""hex"":""" & keys(ii) & """,""count"":" & cnt & _
                               ",""roles"":[" & roles & "]}"
        End If
    Next ii

    Dim jsonOut As String: jsonOut = "[" & vbCrLf
    For ii = 1 To included
        jsonOut = jsonOut & lines(ii)
        If ii < included Then jsonOut = jsonOut & ","
        jsonOut = jsonOut & vbCrLf
    Next ii
    jsonOut = jsonOut & "]"

    ' Write to temp file (used by programmatic callers + clipboard fallback)
    Dim tmpPath As String: tmpPath = Environ("TEMP") & "\decko_palette.json"
    Dim fNum As Integer: fNum = FreeFile()
    Open tmpPath For Output As #fNum
    Print #fNum, jsonOut
    Close #fNum

    ' Set Windows clipboard via PowerShell (synchronous, hidden window).
    ' If PowerShell execution policy or WScript.Shell is blocked, surface a
    ' Debug.Print warning so the failure is visible during development; the
    ' palette JSON is still written to tmpPath and reported in the result.
    On Error Resume Next
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        Debug.Print "[scan_palette] clipboard set failed (WScript.Shell unavailable): " & _
                    Err.Description & " — palette saved at " & tmpPath
        Err.Clear: On Error GoTo 0
        Exit Sub
    End If
    Dim psCmd As String
    psCmd = "powershell -NonInteractive -Command " & _
            """Get-Content -Raw -Encoding UTF8 '" & tmpPath & "' | Set-Clipboard"""
    Dim psRc As Long: psRc = wsh.Run(psCmd, 0, True)
    If Err.Number <> 0 Or psRc <> 0 Then
        Debug.Print "[scan_palette] PowerShell clipboard set failed (rc=" & psRc & _
                    "): " & Err.Description & " — palette saved at " & tmpPath
    End If
    Err.Clear: On Error GoTo 0
End Sub

Private Sub RecordPaletteColor(hexStr As String, _
    countDict As Object, fillDict As Object, fontDict As Object, borderDict As Object, _
    isFill As Boolean, isFont As Boolean, isBorder As Boolean)
    If Not countDict.Exists(hexStr) Then countDict(hexStr) = 0
    countDict(hexStr) = countDict(hexStr) + 1
    If isFill   Then fillDict(hexStr)   = True
    If isFont   Then fontDict(hexStr)   = True
    If isBorder Then borderDict(hexStr) = True
End Sub

Private Function RgbLongToHex(rgbVal As Long) As String
    Dim r As Long: r = rgbVal And &HFF
    Dim g As Long: g = (rgbVal \ 256) And &HFF
    Dim b As Long: b = (rgbVal \ 65536) And &HFF
    RgbLongToHex = "#" & Right("0" & UCase(Hex(r)), 2) & _
                         Right("0" & UCase(Hex(g)), 2) & _
                         Right("0" & UCase(Hex(b)), 2)
End Function

Private Sub ScanShapeForPalette(sh As Shape, _
    countDict As Object, fillDict As Object, fontDict As Object, borderDict As Object)
    On Error Resume Next

    ' Shape fill
    If Not sh.HasTable And Not sh.HasChart Then
        If sh.Fill.Type = msoFillSolid Then
            Err.Clear
            Dim fRgb As Long: fRgb = sh.Fill.ForeColor.RGB
            If Err.Number = 0 Then _
                RecordPaletteColor RgbLongToHex(fRgb), countDict, fillDict, fontDict, borderDict, True, False, False
        End If
        Err.Clear
    End If

    ' Shape border/line
    If Not sh.HasTable Then
        Err.Clear
        Dim lRgb As Long: lRgb = sh.Line.ForeColor.RGB
        If Err.Number = 0 Then _
            RecordPaletteColor RgbLongToHex(lRgb), countDict, fillDict, fontDict, borderDict, False, False, True
        Err.Clear
    End If

    ' Shape text runs
    If sh.HasTextFrame Then
        Dim p As Long
        For p = 1 To sh.TextFrame.TextRange.Paragraphs.Count
            Dim par As TextRange: Set par = sh.TextFrame.TextRange.Paragraphs(p)
            Dim r As Long
            For r = 1 To par.Runs.Count
                Err.Clear
                Dim tRgb As Long: tRgb = par.Runs(r).Font.Color.RGB
                If Err.Number = 0 Then _
                    RecordPaletteColor RgbLongToHex(tRgb), countDict, fillDict, fontDict, borderDict, False, True, False
                Err.Clear
            Next r
        Next p
    End If

    ' Table
    If sh.HasTable Then
        Dim tbl3 As Object: Set tbl3 = sh.Table
        Dim rw As Long, cl As Long
        For rw = 1 To tbl3.Rows.Count
            For cl = 1 To tbl3.Columns.Count
                Dim tcell As Object: Set tcell = tbl3.Cell(rw, cl)
                ' cell fill
                Dim cfType As Long: Err.Clear: cfType = tcell.Shape.Fill.Type
                If Err.Number = 0 And cfType = msoFillSolid Then
                    Dim cfRgb As Long: Err.Clear: cfRgb = tcell.Shape.Fill.ForeColor.RGB
                    If Err.Number = 0 Then _
                        RecordPaletteColor RgbLongToHex(cfRgb), countDict, fillDict, fontDict, borderDict, True, False, False
                End If
                Err.Clear
                ' cell text
                If tcell.Shape.HasTextFrame Then
                    Dim cp As Long
                    For cp = 1 To tcell.Shape.TextFrame.TextRange.Paragraphs.Count
                        Dim cpara As TextRange
                        Set cpara = tcell.Shape.TextFrame.TextRange.Paragraphs(cp)
                        Dim cr As Long
                        For cr = 1 To cpara.Runs.Count
                            Err.Clear
                            Dim ctRgb As Long: ctRgb = cpara.Runs(cr).Font.Color.RGB
                            If Err.Number = 0 Then _
                                RecordPaletteColor RgbLongToHex(ctRgb), countDict, fillDict, fontDict, borderDict, False, True, False
                            Err.Clear
                        Next cr
                    Next cp
                End If
                ' cell borders
                Dim bd As Long
                For bd = 1 To 6
                    Err.Clear
                    Dim cbRgb As Long: cbRgb = tcell.Borders(bd).ForeColor.RGB
                    If Err.Number = 0 Then _
                        RecordPaletteColor RgbLongToHex(cbRgb), countDict, fillDict, fontDict, borderDict, False, False, True
                    Err.Clear
                Next bd
            Next cl
        Next rw
    End If

    ' Group recursion
    If sh.Type = msoGroup Then
        Dim gchild As Shape
        For Each gchild In sh.GroupItems
            ScanShapeForPalette gchild, countDict, fillDict, fontDict, borderDict
        Next gchild
    End If

    On Error GoTo 0
End Sub
