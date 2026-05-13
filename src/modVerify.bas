Attribute VB_Name = "modVerify"
Option Explicit

' =============================================================================
' modVerify.bas — automatic quality-check sweep after action batches.
'
' Goal: catch the failure modes a human eyeballs (overflow, unreadable labels,
' tiny fonts, duplicate shapes) and surface them as structured warnings so the
' LLM can self-correct in the next batch instead of bouncing back to the user.
'
' Design constraints:
'   - Must be fast: hard caps on shape iteration depth, warning count, and
'     deep chart/table inspection. Stops contributing wall-clock latency
'     beyond ~1 second on typical 20-slide decks.
'   - Must be safe: every check wrapped in On Error Resume Next so a single
'     malformed shape never aborts the sweep.
'   - Returns JSON-serializable Collection of Dictionaries so the existing
'     modJSON.ConvertToJson can emit it directly.
'
' Public surface:
'   RunVerificationLoop(scope, maxWarnings)  -> Collection of warning Dicts
'   FormatWarningsSummary(warnings)          -> human-readable one-line summary
' =============================================================================

' ---- Tunables ------------------------------------------------------------
Private Const SLIDE_BOUNDS_TOL_PT As Double = 2#       ' tolerance for off-slide
Private Const TINY_TABLE_FONT_PT As Double = 8#        ' below this = warn
Private Const MAX_WARNINGS_DEFAULT As Long = 100
Private Const MAX_SHAPES_PER_SLIDE_FOR_DEEP As Long = 100
Private Const MAX_CHART_SERIES_FOR_DEEP As Long = 20
Private Const MAX_TABLE_CELLS_FOR_DEEP As Long = 200
Private Const WCAG_CONTRAST_THRESHOLD As Double = 4.5  ' AA standard for body text

' Per-batch warning ceiling (caller passes maxWarnings; default kicks in if <=0).
Private g_warnCap As Long
Private g_warnCount As Long

' =============================================================================
' Public entry point. scope = "deck" or "slide:N".
' Returns a Collection where each element is a Scripting.Dictionary with keys:
'   severity   - "warn" | "info"
'   kind       - short machine-readable label, e.g. "off_slide_shape"
'   slide      - slide number
'   shape_id   - shape Id (or 0 if N/A)
'   message    - human-readable description
'   suggestion - actionable hint for the LLM
' =============================================================================
Public Function RunVerificationLoop(Optional scope As String = "deck", _
                                     Optional maxWarnings As Long = 0) As Object
    Dim warnings As Collection: Set warnings = New Collection
    Set RunVerificationLoop = warnings

    g_warnCap = IIf(maxWarnings <= 0, MAX_WARNINGS_DEFAULT, maxWarnings)
    g_warnCount = 0

    Dim pres As Presentation: Set pres = ActivePresentation
    If pres Is Nothing Then Exit Function

    Dim startSlide As Long, endSlide As Long
    If LCase(scope) = "deck" Then
        startSlide = 1
        endSlide = pres.Slides.Count
    ElseIf LCase(Left(scope, 6)) = "slide:" Then
        startSlide = CLng(Mid(scope, 7))
        endSlide = startSlide
        If startSlide < 1 Or startSlide > pres.Slides.Count Then Exit Function
    Else
        Exit Function
    End If

    Dim slideW As Double: slideW = pres.PageSetup.SlideWidth
    Dim slideH As Double: slideH = pres.PageSetup.SlideHeight

    Dim s As Long
    For s = startSlide To endSlide
        If g_warnCount >= g_warnCap Then Exit For
        On Error Resume Next
        CheckSlide pres.Slides(s), s, slideW, slideH, warnings
        On Error GoTo 0
    Next s
End Function

' ---- Per-slide orchestrator ---------------------------------------------
Private Sub CheckSlide(sl As Slide, slideNum As Long, _
                        slideW As Double, slideH As Double, _
                        warnings As Collection)
    Dim shapeCount As Long: shapeCount = sl.Shapes.Count
    Dim deepOk As Boolean: deepOk = (shapeCount <= MAX_SHAPES_PER_SLIDE_FOR_DEEP)

    ' --- slide-level checks (run once per slide, not per shape) ---
    CheckCrowdedSlide sl, slideNum, shapeCount, warnings

    Dim sh As Shape
    Dim posMap As Object: Set posMap = CreateObject("Scripting.Dictionary")
    Dim colorMap As Object: Set colorMap = CreateObject("Scripting.Dictionary")
    Dim fontMap As Object: Set fontMap = CreateObject("Scripting.Dictionary")

    For Each sh In sl.Shapes
        If g_warnCount >= g_warnCap Then Exit For

        ' --- always-run lightweight checks ---
        CheckOffSlide sh, slideNum, slideW, slideH, warnings
        CheckDuplicatePosition sh, slideNum, posMap, warnings
        CheckZeroSize sh, slideNum, warnings

        If deepOk Then
            CheckTextOverflow sh, slideNum, warnings
            CheckEmptyShape sh, slideNum, warnings
            CheckShapeTextContrast sh, slideNum, warnings
            CheckTinyShapeFont sh, slideNum, warnings
            CheckMixedFontFamilies sh, slideNum, warnings
            CheckPictureAltText sh, slideNum, warnings
            ' Track distinct colors/fonts for slide-level summary
            TrackDistinctFill sh, colorMap
            TrackDistinctFont sh, fontMap

            If sh.HasChart Then
                CheckChartLabels sh, slideNum, warnings
                CheckChartTitle sh, slideNum, warnings
                CheckChartZeroValues sh, slideNum, warnings
                CheckPieTooManySlices sh, slideNum, warnings
            End If

            If sh.HasTable Then CheckTable sh, slideNum, warnings
        End If
    Next sh

    ' --- slide-level aggregate checks (post-iteration) ---
    If deepOk Then
        CheckTooManyDistinctColors slideNum, colorMap, warnings
        CheckTooManyDistinctFonts slideNum, fontMap, warnings
    End If
End Sub

' ---- Check 1: off-slide bounds -------------------------------------------
Private Sub CheckOffSlide(sh As Shape, slideNum As Long, _
                           slideW As Double, slideH As Double, _
                           warnings As Collection)
    On Error Resume Next
    Dim L As Double, t As Double, w As Double, h As Double
    L = sh.Left: t = sh.Top: w = sh.Width: h = sh.Height
    If Err.Number <> 0 Then Err.Clear: Exit Sub

    Dim msg As String: msg = ""
    Dim overhang As Double

    If L < -SLIDE_BOUNDS_TOL_PT Then
        overhang = -L
        msg = "extends " & Format(overhang, "0.#") & " pt past LEFT edge of slide"
    ElseIf (L + w) > (slideW + SLIDE_BOUNDS_TOL_PT) Then
        overhang = (L + w) - slideW
        msg = "extends " & Format(overhang, "0.#") & " pt past RIGHT edge of slide"
    ElseIf t < -SLIDE_BOUNDS_TOL_PT Then
        overhang = -t
        msg = "extends " & Format(overhang, "0.#") & " pt past TOP edge of slide"
    ElseIf (t + h) > (slideH + SLIDE_BOUNDS_TOL_PT) Then
        overhang = (t + h) - slideH
        msg = "extends " & Format(overhang, "0.#") & " pt past BOTTOM edge of slide"
    End If

    If Len(msg) > 0 Then
        AddWarning warnings, "warn", "off_slide_shape", slideNum, sh.Id, _
            sh.Name & " " & msg, _
            "move or resize shape inside slide bounds (slide is " & _
            Format(slideW, "0") & " x " & Format(slideH, "0") & " pt)"
    End If
End Sub

' ---- Check 2: duplicate position (likely accidental overlap) -------------
Private Sub CheckDuplicatePosition(sh As Shape, slideNum As Long, _
                                    posMap As Object, warnings As Collection)
    On Error Resume Next
    Dim key As String
    key = Format(sh.Left, "0.0") & "," & Format(sh.Top, "0.0") & "," & _
          Format(sh.Width, "0.0") & "," & Format(sh.Height, "0.0")
    If posMap.Exists(key) Then
        Dim otherId As Long: otherId = posMap(key)
        AddWarning warnings, "info", "duplicate_position", slideNum, sh.Id, _
            sh.Name & " has identical bounds to shape Id " & otherId & " (likely accidental duplicate)", _
            "delete one of the shapes if duplicate, or offset to disambiguate"
    Else
        posMap(key) = sh.Id
    End If
End Sub

' ---- Check 3: text-frame content overflows its frame ---------------------
Private Sub CheckTextOverflow(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    If Not sh.HasTextFrame Then Exit Sub
    If Not sh.TextFrame.HasText Then Exit Sub

    ' Skip if shrink-to-fit already enabled — overflow will auto-shrink
    Dim autoFit As Long: autoFit = sh.TextFrame2.AutoSize
    If autoFit = 2 Then Exit Sub   ' msoAutoSizeTextToFitShape

    Dim bh As Double: bh = sh.TextFrame.TextRange.BoundHeight
    Dim bw As Double: bw = sh.TextFrame.TextRange.BoundWidth
    If Err.Number <> 0 Then Err.Clear: Exit Sub

    Dim overflowH As Boolean: overflowH = bh > sh.Height + 2
    Dim overflowW As Boolean: overflowW = bw > sh.Width + 5

    If overflowH Or overflowW Then
        Dim dirStr As String
        If overflowH And overflowW Then
            dirStr = "vertically and horizontally"
        ElseIf overflowH Then
            dirStr = "vertically"
        Else
            dirStr = "horizontally"
        End If
        AddWarning warnings, "warn", "text_overflow", slideNum, sh.Id, _
            sh.Name & " text overflows shape " & dirStr & _
            " (text " & Format(bw, "0") & "x" & Format(bh, "0") & " pt, shape " & _
            Format(sh.Width, "0") & "x" & Format(sh.Height, "0") & " pt)", _
            "reduce font size, enable shrink-to-fit (set_text_autofit shrink), or enlarge shape"
    End If
End Sub

' ---- Check 4: empty placeholder / orphan shape ---------------------------
Private Sub CheckEmptyShape(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    ' Skip groups, tables, charts, pictures — they're not "empty" by definition
    If sh.Type = msoGroup Or sh.HasTable Or sh.HasChart Then Exit Sub
    If sh.Type = msoPicture Or sh.Type = msoLinkedPicture Then Exit Sub
    If sh.Type = msoLine Then Exit Sub

    ' Has visible text?
    Dim hasText As Boolean: hasText = False
    If sh.HasTextFrame Then
        If sh.TextFrame.HasText Then
            If Len(Trim(sh.TextFrame.TextRange.Text)) > 0 Then hasText = True
        End If
    End If
    If hasText Then Exit Sub

    ' Has visible fill?
    Dim hasFill As Boolean: hasFill = (sh.Fill.Visible = msoTrue)
    ' Has visible line?
    Dim hasLine As Boolean: hasLine = (sh.Line.Visible = msoTrue)

    If Not hasFill And Not hasLine Then
        AddWarning warnings, "info", "empty_shape", slideNum, sh.Id, _
            sh.Name & " has no text, no fill, and no outline (invisible placeholder)", _
            "delete the shape or give it visible content"
    End If
End Sub

' ---- Check 5: chart data-label contrast vs series fill -------------------
Private Sub CheckChartLabels(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    Dim ch As Object: Set ch = sh.Chart
    Dim seriesCount As Long: seriesCount = ch.SeriesCollection.Count
    If seriesCount > MAX_CHART_SERIES_FOR_DEEP Then Exit Sub

    ' Only check chart types where data labels sit ON the series fill: column/bar.
    ' Line/scatter/area labels sit above markers and the contrast issue is rare.
    Dim ct As Long: ct = ch.ChartType
    If Not IsBarColumnChart(ct) Then Exit Sub

    Dim s As Long
    For s = 1 To seriesCount
        If g_warnCount >= g_warnCap Then Exit For
        Dim ser As Object: Set ser = ch.SeriesCollection(s)
        If Not ser.HasDataLabels Then GoTo NextSeries

        ' Fill color of the series
        Dim fillRgb As Long: fillRgb = -1
        fillRgb = ser.Format.Fill.ForeColor.RGB
        If fillRgb < 0 Then GoTo NextSeries

        ' Default data-label color (auto / black). We approximate by reading
        ' the first label's font color; if it's near-black we warn for dark
        ' fills, near-white for light fills, etc.
        Dim lblRgb As Long: lblRgb = -1
        lblRgb = ser.DataLabels.Font.Color.RGB
        If lblRgb < 0 Then lblRgb = RGB(0, 0, 0)   ' default to black

        Dim contrast As Double: contrast = ContrastRatio(fillRgb, lblRgb)
        If contrast < WCAG_CONTRAST_THRESHOLD Then
            Dim fillHex As String: fillHex = RgbToHex(fillRgb)
            Dim lblHex As String: lblHex = RgbToHex(lblRgb)
            Dim recommend As String
            If IsDarkColor(fillRgb) Then
                recommend = "#FFFFFF"
            Else
                recommend = "#000000"
            End If
            AddWarning warnings, "warn", "chart_label_contrast", slideNum, sh.Id, _
                "series " & s & " (fill " & fillHex & ") has data labels in " & lblHex & _
                " — contrast ratio " & Format(contrast, "0.0") & " is below WCAG AA (4.5)", _
                "set_chart_series slide=" & slideNum & " shape_id=" & sh.Id & _
                " series_index=" & s & " props.label_color=" & recommend
        End If
NextSeries:
    Next s
End Sub

' ---- Check 6: table — tiny fonts + cell font-vs-fill contrast ------------
Private Sub CheckTable(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    Dim tbl As Table: Set tbl = sh.Table
    Dim rowsN As Long: rowsN = tbl.Rows.Count
    Dim colsN As Long: colsN = tbl.Columns.Count
    Dim cellCount As Long: cellCount = rowsN * colsN
    If cellCount > MAX_TABLE_CELLS_FOR_DEEP Then Exit Sub

    Dim r As Long, c As Long
    For r = 1 To rowsN
        For c = 1 To colsN
            If g_warnCount >= g_warnCap Then Exit Sub
            Dim cs As Shape: Set cs = tbl.Cell(r, c).Shape
            If Not cs.HasTextFrame Then GoTo NextCell
            If Not cs.TextFrame.HasText Then GoTo NextCell

            ' Tiny font check
            Dim sz As Double: sz = cs.TextFrame.TextRange.Font.Size
            If sz > 0 And sz < TINY_TABLE_FONT_PT Then
                AddWarning warnings, "warn", "tiny_table_font", slideNum, sh.Id, _
                    "cell (" & r & "," & c & ") font size " & Format(sz, "0.#") & _
                    "pt is below " & TINY_TABLE_FONT_PT & "pt — likely unreadable", _
                    "increase font: set_cell_font_size slide=" & slideNum & " shape_id=" & sh.Id & _
                    " row=" & r & " col=" & c & " value=10 (or auto_fit_table_text)"
            End If

            ' Contrast check: cell fill color vs font color
            If cs.Fill.Visible = msoTrue Then
                Dim cellFill As Long: cellFill = cs.Fill.ForeColor.RGB
                Dim cellFont As Long: cellFont = cs.TextFrame.TextRange.Font.Color.RGB
                If cellFill >= 0 And cellFont >= 0 Then
                    Dim ctr As Double: ctr = ContrastRatio(cellFill, cellFont)
                    If ctr < WCAG_CONTRAST_THRESHOLD Then
                        Dim recCol As String
                        recCol = IIf(IsDarkColor(cellFill), "#FFFFFF", "#000000")
                        AddWarning warnings, "warn", "cell_text_contrast", slideNum, sh.Id, _
                            "cell (" & r & "," & c & ") font color " & RgbToHex(cellFont) & _
                            " on fill " & RgbToHex(cellFill) & " — contrast " & _
                            Format(ctr, "0.0") & " below AA threshold (4.5)", _
                            "set_cell_font_color slide=" & slideNum & " shape_id=" & sh.Id & _
                            " row=" & r & " col=" & c & " value=" & recCol
                    End If
                End If
            End If
NextCell:
        Next c
    Next r
End Sub

' ---- Check 7: zero-size / micro-size shape -------------------------------
Private Sub CheckZeroSize(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    Dim w As Double: w = sh.Width
    Dim h As Double: h = sh.Height
    If Err.Number <> 0 Then Err.Clear: Exit Sub
    If w < 1 Or h < 1 Then
        AddWarning warnings, "warn", "zero_size_shape", slideNum, sh.Id, _
            sh.Name & " has size " & Format(w, "0.#") & " x " & Format(h, "0.#") & _
            " pt — effectively invisible", _
            "resize_shape (width and height >= 10) or delete_shape if unwanted"
    End If
End Sub

' ---- Check 8: generic shape text vs fill contrast ------------------------
' Same idea as cell_text_contrast but for any non-table text-bearing shape
' that has a visible solid fill. Skips groups, charts, tables, pictures.
Private Sub CheckShapeTextContrast(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    If sh.Type = msoGroup Or sh.HasTable Or sh.HasChart Then Exit Sub
    If sh.Type = msoPicture Or sh.Type = msoLinkedPicture Then Exit Sub
    If Not sh.HasTextFrame Then Exit Sub
    If Not sh.TextFrame.HasText Then Exit Sub
    If sh.Fill.Visible <> msoTrue Then Exit Sub
    Dim fillRgb As Long: fillRgb = sh.Fill.ForeColor.RGB
    Dim fontRgb As Long: fontRgb = sh.TextFrame.TextRange.Font.Color.RGB
    If fillRgb < 0 Or fontRgb < 0 Then Exit Sub
    Dim ratio As Double: ratio = ContrastRatio(fillRgb, fontRgb)
    If ratio < WCAG_CONTRAST_THRESHOLD Then
        Dim rec As String: rec = IIf(IsDarkColor(fillRgb), "#FFFFFF", "#000000")
        AddWarning warnings, "warn", "shape_text_contrast", slideNum, sh.Id, _
            sh.Name & " text " & RgbToHex(fontRgb) & " on fill " & RgbToHex(fillRgb) & _
            " — contrast " & Format(ratio, "0.0") & " below AA threshold (4.5)", _
            "set_font_color slide=" & slideNum & " shape_id=" & sh.Id & " value=" & rec
    End If
End Sub

' ---- Check 9: tiny font on any text shape (not just tables) --------------
Private Sub CheckTinyShapeFont(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    If sh.HasTable Or sh.HasChart Then Exit Sub
    If Not sh.HasTextFrame Then Exit Sub
    If Not sh.TextFrame.HasText Then Exit Sub
    Dim sz As Double: sz = sh.TextFrame.TextRange.Font.Size
    If sz > 0 And sz < TINY_TABLE_FONT_PT Then
        AddWarning warnings, "warn", "tiny_shape_font", slideNum, sh.Id, _
            sh.Name & " font size " & Format(sz, "0.#") & "pt is below readable threshold (" & _
            TINY_TABLE_FONT_PT & "pt)", _
            "set_font_size slide=" & slideNum & " shape_id=" & sh.Id & " value=10 (or larger)"
    End If
End Sub

' ---- Check 10: too many fonts mixed in one shape's text ------------------
Private Sub CheckMixedFontFamilies(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    If sh.HasTable Or sh.HasChart Then Exit Sub
    If Not sh.HasTextFrame Then Exit Sub
    If Not sh.TextFrame.HasText Then Exit Sub
    Dim names As Object: Set names = CreateObject("Scripting.Dictionary")
    Dim paraN As Long: paraN = sh.TextFrame.TextRange.Paragraphs().Count
    Dim p As Long
    For p = 1 To paraN
        Dim para As TextRange: Set para = sh.TextFrame.TextRange.Paragraphs(p)
        Dim runN As Long: runN = para.Runs.Count
        Dim r As Long
        For r = 1 To runN
            Dim fn As String: fn = para.Runs(r).Font.Name
            If Len(fn) > 0 Then
                If Not names.Exists(fn) Then names.Add fn, 1
            End If
        Next r
    Next p
    If names.Count >= 3 Then
        AddWarning warnings, "info", "mixed_font_families", slideNum, sh.Id, _
            sh.Name & " uses " & names.Count & " different fonts in one shape (visually noisy)", _
            "consolidate to 1-2 fonts via set_paragraph_font_name or set_run_font_name"
    End If
End Sub

' ---- Check 11: picture without alt text (accessibility) ------------------
Private Sub CheckPictureAltText(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    If sh.Type <> msoPicture And sh.Type <> msoLinkedPicture Then Exit Sub
    Dim alt As String: alt = sh.AlternativeText
    If Len(Trim(alt)) = 0 Then
        AddWarning warnings, "info", "picture_no_alt_text", slideNum, sh.Id, _
            sh.Name & " is a picture with no alt text (fails accessibility)", _
            "set_shape_alt_text slide=" & slideNum & " shape_id=" & sh.Id & " value=""<describe image>"""
    End If
End Sub

' ---- Check 12: chart with no title --------------------------------------
Private Sub CheckChartTitle(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    Dim ch As Object: Set ch = sh.Chart
    If Not ch.HasTitle Then
        AddWarning warnings, "info", "chart_no_title", slideNum, sh.Id, _
            sh.Name & " is a chart with no title", _
            "set_chart_title slide=" & slideNum & " shape_id=" & sh.Id & " value=""<title>"""
    End If
End Sub

' ---- Check 13: chart with all-zero series (broken data) ------------------
Private Sub CheckChartZeroValues(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    Dim ch As Object: Set ch = sh.Chart
    Dim seriesCount As Long: seriesCount = ch.SeriesCollection.Count
    If seriesCount > MAX_CHART_SERIES_FOR_DEEP Then Exit Sub
    Dim totalNonZero As Long: totalNonZero = 0
    Dim s As Long
    For s = 1 To seriesCount
        Dim vals As Variant: vals = ch.SeriesCollection(s).Values
        If IsArray(vals) Then
            Dim i As Long
            For i = LBound(vals) To UBound(vals)
                If IsNumeric(vals(i)) Then
                    If CDbl(vals(i)) <> 0 Then totalNonZero = totalNonZero + 1
                End If
            Next i
        End If
    Next s
    If totalNonZero = 0 Then
        AddWarning warnings, "warn", "chart_all_zero_values", slideNum, sh.Id, _
            sh.Name & " chart contains only zero values across all series — likely missing data", _
            "use set_series_values to populate real data"
    End If
End Sub

' ---- Check 14: pie/doughnut chart with too many slices -------------------
Private Sub CheckPieTooManySlices(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    Dim ch As Object: Set ch = sh.Chart
    Dim ct As Long: ct = ch.ChartType
    ' XlChartType: pie=5, pie3d=70, pieExploded=69, pieExploded3D=70, doughnut=-4120
    If Not (ct = 5 Or ct = 69 Or ct = 70 Or ct = -4120) Then Exit Sub
    If ch.SeriesCollection.Count = 0 Then Exit Sub
    Dim catCount As Long: catCount = ch.SeriesCollection(1).Points.Count
    If catCount > 8 Then
        AddWarning warnings, "info", "pie_too_many_slices", slideNum, sh.Id, _
            sh.Name & " is a pie/doughnut with " & catCount & _
            " slices — labels likely overlap and the audience cannot compare", _
            "consolidate to <=7 slices (group small categories into 'Other') or switch to bar chart"
    End If
End Sub

' ---- Slide-level: too many shapes ----------------------------------------
Private Sub CheckCrowdedSlide(sl As Slide, slideNum As Long, shapeCount As Long, warnings As Collection)
    If shapeCount > 40 Then
        AddWarning warnings, "info", "crowded_slide", slideNum, 0, _
            "slide " & slideNum & " has " & shapeCount & " shapes (>40) — visually crowded", _
            "split into multiple slides or group related elements"
    End If
End Sub

' ---- Tracking helpers for slide-level aggregates -------------------------
Private Sub TrackDistinctFill(sh As Shape, colorMap As Object)
    On Error Resume Next
    If sh.HasTable Or sh.HasChart Then Exit Sub
    If sh.Type = msoGroup Or sh.Type = msoLine Or sh.Type = msoPicture Then Exit Sub
    If sh.Fill.Visible <> msoTrue Then Exit Sub
    Dim k As String: k = CStr(sh.Fill.ForeColor.RGB)
    If Not colorMap.Exists(k) Then colorMap.Add k, 1
End Sub

Private Sub TrackDistinctFont(sh As Shape, fontMap As Object)
    On Error Resume Next
    If Not sh.HasTextFrame Then Exit Sub
    If Not sh.TextFrame.HasText Then Exit Sub
    Dim fn As String: fn = sh.TextFrame.TextRange.Font.Name
    If Len(fn) > 0 Then
        If Not fontMap.Exists(fn) Then fontMap.Add fn, 1
    End If
End Sub

Private Sub CheckTooManyDistinctColors(slideNum As Long, colorMap As Object, warnings As Collection)
    If colorMap.Count > 10 Then
        AddWarning warnings, "info", "too_many_colors", slideNum, 0, _
            "slide " & slideNum & " uses " & colorMap.Count & _
            " distinct fill colors (>10) — palette feels disorganized", _
            "constrain to brand palette (3-5 fills) via recolor_palette_deck_wide or recolor_fill_match"
    End If
End Sub

Private Sub CheckTooManyDistinctFonts(slideNum As Long, fontMap As Object, warnings As Collection)
    If fontMap.Count > 3 Then
        AddWarning warnings, "info", "too_many_fonts", slideNum, 0, _
            "slide " & slideNum & " uses " & fontMap.Count & _
            " distinct font families (>3) — looks inconsistent", _
            "unify via swap_font_deck_wide or set_theme_font"
    End If
End Sub

' ---- Helpers --------------------------------------------------------------

Private Sub AddWarning(warnings As Collection, severity As String, kind As String, _
                        slideNum As Long, shapeId As Long, _
                        message As String, suggestion As String)
    If g_warnCount >= g_warnCap Then Exit Sub
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("severity") = severity
    d("kind") = kind
    d("slide") = slideNum
    d("shape_id") = shapeId
    d("message") = message
    d("suggestion") = suggestion
    warnings.Add d
    g_warnCount = g_warnCount + 1
End Sub

Private Function IsBarColumnChart(ct As Long) As Boolean
    ' XlChartType values for clustered/stacked/100% column and bar (2D + 3D)
    Select Case ct
        Case 51, 52, 53, 54, 55, 56     ' column variants
            IsBarColumnChart = True
        Case 57, 58, 59, 60, 61, 62     ' bar variants
            IsBarColumnChart = True
        Case Else
            IsBarColumnChart = False
    End Select
End Function

' WCAG relative luminance + contrast ratio.
' Returns ratio >= 1.0 (higher = better contrast). AA body text needs >= 4.5.
Public Function ContrastRatio(rgb1 As Long, rgb2 As Long) As Double
    Dim l1 As Double: l1 = RelativeLuminance(rgb1)
    Dim l2 As Double: l2 = RelativeLuminance(rgb2)
    Dim lighter As Double, darker As Double
    If l1 > l2 Then lighter = l1: darker = l2 Else lighter = l2: darker = l1
    ContrastRatio = (lighter + 0.05) / (darker + 0.05)
End Function

Private Function RelativeLuminance(rgbVal As Long) As Double
    ' Extract 0-255 channels (VBA RGB is BGR: blue in high byte, red in low byte)
    Dim r As Long, g As Long, b As Long
    r = rgbVal And &HFF
    g = (rgbVal \ &H100) And &HFF
    b = (rgbVal \ &H10000) And &HFF
    Dim rs As Double: rs = LinearizeChannel(r / 255#)
    Dim gs As Double: gs = LinearizeChannel(g / 255#)
    Dim bs As Double: bs = LinearizeChannel(b / 255#)
    RelativeLuminance = 0.2126 * rs + 0.7152 * gs + 0.0722 * bs
End Function

Private Function LinearizeChannel(c As Double) As Double
    If c <= 0.03928 Then
        LinearizeChannel = c / 12.92
    Else
        LinearizeChannel = ((c + 0.055) / 1.055) ^ 2.4
    End If
End Function

Public Function IsDarkColor(rgbVal As Long) As Boolean
    IsDarkColor = (RelativeLuminance(rgbVal) < 0.5)
End Function

Public Function RgbToHex(rgbVal As Long) As String
    Dim r As Long, g As Long, b As Long
    r = rgbVal And &HFF
    g = (rgbVal \ &H100) And &HFF
    b = (rgbVal \ &H10000) And &HFF
    RgbToHex = "#" & Right("00" & Hex(r), 2) & Right("00" & Hex(g), 2) & Right("00" & Hex(b), 2)
End Function

' Human-readable one-line summary of warnings collection.
Public Function FormatWarningsSummary(warnings As Collection) As String
    If warnings.Count = 0 Then
        FormatWarningsSummary = "verification: clean (0 warnings)"
        Exit Function
    End If
    Dim warnN As Long: warnN = 0
    Dim infoN As Long: infoN = 0
    Dim i As Long
    For i = 1 To warnings.Count
        If LCase(CStr(warnings(i)("severity"))) = "warn" Then
            warnN = warnN + 1
        Else
            infoN = infoN + 1
        End If
    Next i
    FormatWarningsSummary = "verification: " & warnN & " warning(s), " & infoN & " info"
End Function

' Serialize warnings Collection to compact JSON for the action-log return value.
Public Function WarningsToJson(warnings As Collection) As String
    If warnings.Count = 0 Then
        WarningsToJson = "[]"
        Exit Function
    End If
    Dim sb As String: sb = "["
    Dim i As Long
    For i = 1 To warnings.Count
        Dim w As Object: Set w = warnings(i)
        If i > 1 Then sb = sb & ","
        sb = sb & "{""severity"":""" & JsonEscape(CStr(w("severity"))) & """"
        sb = sb & ",""kind"":""" & JsonEscape(CStr(w("kind"))) & """"
        sb = sb & ",""slide"":" & CStr(w("slide"))
        sb = sb & ",""shape_id"":" & CStr(w("shape_id"))
        sb = sb & ",""message"":""" & JsonEscape(CStr(w("message"))) & """"
        sb = sb & ",""suggestion"":""" & JsonEscape(CStr(w("suggestion"))) & """}"
    Next i
    WarningsToJson = sb & "]"
End Function

Private Function JsonEscape(s As String) As String
    Dim out As String: out = s
    out = Replace(out, "\", "\\")
    out = Replace(out, """", "\""")
    out = Replace(out, Chr(13), "\n")
    out = Replace(out, Chr(10), "")
    out = Replace(out, Chr(9), "\t")
    JsonEscape = out
End Function
