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
    Dim textMap As Object: Set textMap = CreateObject("Scripting.Dictionary")
    Dim chartShapes As Collection: Set chartShapes = New Collection
    Dim coveringShapes As Collection: Set coveringShapes = New Collection

    ' --- slide-level pre-iteration: title check ---
    CheckSlideTitle sl, slideNum, warnings

    For Each sh In sl.Shapes
        If g_warnCount >= g_warnCap Then Exit For

        ' --- always-run lightweight checks ---
        CheckOffSlide sh, slideNum, slideW, slideH, warnings
        CheckDuplicatePosition sh, slideNum, posMap, warnings
        CheckZeroSize sh, slideNum, warnings
        CheckShapeInSafeArea sh, slideNum, slideW, slideH, warnings
        CheckOrphanConnector sh, slideNum, warnings

        If deepOk Then
            CheckTextOverflow sh, slideNum, warnings
            CheckEmptyShape sh, slideNum, warnings
            CheckShapeTextContrast sh, slideNum, warnings
            CheckTinyShapeFont sh, slideNum, warnings
            CheckHugeBodyFont sh, slideNum, warnings
            CheckMixedFontFamilies sh, slideNum, warnings
            CheckPictureAltText sh, slideNum, warnings
            CheckPlaceholderText sh, slideNum, warnings
            CheckTrailingWhitespace sh, slideNum, warnings
            CheckBrokenInternalHyperlink sh, slideNum, warnings
            ' Track distinct colors/fonts/texts for slide-level aggregates
            TrackDistinctFill sh, colorMap
            TrackDistinctFont sh, fontMap
            TrackDuplicateText sh, slideNum, textMap, warnings
            ' Track charts + non-chart shapes for Z-order check
            If sh.HasChart Then chartShapes.Add sh Else coveringShapes.Add sh

            If sh.HasChart Then
                CheckChartLabels sh, slideNum, warnings
                CheckChartTitle sh, slideNum, warnings
                CheckChartZeroValues sh, slideNum, warnings
                CheckPieTooManySlices sh, slideNum, warnings
                CheckChartDefaultSeriesNames sh, slideNum, warnings
                CheckChartLegendPointless sh, slideNum, warnings
                CheckChartAxisUnitsMismatch sh, slideNum, warnings
            End If

            If sh.HasTable Then
                CheckTable sh, slideNum, warnings
                CheckTableColumnOverflow sh, slideNum, warnings
            End If
        End If
    Next sh

    ' --- slide-level aggregate checks (post-iteration) ---
    If deepOk Then
        CheckTooManyDistinctColors slideNum, colorMap, warnings
        CheckTooManyDistinctFonts slideNum, fontMap, warnings
        CheckChartCovered slideNum, chartShapes, coveringShapes, warnings
        CheckRowColumnAlignment sl, slideNum, warnings
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

' ====================================================================
' WAVE 2 CHECKS — broader coverage across every Decko domain.
' ====================================================================

' ---- Check 15: shape inside slide but cramped within 12pt margin --------
Private Sub CheckShapeInSafeArea(sh As Shape, slideNum As Long, _
                                  slideW As Double, slideH As Double, _
                                  warnings As Collection)
    On Error Resume Next
    ' Skip if already flagged off_slide (caller logic) — same shape would double-warn
    Dim L As Double: L = sh.Left
    Dim t As Double: t = sh.Top
    Dim w As Double: w = sh.Width
    Dim h As Double: h = sh.Height
    If L < -SLIDE_BOUNDS_TOL_PT Or L + w > slideW + SLIDE_BOUNDS_TOL_PT Then Exit Sub
    If t < -SLIDE_BOUNDS_TOL_PT Or t + h > slideH + SLIDE_BOUNDS_TOL_PT Then Exit Sub
    ' Skip tiny shapes — margins don't matter for icons/dots
    If w < 30 And h < 30 Then Exit Sub
    Const SAFE_MARGIN As Double = 12
    Dim edges As String: edges = ""
    If L < SAFE_MARGIN Then edges = edges & "left "
    If t < SAFE_MARGIN Then edges = edges & "top "
    If slideW - (L + w) < SAFE_MARGIN Then edges = edges & "right "
    If slideH - (t + h) < SAFE_MARGIN Then edges = edges & "bottom "
    If Len(edges) > 0 Then
        AddWarning warnings, "info", "cramped_to_edge", slideNum, sh.Id, _
            sh.Name & " is within " & SAFE_MARGIN & "pt of " & Trim(edges) & "edge(s)", _
            "move inward by ~" & SAFE_MARGIN & "pt for breathing room (move_shape or set_pos)"
    End If
End Sub

' ---- Check 16: orphan / free-floating connector -------------------------
Private Sub CheckOrphanConnector(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    If Not sh.Connector Then Exit Sub
    Dim beginOk As Boolean: beginOk = sh.ConnectorFormat.BeginConnected
    Dim endOk As Boolean: endOk = sh.ConnectorFormat.EndConnected
    If Not (beginOk And endOk) Then
        Dim issue As String
        If Not beginOk And Not endOk Then
            issue = "both endpoints"
        ElseIf Not beginOk Then
            issue = "begin endpoint"
        Else
            issue = "end endpoint"
        End If
        AddWarning warnings, "warn", "orphan_connector", slideNum, sh.Id, _
            sh.Name & " is a connector with " & issue & " not attached to a shape", _
            "delete_shape or reconnect_connector slide=" & slideNum & " shape_id=" & sh.Id
    End If
End Sub

' ---- Check 17: oversized body font (>40pt) on non-title shape -----------
Private Sub CheckHugeBodyFont(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    If sh.HasTable Or sh.HasChart Then Exit Sub
    If Not sh.HasTextFrame Then Exit Sub
    If Not sh.TextFrame.HasText Then Exit Sub
    ' Skip if shape is a title placeholder — titles are SUPPOSED to be large
    If sh.Type = msoPlaceholder Then
        If sh.PlaceholderFormat.Type = ppPlaceholderTitle Or _
           sh.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then Exit Sub
    End If
    Dim sz As Double: sz = sh.TextFrame.TextRange.Font.Size
    If sz > 40 Then
        AddWarning warnings, "info", "very_large_body_font", slideNum, sh.Id, _
            sh.Name & " uses " & Format(sz, "0") & "pt font on a non-title shape — likely meant to be smaller", _
            "set_font_size slide=" & slideNum & " shape_id=" & sh.Id & " value=14 (or appropriate body size)"
    End If
End Sub

' ---- Check 18: placeholder text never replaced ('Click to add ...') -----
Private Sub CheckPlaceholderText(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    If Not sh.HasTextFrame Then Exit Sub
    If Not sh.TextFrame.HasText Then Exit Sub
    Dim t As String: t = LCase(sh.TextFrame.TextRange.Text)
    If InStr(t, "click to add") > 0 Or InStr(t, "click here to add") > 0 Then
        AddWarning warnings, "warn", "placeholder_text_present", slideNum, sh.Id, _
            sh.Name & " still contains default placeholder prompt text", _
            "set_text or set_paragraph_text to replace the placeholder"
    End If
End Sub

' ---- Check 19: trailing whitespace in text shapes -----------------------
Private Sub CheckTrailingWhitespace(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    If sh.HasTable Or sh.HasChart Then Exit Sub
    If Not sh.HasTextFrame Then Exit Sub
    If Not sh.TextFrame.HasText Then Exit Sub
    Dim t As String: t = sh.TextFrame.TextRange.Text
    ' Strip paragraph terminators first
    Do While Len(t) > 0 And (Right(t, 1) = Chr(13) Or Right(t, 1) = Chr(10))
        t = Left(t, Len(t) - 1)
    Loop
    If Len(t) >= 3 Then
        If Right(t, 3) = "   " Then
            AddWarning warnings, "info", "trailing_whitespace", slideNum, sh.Id, _
                sh.Name & " text ends with 3+ trailing spaces (leftover from edits)", _
                "find_replace_text or set_text to clean trailing whitespace"
        End If
    End If
End Sub

' ---- Check 20: broken internal hyperlink (#slide:N out of range) --------
Private Sub CheckBrokenInternalHyperlink(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    Dim deckCount As Long: deckCount = ActivePresentation.Slides.Count
    ' Shape-level click action
    Dim url As String: url = sh.ActionSettings(1).Hyperlink.Address
    If Len(url) > 0 Then
        If LCase(Left(url, 7)) = "#slide:" Then
            Dim n As Long: n = Val(Mid(url, 8))
            If n < 1 Or n > deckCount Then
                AddWarning warnings, "warn", "broken_internal_hyperlink", slideNum, sh.Id, _
                    sh.Name & " has a shape-level hyperlink to #slide:" & n & _
                    " but deck only has " & deckCount & " slides", _
                    "set_shape_hyperlink slide=" & slideNum & " shape_id=" & sh.Id & _
                    " value=#slide:<valid N> (or """" to clear)"
            End If
        End If
    End If
    ' Run-level hyperlinks
    If Not sh.HasTextFrame Then Exit Sub
    If Not sh.TextFrame.HasText Then Exit Sub
    Dim paraN As Long: paraN = sh.TextFrame.TextRange.Paragraphs().Count
    Dim p As Long
    For p = 1 To paraN
        Dim para As TextRange: Set para = sh.TextFrame.TextRange.Paragraphs(p)
        Dim r As Long
        For r = 1 To para.Runs.Count
            Dim ru As TextRange: Set ru = para.Runs(r)
            Dim u As String: u = ru.ActionSettings(1).Hyperlink.Address
            If LCase(Left(u, 7)) = "#slide:" Then
                Dim slideStr As String: slideStr = Mid(u, 8)
                Dim slideOk As Boolean: slideOk = (Len(slideStr) > 0)
                Dim ci As Long
                For ci = 1 To Len(slideStr)
                    If Mid(slideStr, ci, 1) < "0" Or Mid(slideStr, ci, 1) > "9" Then
                        slideOk = False
                        Exit For
                    End If
                Next ci
                If Not slideOk Then GoTo NextHyperlink
                Dim n2 As Long: n2 = CLng(slideStr)
                If n2 < 1 Or n2 > deckCount Then
                    AddWarning warnings, "warn", "broken_internal_hyperlink", slideNum, sh.Id, _
                        sh.Name & " paragraph " & (p - 1) & " run " & (r - 1) & _
                        " hyperlinks to #slide:" & n2 & " but deck has only " & deckCount & " slides", _
                        "set_run_hyperlink slide=" & slideNum & " shape_id=" & sh.Id & _
                        " paragraph_index=" & (p - 1) & " run_index=" & (r - 1) & " value=""""  (clear)"
                End If
            End If
NextHyperlink:
        Next r
    Next p
End Sub

' ---- Check 21: duplicate text content (two shapes same text on slide) ---
Private Sub TrackDuplicateText(sh As Shape, slideNum As Long, _
                                textMap As Object, warnings As Collection)
    On Error Resume Next
    If sh.HasTable Or sh.HasChart Then Exit Sub
    If Not sh.HasTextFrame Then Exit Sub
    If Not sh.TextFrame.HasText Then Exit Sub
    Dim t As String: t = Trim(sh.TextFrame.TextRange.Text)
    If Len(t) < 8 Then Exit Sub   ' skip short labels like "Q1" / "Yes"
    Dim key As String: key = LCase(t)
    If textMap.Exists(key) Then
        Dim otherId As Long: otherId = textMap(key)
        AddWarning warnings, "info", "duplicate_text_content", slideNum, sh.Id, _
            sh.Name & " contains the same text as shape Id " & otherId & _
            " on this slide (possible accidental duplicate)", _
            "delete one of the shapes or differentiate the text"
    Else
        textMap(key) = sh.Id
    End If
End Sub

' ---- Check 22: chart series still named 'Series 1' / 'Series 2' --------
Private Sub CheckChartDefaultSeriesNames(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    Dim ch As Object: Set ch = sh.Chart
    Dim seriesCount As Long: seriesCount = ch.SeriesCollection.Count
    If seriesCount > MAX_CHART_SERIES_FOR_DEEP Then Exit Sub
    Dim s As Long
    For s = 1 To seriesCount
        Dim nm As String: nm = ch.SeriesCollection(s).Name
        If LCase(Left(nm, 7)) = "series " Then
            AddWarning warnings, "info", "chart_default_series_name", slideNum, sh.Id, _
                sh.Name & " series " & s & " still named '" & nm & "' (default placeholder)", _
                "set_series_name slide=" & slideNum & " shape_id=" & sh.Id & _
                " series_index=" & s & " value=""<descriptive name>"""
        End If
    Next s
End Sub

' ---- Check 23: chart with 1 series but legend visible (pointless) ------
Private Sub CheckChartLegendPointless(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    Dim ch As Object: Set ch = sh.Chart
    If ch.SeriesCollection.Count = 1 And ch.HasLegend Then
        AddWarning warnings, "info", "chart_pointless_legend", slideNum, sh.Id, _
            sh.Name & " has only 1 series but the legend is visible (legend conveys no extra info)", _
            "set_chart_legend slide=" & slideNum & " shape_id=" & sh.Id & " props.visible=false"
    End If
End Sub

' ---- Check 24: chart value-axis range mismatched with data magnitude ---
Private Sub CheckChartAxisUnitsMismatch(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    Dim ch As Object: Set ch = sh.Chart
    Dim seriesCount As Long: seriesCount = ch.SeriesCollection.Count
    If seriesCount = 0 Or seriesCount > MAX_CHART_SERIES_FOR_DEEP Then Exit Sub
    Dim maxVal As Double: maxVal = -1E+30
    Dim minVal As Double: minVal = 1E+30
    Dim sawData As Boolean: sawData = False
    Dim s As Long
    For s = 1 To seriesCount
        Dim vals As Variant: vals = ch.SeriesCollection(s).Values
        If IsArray(vals) Then
            Dim i As Long
            For i = LBound(vals) To UBound(vals)
                If IsNumeric(vals(i)) Then
                    sawData = True
                    Dim v As Double: v = CDbl(vals(i))
                    If v > maxVal Then maxVal = v
                    If v < minVal Then minVal = v
                End If
            Next i
        End If
    Next s
    If Not sawData Then Exit Sub
    Dim axMin As Double: axMin = ch.Axes(2).MinimumScale
    Dim axMax As Double: axMax = ch.Axes(2).MaximumScale
    ' Flag if axis max < 30% of data max (data clipped) OR axis max > 10x data max (wasted space)
    If maxVal > 0 Then
        If axMax > 0 And axMax < maxVal * 0.3 Then
            AddWarning warnings, "warn", "chart_axis_clips_data", slideNum, sh.Id, _
                sh.Name & " value axis max (" & Format(axMax, "0.#") & ") is below data max (" & _
                Format(maxVal, "0.#") & ") — bars will be clipped", _
                "set_chart_axis slide=" & slideNum & " shape_id=" & sh.Id & _
                " axis=y props.max=" & Format(maxVal * 1.1, "0")
        ElseIf axMax > maxVal * 10 Then
            AddWarning warnings, "info", "chart_axis_excess_headroom", slideNum, sh.Id, _
                sh.Name & " value axis max (" & Format(axMax, "0") & ") is >10x data max (" & _
                Format(maxVal, "0") & ") — most of plot area is empty", _
                "set_chart_axis slide=" & slideNum & " shape_id=" & sh.Id & _
                " axis=y props.max=" & Format(maxVal * 1.1, "0")
        End If
    End If
End Sub

' ---- Check 25: table column widths sum exceeds table width --------------
Private Sub CheckTableColumnOverflow(sh As Shape, slideNum As Long, warnings As Collection)
    On Error Resume Next
    Dim tbl As Table: Set tbl = sh.Table
    Dim sumW As Double: sumW = 0
    Dim c As Long
    For c = 1 To tbl.Columns.Count
        sumW = sumW + tbl.Columns(c).Width
    Next c
    If sumW > sh.Width + 2 Then
        AddWarning warnings, "warn", "table_column_overflow", slideNum, sh.Id, _
            sh.Name & " column widths sum to " & Format(sumW, "0") & "pt but table is only " & _
            Format(sh.Width, "0") & "pt wide", _
            "set_table_col_width on individual columns to fit, or resize the table"
    End If
End Sub

' ---- Check 26: slide has no title or title placeholder is empty --------
Private Sub CheckSlideTitle(sl As Slide, slideNum As Long, warnings As Collection)
    On Error Resume Next
    ' Skip section-header / blank-layout slides where missing title is intentional.
    Dim layoutType As Long: layoutType = sl.Layout
    ' ppLayoutBlank = 12, ppLayoutSectionHeader = 17 (not always present in enum, skip layout=12)
    If layoutType = 12 Then Exit Sub
    Dim hasTitle As Boolean: hasTitle = False
    Dim ph As Shape
    Dim i As Long
    For i = 1 To sl.Shapes.Placeholders.Count
        Set ph = sl.Shapes.Placeholders(i)
        If ph.PlaceholderFormat.Type = ppPlaceholderTitle Or _
           ph.PlaceholderFormat.Type = ppPlaceholderCenterTitle Then
            hasTitle = True
            ' Title placeholder present; check if it has actual text
            If ph.HasTextFrame Then
                If ph.TextFrame.HasText Then
                    Dim txt As String: txt = Trim(ph.TextFrame.TextRange.Text)
                    If Len(txt) > 0 And InStr(LCase(txt), "click to add") = 0 Then Exit Sub
                End If
            End If
            AddWarning warnings, "info", "slide_empty_title", slideNum, ph.Id, _
                "slide " & slideNum & " title placeholder is empty or contains default prompt", _
                "set_text or set_paragraph_text on the title placeholder"
            Exit Sub
        End If
    Next i
    If Not hasTitle Then
        AddWarning warnings, "info", "slide_no_title", slideNum, 0, _
            "slide " & slideNum & " has no title placeholder (screen readers + TOC rely on it)", _
            "use a layout that includes a title, or add_text_box for a title"
    End If
End Sub

' ---- Check 27: chart fully covered by another shape (Z-order mistake) ---
Private Sub CheckChartCovered(slideNum As Long, chartShapes As Collection, _
                               coveringShapes As Collection, warnings As Collection)
    On Error Resume Next
    Dim ch As Shape, cv As Shape
    Dim i As Long, j As Long
    For i = 1 To chartShapes.Count
        Set ch = chartShapes(i)
        For j = 1 To coveringShapes.Count
            Set cv = coveringShapes(j)
            If cv.ZOrderPosition <= ch.ZOrderPosition Then GoTo NextCv
            ' Bounds containment with 2pt tolerance
            If cv.Left <= ch.Left + 2 And _
               cv.Top <= ch.Top + 2 And _
               (cv.Left + cv.Width) >= (ch.Left + ch.Width) - 2 And _
               (cv.Top + cv.Height) >= (ch.Top + ch.Height) - 2 Then
                AddWarning warnings, "warn", "chart_covered_by_shape", slideNum, ch.Id, _
                    "chart '" & ch.Name & "' is fully covered by shape '" & cv.Name & _
                    "' (Id " & cv.Id & ") which is above it in Z-order", _
                    "z_order slide=" & slideNum & " shape_id=" & cv.Id & " order=back, " & _
                    "or delete the covering shape"
                Exit For
            End If
NextCv:
        Next j
    Next i
End Sub

' ---- Check 28: inconsistent row / column alignment ---------------------
' Finds shapes that look like they're in the same row (similar tops within
' 1-5pt) and warns if their tops aren't pixel-aligned. Same for columns.
Private Sub CheckRowColumnAlignment(sl As Slide, slideNum As Long, warnings As Collection)
    On Error Resume Next
    If sl.Shapes.Count > 30 Then Exit Sub   ' too noisy on busy slides
    Dim tops As Object: Set tops = CreateObject("Scripting.Dictionary")
    Dim lefts As Object: Set lefts = CreateObject("Scripting.Dictionary")
    Dim sh As Shape
    For Each sh In sl.Shapes
        If sh.Type = msoLine Then GoTo NextSh
        If sh.Connector Then GoTo NextSh
        ' Quantize to nearest 0.5pt to detect "almost-aligned" groups
        Dim tkey As String: tkey = CStr(Int(sh.Top * 2) / 2)
        Dim lkey As String: lkey = CStr(Int(sh.Left * 2) / 2)
        If tops.Exists(tkey) Then
            tops(tkey) = tops(tkey) & "," & sh.Id & ":" & sh.Top
        Else
            tops(tkey) = sh.Id & ":" & sh.Top
        End If
        If lefts.Exists(lkey) Then
            lefts(lkey) = lefts(lkey) & "," & sh.Id & ":" & sh.Left
        Else
            lefts(lkey) = sh.Id & ":" & sh.Left
        End If
NextSh:
    Next sh
    ' Look for near-aligned rows: tops differing by 0.1 to 4pt across 3+ shapes
    DetectMisalignmentBand sl, slideNum, "row", warnings
    DetectMisalignmentBand sl, slideNum, "column", warnings
End Sub

' Group shapes by similar Y (or X), flag groups where individual tops (or
' lefts) differ from each other by 1-5pt — looks intentional but isn't aligned.
Private Sub DetectMisalignmentBand(sl As Slide, slideNum As Long, _
                                    axis As String, warnings As Collection)
    On Error Resume Next
    Dim n As Long: n = sl.Shapes.Count
    If n < 3 Or n > 30 Then Exit Sub
    Dim ids() As Long: ReDim ids(1 To n)
    Dim coord() As Double: ReDim coord(1 To n)
    Dim names() As String: ReDim names(1 To n)
    Dim i As Long, count As Long: count = 0
    Dim sh As Shape
    For Each sh In sl.Shapes
        If sh.Type = msoLine Or sh.Connector Then GoTo NextSh2
        count = count + 1
        ids(count) = sh.Id
        names(count) = sh.Name
        If axis = "row" Then
            coord(count) = sh.Top
        Else
            coord(count) = sh.Left
        End If
NextSh2:
    Next sh
    If count < 3 Then Exit Sub
    ' Cluster shapes whose coord is within 5pt of each other; if cluster has
    ' >=3 members AND any pair differs by >0.5pt, flag the band.
    Dim flagged As Object: Set flagged = CreateObject("Scripting.Dictionary")
    Dim a As Long, b As Long
    For a = 1 To count - 2
        If flagged.Exists(CStr(ids(a))) Then GoTo NextA
        Dim members As String: members = names(a) & "(" & Format(coord(a), "0.0") & ")"
        Dim memberCount As Long: memberCount = 1
        Dim maxDelta As Double: maxDelta = 0
        For b = a + 1 To count
            If Abs(coord(b) - coord(a)) <= 5 And Abs(coord(b) - coord(a)) > 0.5 Then
                members = members & ", " & names(b) & "(" & Format(coord(b), "0.0") & ")"
                memberCount = memberCount + 1
                If Abs(coord(b) - coord(a)) > maxDelta Then maxDelta = Abs(coord(b) - coord(a))
                flagged(CStr(ids(b))) = 1
            End If
        Next b
        If memberCount >= 3 Then
            flagged(CStr(ids(a))) = 1
            AddWarning warnings, "info", "inconsistent_" & axis & "_alignment", slideNum, ids(a), _
                memberCount & " shapes appear to share a " & axis & " but their " & _
                IIf(axis = "row", "tops", "lefts") & " differ by up to " & _
                Format(maxDelta, "0.#") & "pt: " & members, _
                "align_shapes slide=" & slideNum & " shape_ids=[<ids>] anchor=" & _
                IIf(axis = "row", "top", "left")
        End If
NextA:
    Next a
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

' =============================================================================
' "Fix This" button helpers — read sidecar JSON, format as LLM-ready prompt,
' optionally copy to clipboard. Public so the form button can call them and
' so they're testable from Application.Run.
' =============================================================================

' Read warnings.json sidecar, format as LLM prompt. Returns empty string on
' missing/empty/unparseable file (caller checks Len and handles).
Public Function BuildLLMPromptFromWarnings(Optional warningsPath As String = "") As String
    BuildLLMPromptFromWarnings = ""
    Dim path As String
    If Len(warningsPath) = 0 Then
        path = ActivePresentation.FullName & ".warnings.json"
    Else
        path = warningsPath
    End If

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function

    Dim raw As String
    Dim fnum As Integer: fnum = FreeFile
    Open path For Input As #fnum
    If LOF(fnum) > 0 Then raw = Input$(LOF(fnum), fnum)
    Close #fnum

    Dim parsed As Object
    On Error Resume Next
    Set parsed = modJSON.ParseJson(raw)
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    If Not parsed.Exists("warnings") Then Exit Function

    Dim warns As Object: Set warns = parsed("warnings")
    If warns.Count = 0 Then Exit Function

    Dim sb As String
    sb = "Decko's verification loop found " & warns.Count & " quality issue(s) " & _
         "in the active deck. Each warning has a SUGGESTION field with the " & _
         "literal action(s) to use as a fix. Return a single JSON object " & _
         "{""actions"":[...]} that resolves every warning. Use the suggestions " & _
         "as guidance; combine related fixes where it makes sense." & vbCrLf & vbCrLf
    sb = sb & "WARNINGS:" & vbCrLf

    Dim i As Long
    For i = 1 To warns.Count
        Dim w As Object: Set w = warns(i)
        Dim head As String
        head = "[" & CStr(w("severity")) & "] " & CStr(w("kind")) & " (slide " & CStr(w("slide"))
        If CLng(w("shape_id")) <> 0 Then head = head & ", shape_id " & CStr(w("shape_id"))
        head = head & ")"
        sb = sb & vbCrLf & head & vbCrLf
        sb = sb & "  ISSUE: " & CStr(w("message")) & vbCrLf
        sb = sb & "  SUGGESTION: " & CStr(w("suggestion")) & vbCrLf
    Next i
    BuildLLMPromptFromWarnings = sb
End Function

' One-call helper: build prompt + put on clipboard. Returns warning count
' (0 if nothing to copy). Form's btnFixThis_Click calls this.
Public Function CopyWarningsPromptToClipboard() As Long
    Dim prompt As String: prompt = BuildLLMPromptFromWarnings()
    If Len(prompt) = 0 Then
        CopyWarningsPromptToClipboard = 0
        Exit Function
    End If
    Dim dobj As MSForms.DataObject
    Set dobj = New MSForms.DataObject
    dobj.SetText prompt
    dobj.PutInClipboard
    ' Count warnings cheaply: scan for "WARNINGS:" then count lines starting with "["
    Dim n As Long: n = 0
    Dim pos As Long: pos = 1
    Do
        pos = InStr(pos, prompt, vbCrLf & "[")
        If pos = 0 Then Exit Do
        n = n + 1
        pos = pos + 1
    Loop
    CopyWarningsPromptToClipboard = n
End Function
