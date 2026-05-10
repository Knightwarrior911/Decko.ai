Attribute VB_Name = "modActionsChart"
Option Explicit

' Create a native PowerPoint chart with categories + series data.
' chart_type: columnclustered, columnstacked, barclustered, barstacked, line,
'             linemarkers, pie, area, scatter, doughnut.
' categories: Collection of strings (e.g. ["2022","2023","2024","2025"])
' series:     Collection of dicts, each with keys:
'               name (string), values (Collection of numbers), color (optional hex)
Public Sub Do_add_chart(slideNum As Long, chartType As String, _
                         leftPt As Single, topPt As Single, _
                         widthPt As Single, heightPt As Single, _
                         ByVal categories As Object, ByVal series As Object, _
                         Optional refName As String = "", _
                         Optional showLegend As Boolean = True, _
                         Optional showValues As Boolean = False, _
                         Optional titleText As String = "", _
                         Optional cleanStyle As Boolean = False, _
                         Optional valueFormat As String = "")
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 11010, "Do_add_chart", "slide_out_of_range"
    End If
    Dim sl As Slide: Set sl = pres.Slides(slideNum)
    Dim chartTypeNum As Long: chartTypeNum = ChartTypeFromName(chartType)

    Dim sh As Shape
    Set sh = sl.Shapes.AddChart2(-1, chartTypeNum, leftPt, topPt, widthPt, heightPt, True)
    If Len(refName) > 0 Then sh.Name = refName
    Dim ch As Chart: Set ch = sh.Chart

    Dim catCount As Long: catCount = categories.Count
    Dim seriesCount As Long: seriesCount = series.Count

    ' Build category array
    Dim catArr() As Variant
    ReDim catArr(1 To catCount)
    Dim r As Long
    For r = 1 To catCount
        catArr(r) = CStr(categories(r))
    Next r

    ' Detect chart types that don't support standard SeriesCollection writes:
    ' waterfall (119), histogram (118), pareto (122), boxwhisker (121),
    ' treemap (117), sunburst (120), funnel (123). For these, the chart is
    ' created with AddChart2 default placeholder data and series isn't manipulated.
    ' Caller can supply data via ChartData approach or live with placeholders.
    Dim isSpecial As Boolean
    isSpecial = (chartTypeNum = 117 Or chartTypeNum = 118 Or chartTypeNum = 119 Or _
                 chartTypeNum = 120 Or chartTypeNum = 121 Or chartTypeNum = 122 Or _
                 chartTypeNum = 123)

    Dim s As Long
    If Not isSpecial Then
        Dim existingCount As Long: existingCount = ch.SeriesCollection.Count
        Do While ch.SeriesCollection.Count > seriesCount
            ch.SeriesCollection(ch.SeriesCollection.Count).Delete
        Loop
        Do While ch.SeriesCollection.Count < seriesCount
            ch.SeriesCollection.NewSeries
        Loop
        For s = 1 To seriesCount
            Dim si As Object: Set si = series(s)
            Dim valsCol As Object: Set valsCol = si("values")
            Dim valArr() As Variant
            ReDim valArr(1 To catCount)
            For r = 1 To catCount
                valArr(r) = CDbl(valsCol(r))
            Next r
            With ch.SeriesCollection(s)
                .Name = CStr(si("name"))
                .Values = valArr
                .XValues = catArr
            End With
        Next s
    End If

    ch.HasLegend = showLegend
    If Len(titleText) > 0 Then
        ch.HasTitle = True
        ch.ChartTitle.Text = titleText
    Else
        ch.HasTitle = False
    End If

    ' Apply colors per series if provided
    For s = 1 To seriesCount
        Set si = series(s)
        If si.Exists("color") Then
            Dim hex As String: hex = CStr(si("color"))
            On Error Resume Next
            ch.SeriesCollection(s).Format.Fill.ForeColor.RGB = modActions.HexToRgb(hex)
            ch.SeriesCollection(s).Format.Line.ForeColor.RGB = modActions.HexToRgb(hex)
            On Error GoTo 0
        End If
    Next s

    ' Show data labels if requested or clean style demands them
    If showValues Or cleanStyle Then
        On Error Resume Next
        For s = 1 To seriesCount
            ch.SeriesCollection(s).HasDataLabels = True
            If cleanStyle Then
                ' xlLabelPositionOutsideEnd = 0 — above bars / outside markers
                ch.SeriesCollection(s).DataLabels.Position = 0
                ch.SeriesCollection(s).DataLabels.Font.Italic = True
            End If
            If Len(valueFormat) > 0 Then
                ch.SeriesCollection(s).DataLabels.NumberFormat = valueFormat
            End If
        Next s
        On Error GoTo 0
    End If

    ' Clean style: hide y-axis labels + line (keep scale active so series render),
    ' delete gridlines, hide chart/plot borders & fills.
    ' DO NOT use HasAxis=False — that removes the axis entirely, killing the scale
    ' that the bars depend on for sizing (especially in combo charts).
    If cleanStyle Then
        On Error Resume Next
        ' Primary value axis — hide visual but keep scale
        ch.Axes(2, 1).TickLabelPosition = -4142     ' xlNone
        ch.Axes(2, 1).Format.Line.Visible = msoFalse
        ch.Axes(2, 1).MajorGridlines.Delete
        ' Secondary value axis (may not exist; On Error swallows)
        ch.Axes(2, 2).TickLabelPosition = -4142
        ch.Axes(2, 2).Format.Line.Visible = msoFalse
        ch.Axes(2, 2).MajorGridlines.Delete
        ' Category axis gridlines
        ch.Axes(1).MajorGridlines.Delete
        ' Frame
        ch.ChartArea.Format.Line.Visible = msoFalse
        ch.PlotArea.Format.Line.Visible = msoFalse
        ch.ChartArea.Format.Fill.Visible = msoFalse
        ch.PlotArea.Format.Fill.Visible = msoFalse
        On Error GoTo 0
    End If
End Sub

' Convert 1-based column number to Excel letter (1=A, 2=B, ..., 27=AA)
Private Function ColLetter(n As Long) As String
    Dim s As String
    Do While n > 0
        Dim rem_ As Long: rem_ = (n - 1) Mod 26
        s = Chr(65 + rem_) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function

' --- Granular chart-element actions: parity with PowerPoint's chart UI -----

' Chart-group / chart-level visual properties.
'   props (Object): any subset of:
'     gap_width (0-500)        — space between category groups (column/bar charts)
'     overlap (-100 to 100)    — overlap between series in same category
'     bar_shape (string)       — "box", "cone", "cone_to_max", "cylinder", "pyramid", "pyramid_to_max"
'     vary_by_categories (bool)— color each point in a single-series chart
'     reverse_categories (bool)— reverse category-axis plot order
'     reverse_series (bool)    — reverse series plot order
'     scale_type (string)      — "linear" or "logarithmic" for value axis
'     doughnut_hole_size (10-90)— inner-radius % for doughnut charts
Public Sub Do_set_chart_format(slideNum As Long, shapeId As Long, ByVal props As Object)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11050, "Do_set_chart_format", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11051, "Do_set_chart_format", "not a chart"
    Dim ch As Object: Set ch = sh.Chart
    On Error Resume Next
    Dim cg As Object: Set cg = ch.ChartGroups(1)
    If props.Exists("gap_width") Then cg.GapWidth = modActions.ToLong(props("gap_width"))
    If props.Exists("overlap") Then cg.Overlap = modActions.ToLong(props("overlap"))
    If props.Exists("vary_by_categories") Then cg.VaryByCategories = modActions.ToBool(props("vary_by_categories"))
    If props.Exists("doughnut_hole_size") Then cg.DoughnutHoleSize = modActions.ToLong(props("doughnut_hole_size"))
    If props.Exists("bar_shape") Then
        Select Case LCase(CStr(props("bar_shape")))
            Case "box":            cg.BarShape = 0
            Case "pyramid":        cg.BarShape = 1
            Case "pyramid_to_max": cg.BarShape = 2
            Case "cylinder":       cg.BarShape = 3
            Case "cone":           cg.BarShape = 4
            Case "cone_to_max":    cg.BarShape = 5
        End Select
    End If
    If props.Exists("reverse_categories") Then
        ch.Axes(1).ReversePlotOrder = modActions.ToBool(props("reverse_categories"))
    End If
    If props.Exists("reverse_series") Then
        ch.Axes(2).ReversePlotOrder = modActions.ToBool(props("reverse_series"))
    End If
    If props.Exists("scale_type") Then
        Select Case LCase(CStr(props("scale_type")))
            Case "linear":      ch.Axes(2).ScaleType = -4132    ' xlScaleLinear
            Case "logarithmic": ch.Axes(2).ScaleType = -4133    ' xlScaleLogarithmic
        End Select
    End If
    ' Drop lines (line charts) - vertical line from each marker to category axis
    If props.Exists("drop_lines") Then cg.HasDropLines = modActions.ToBool(props("drop_lines"))
    ' Hi-lo lines (line charts with multiple series) - vertical line between max/min markers
    If props.Exists("hi_lo_lines") Then cg.HasHiLoLines = modActions.ToBool(props("hi_lo_lines"))
    ' Up-down bars (line charts with 2 series) - boxes showing diff between series
    If props.Exists("up_down_bars") Then cg.HasUpDownBars = modActions.ToBool(props("up_down_bars"))
    On Error GoTo 0
End Sub

' Add a trendline to a chart series.
'   seriesIndex: 1-based
'   props (Object):
'     type (string)             — "linear" | "log" | "polynomial" | "power" |
'                                 "exponential" | "moving_avg"
'     order (long)              — for polynomial (2-6)
'     period (long)             — for moving_avg (2..)
'     forward (number)          — periods to extend forward
'     backward (number)         — periods to extend backward
'     intercept (number)        — y-intercept
'     display_equation (bool)
'     display_r_squared (bool)
'     name (string)             — custom trendline name
'     color (hex), weight (number), dash (solid/dash/dot/...)
Public Sub Do_add_chart_trendline(slideNum As Long, shapeId As Long, _
                                    seriesIndex As Long, ByVal props As Object)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11060, "Do_add_chart_trendline", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11061, "Do_add_chart_trendline", "not a chart"
    Dim ch As Object: Set ch = sh.Chart
    If seriesIndex < 1 Or seriesIndex > ch.SeriesCollection.Count Then
        Err.Raise vbObjectError + 11062, "Do_add_chart_trendline", _
                  "series_index out of range 1.." & ch.SeriesCollection.Count
    End If
    Dim ser As Object: Set ser = ch.SeriesCollection(seriesIndex)
    Dim trType As Long: trType = -4132   ' xlLinear default
    If props.Exists("type") Then
        Select Case LCase(CStr(props("type")))
            Case "linear":      trType = -4132    ' xlLinear
            Case "log":         trType = -4133    ' xlLogarithmic
            Case "polynomial":  trType = 3        ' xlPolynomial
            Case "power":       trType = 4        ' xlPower
            Case "exponential": trType = 5        ' xlExponential
            Case "moving_avg":  trType = 6        ' xlMovingAvg
        End Select
    End If
    Dim tr As Object: Set tr = ser.Trendlines.Add(Type:=trType)
    On Error Resume Next
    If props.Exists("order") Then tr.Order = modActions.ToLong(props("order"))
    If props.Exists("period") Then tr.Period = modActions.ToLong(props("period"))
    If props.Exists("forward") Then tr.Forward = CDbl(props("forward"))
    If props.Exists("backward") Then tr.Backward = CDbl(props("backward"))
    If props.Exists("intercept") Then
        tr.InterceptIsAuto = False
        tr.Intercept = CDbl(props("intercept"))
    End If
    If props.Exists("display_equation") Then tr.DisplayEquation = modActions.ToBool(props("display_equation"))
    If props.Exists("display_r_squared") Then tr.DisplayRSquared = modActions.ToBool(props("display_r_squared"))
    If props.Exists("name") Then
        tr.NameIsAuto = False
        tr.Name = CStr(props("name"))
    End If
    If props.Exists("color") Then
        tr.Format.Line.ForeColor.RGB = modActions.HexToRgb(CStr(props("color")))
    End If
    If props.Exists("weight") Then tr.Format.Line.Weight = CDbl(props("weight"))
    If props.Exists("dash") Then
        Select Case LCase(CStr(props("dash")))
            Case "solid":         tr.Format.Line.DashStyle = msoLineSolid
            Case "dash":          tr.Format.Line.DashStyle = msoLineDash
            Case "dot":           tr.Format.Line.DashStyle = msoLineSquareDot
            Case "round_dot":     tr.Format.Line.DashStyle = msoLineRoundDot
            Case "dash_dot":      tr.Format.Line.DashStyle = msoLineDashDot
            Case "long_dash":     tr.Format.Line.DashStyle = msoLineLongDash
            Case "long_dash_dot": tr.Format.Line.DashStyle = msoLineLongDashDot
        End Select
    End If
    On Error GoTo 0
End Sub

' Add error bars to a chart series.
'   props (Object):
'     direction (string)        — "x" | "y"
'     include (string)          — "both" | "plus" | "minus"
'     type (string)             — "fixed" | "percent" | "stdev" | "stderr" | "custom"
'     amount (number)           — value for fixed/percent/stdev
'     end_style (string)        — "cap" | "no_cap"
'     color (hex), weight (number)
Public Sub Do_set_chart_error_bars(slideNum As Long, shapeId As Long, _
                                     seriesIndex As Long, ByVal props As Object)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11070, "Do_set_chart_error_bars", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11071, "Do_set_chart_error_bars", "not a chart"
    Dim ch As Object: Set ch = sh.Chart
    Dim ser As Object: Set ser = ch.SeriesCollection(seriesIndex)

    Dim direction As Long: direction = 2   ' xlY
    Dim include As Long: include = 1       ' xlBoth
    Dim ebType As Long: ebType = 1         ' xlErrorBarTypeFixedValue
    Dim amount As Double: amount = 1
    Dim minusValues As Double: minusValues = 1

    If props.Exists("direction") Then
        Select Case LCase(CStr(props("direction")))
            Case "x": direction = 1   ' xlX
            Case "y": direction = 2   ' xlY
        End Select
    End If
    If props.Exists("include") Then
        Select Case LCase(CStr(props("include")))
            Case "both":  include = 1   ' xlErrorBarIncludeBoth
            Case "plus":  include = 2   ' xlErrorBarIncludePlusValues
            Case "minus": include = 3   ' xlErrorBarIncludeMinusValues
        End Select
    End If
    If props.Exists("type") Then
        Select Case LCase(CStr(props("type")))
            Case "fixed":   ebType = 1   ' xlErrorBarTypeFixedValue
            Case "percent": ebType = 2   ' xlErrorBarTypePercent
            Case "stdev":   ebType = 3   ' xlErrorBarTypeStDev
            Case "stderr":  ebType = 4   ' xlErrorBarTypeStError
            Case "custom":  ebType = -4114  ' xlErrorBarTypeCustom
        End Select
    End If
    If props.Exists("amount") Then amount = CDbl(props("amount"))

    On Error Resume Next
    ser.HasErrorBars = True
    ser.ErrorBar Direction:=direction, Include:=include, Type:=ebType, Amount:=amount
    ' Force error bars visible — defaults are sometimes hidden after method call
    ser.ErrorBars.Format.Line.Visible = msoTrue
    ser.ErrorBars.Format.Line.Weight = 1.5
    ser.ErrorBars.Format.Line.ForeColor.RGB = modActions.HexToRgb("#000000")
    If props.Exists("end_style") Then
        Select Case LCase(CStr(props("end_style")))
            Case "no_cap": ser.ErrorBars.EndStyle = 2   ' xlNoCap
            Case "cap":    ser.ErrorBars.EndStyle = 1   ' xlCap
        End Select
    End If
    If props.Exists("color") Then
        ser.ErrorBars.Format.Line.ForeColor.RGB = modActions.HexToRgb(CStr(props("color")))
    End If
    If props.Exists("weight") Then ser.ErrorBars.Format.Line.Weight = CDbl(props("weight"))
    On Error GoTo 0
End Sub



' Set properties on a chart axis: visibility, label position, line, font, format, scale.
'   axis: "x" / "y" / "category" / "value" — selects axis
'   props (Object): any subset of:
'     visible (bool)                — show/hide entire axis
'     line_visible (bool)           — show/hide axis line only
'     tick_label_position (string)  — "low" | "high" | "next_to_axis" | "none"
'     label_color (hex), label_size (long), label_bold (bool), label_italic (bool)
'     number_format (string)        — Excel format e.g. "0.0%" or '0"%";(0)"%"'
'     min, max, major_unit, minor_unit (number) — value-axis scale
'     title (string)                — axis title text
Public Sub Do_set_chart_axis(slideNum As Long, shapeId As Long, _
                              axis As String, ByVal props As Object)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11020, "Do_set_chart_axis", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11021, "Do_set_chart_axis", "not a chart"
    Dim ch As Object: Set ch = sh.Chart

    Dim axNum As Long
    Dim axGroup As Long: axGroup = 1   ' default xlPrimary
    Select Case LCase(Trim(axis))
        Case "x", "category", "xlcategory":  axNum = 1   ' xlCategory
        Case "y", "value", "xlvalue":        axNum = 2   ' xlValue (primary)
        Case "y2", "value_secondary", "secondary":
            axNum = 2: axGroup = 2                       ' xlValue secondary
        Case "x2", "category_secondary":
            axNum = 1: axGroup = 2
        Case Else: Err.Raise vbObjectError + 11022, "Do_set_chart_axis", _
                              "axis must be x/y/category/value/y2/secondary, got: " & axis
    End Select

    On Error Resume Next
    If props.Exists("visible") Then
        ch.HasAxis(axNum, axGroup) = modActions.ToBool(props("visible"))
    End If
    Dim ax As Object: Set ax = ch.Axes(axNum, axGroup)
    If ax Is Nothing Then Exit Sub

    If props.Exists("line_visible") Then
        If modActions.ToBool(props("line_visible")) Then
            ax.Format.Line.Visible = msoTrue
        Else
            ax.Format.Line.Visible = msoFalse
        End If
    End If
    If props.Exists("tick_label_position") Then
        ' xlLow=-4134, xlHigh=-4127, xlNextToAxis=4, xlNone=-4142
        Dim posKey As String: posKey = LCase(CStr(props("tick_label_position")))
        Select Case posKey
            Case "low":           ax.TickLabelPosition = -4134
            Case "high":          ax.TickLabelPosition = -4127
            Case "next_to_axis":  ax.TickLabelPosition = 4
            Case "none":          ax.TickLabelPosition = -4142
        End Select
    End If
    If props.Exists("label_color") Then
        ax.TickLabels.Font.Color.RGB = modActions.HexToRgb(CStr(props("label_color")))
    End If
    If props.Exists("label_size") Then
        ax.TickLabels.Font.Size = modActions.ToLong(props("label_size"))
    End If
    If props.Exists("label_bold") Then
        ax.TickLabels.Font.Bold = modActions.ToBool(props("label_bold"))
    End If
    If props.Exists("label_italic") Then
        ax.TickLabels.Font.Italic = modActions.ToBool(props("label_italic"))
    End If
    If props.Exists("number_format") Then
        ax.TickLabels.NumberFormat = CStr(props("number_format"))
    End If
    If props.Exists("min") Then
        ax.MinimumScale = CDbl(props("min"))
    End If
    If props.Exists("max") Then
        ax.MaximumScale = CDbl(props("max"))
    End If
    If props.Exists("major_unit") Then
        ax.MajorUnit = CDbl(props("major_unit"))
    End If
    If props.Exists("minor_unit") Then
        ax.MinorUnit = CDbl(props("minor_unit"))
    End If
    If props.Exists("title") Then
        ax.HasTitle = True
        ax.AxisTitle.Text = CStr(props("title"))
    End If
    ' Tick label rotation in degrees (-90 to 90)
    If props.Exists("label_rotation") Then
        ax.TickLabels.Orientation = modActions.ToLong(props("label_rotation"))
    End If
    ' Major tick mark style: "outside" | "inside" | "cross" | "none"
    If props.Exists("major_tick_mark") Then
        Select Case LCase(CStr(props("major_tick_mark")))
            Case "outside": ax.MajorTickMark = 4   ' xlTickMarkOutside
            Case "inside":  ax.MajorTickMark = 2   ' xlTickMarkInside
            Case "cross":   ax.MajorTickMark = 3   ' xlTickMarkCross
            Case "none":    ax.MajorTickMark = -4142 ' xlTickMarkNone
        End Select
    End If
    On Error GoTo 0
End Sub

' Set properties on a single series within a chart.
'   series_index: 1-based
'   props (Object): any subset of:
'     name (string), fill (hex), line_color (hex), line_weight (number)
'     marker_style (string)  — "circle"/"square"/"triangle"/"diamond"/"x"/"none"
'     marker_size (number)
'     show_labels (bool)
'     label_position (string)  — "outside_end"/"inside_end"/"center"/"above"/"below"/"left"/"right"
'     label_format (string)    — Excel number format
'     label_size, label_bold, label_italic, label_color
'     custom_labels (array)    — per-point label override (string per category)
Public Sub Do_set_chart_series(slideNum As Long, shapeId As Long, _
                                seriesIndex As Long, ByVal props As Object)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11030, "Do_set_chart_series", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11031, "Do_set_chart_series", "not a chart"
    Dim ch As Object: Set ch = sh.Chart
    If seriesIndex < 1 Or seriesIndex > ch.SeriesCollection.Count Then
        Err.Raise vbObjectError + 11032, "Do_set_chart_series", _
                  "series_index out of range 1.." & ch.SeriesCollection.Count
    End If
    Dim ser As Object: Set ser = ch.SeriesCollection(seriesIndex)

    On Error Resume Next
    If props.Exists("name") Then ser.Name = CStr(props("name"))
    If props.Exists("fill") Then
        ser.Format.Fill.Solid
        ser.Format.Fill.ForeColor.RGB = modActions.HexToRgb(CStr(props("fill")))
    End If
    ' Gradient fill — props.gradient_fill = { from, to, direction }
    If props.Exists("gradient_fill") Then
        Dim gf As Object: Set gf = props("gradient_fill")
        Dim gfDir As Long: gfDir = 1   ' msoGradientHorizontal default
        If gf.Exists("direction") Then
            Select Case LCase(CStr(gf("direction")))
                Case "horizontal":    gfDir = 1   ' msoGradientHorizontal
                Case "vertical":      gfDir = 2   ' msoGradientVertical
                Case "diagonal_up":   gfDir = 3   ' msoGradientDiagonalUp
                Case "diagonal_down": gfDir = 4   ' msoGradientDiagonalDown
                Case "from_corner":   gfDir = 5   ' msoGradientFromCorner
                Case "from_center":   gfDir = 7   ' msoGradientFromCenter
            End Select
        End If
        ser.Format.Fill.TwoColorGradient gfDir, 1
        If gf.Exists("from") Then
            ser.Format.Fill.ForeColor.RGB = modActions.HexToRgb(CStr(gf("from")))
        End If
        If gf.Exists("to") Then
            ser.Format.Fill.BackColor.RGB = modActions.HexToRgb(CStr(gf("to")))
        End If
    End If
    If props.Exists("fill_visible") Then
        If modActions.ToBool(props("fill_visible")) Then
            ser.Format.Fill.Visible = msoTrue
        Else
            ' Force no fill — Format.Fill.Visible alone doesn't override "Automatic".
            ' DO NOT touch Line here — caller uses line_color/line_dash/line_weight for borders
            ser.Format.Fill.Visible = msoFalse
            ser.Format.Fill.Transparency = 1
            ser.Interior.ColorIndex = -4142   ' xlNone (Excel chart API)
        End If
    End If
    If props.Exists("line_color") Then
        ser.Format.Line.ForeColor.RGB = modActions.HexToRgb(CStr(props("line_color")))
    End If
    If props.Exists("line_weight") Then
        ser.Format.Line.Weight = CDbl(props("line_weight"))
    End If
    If props.Exists("line_dash") Then
        Select Case LCase(CStr(props("line_dash")))
            Case "solid":          ser.Format.Line.DashStyle = msoLineSolid
            Case "dash":           ser.Format.Line.DashStyle = msoLineDash
            Case "dot":            ser.Format.Line.DashStyle = msoLineSquareDot
            Case "round_dot":      ser.Format.Line.DashStyle = msoLineRoundDot
            Case "dash_dot":       ser.Format.Line.DashStyle = msoLineDashDot
            Case "long_dash":      ser.Format.Line.DashStyle = msoLineLongDash
            Case "long_dash_dot":  ser.Format.Line.DashStyle = msoLineLongDashDot
        End Select
    End If
    ' Per-series chart type (for combo charts) — line/column/bar/area
    If props.Exists("chart_type") Then
        ser.ChartType = ChartTypeFromName(CStr(props("chart_type")))
    End If
    ' Axis group: 1 = primary (default), 2 = secondary (creates secondary y-axis)
    If props.Exists("axis_group") Then
        Select Case LCase(CStr(props("axis_group")))
            Case "primary":   ser.AxisGroup = 1
            Case "secondary": ser.AxisGroup = 2
        End Select
    End If
    If props.Exists("marker_style") Then
        ' xlMarkerStyle: Circle=8, Square=1, Triangle=3, Diamond=2, X=-4168, None=-4142
        Select Case LCase(CStr(props("marker_style")))
            Case "circle":   ser.MarkerStyle = 8
            Case "square":   ser.MarkerStyle = 1
            Case "triangle": ser.MarkerStyle = 3
            Case "diamond":  ser.MarkerStyle = 2
            Case "x":        ser.MarkerStyle = -4168
            Case "none":     ser.MarkerStyle = -4142
        End Select
    End If
    If props.Exists("marker_size") Then ser.MarkerSize = CLng(props("marker_size"))
    If props.Exists("marker_fill") Then
        ser.MarkerBackgroundColor = modActions.HexToRgb(CStr(props("marker_fill")))
    End If
    If props.Exists("marker_line") Then
        ser.MarkerForegroundColor = modActions.HexToRgb(CStr(props("marker_line")))
    End If
    If props.Exists("show_labels") Then
        ser.HasDataLabels = modActions.ToBool(props("show_labels"))
    End If
    If ser.HasDataLabels Then
        Dim lbls As Object: Set lbls = ser.DataLabels
        If props.Exists("label_position") Then
            Select Case LCase(CStr(props("label_position")))
                Case "outside_end", "above": lbls.Position = 0    ' xlLabelPositionOutsideEnd
                Case "inside_end":           lbls.Position = 3    ' xlLabelPositionInsideEnd
                Case "center":               lbls.Position = -4108 ' xlLabelPositionCenter
                Case "below":                lbls.Position = 1    ' xlLabelPositionBelow
                Case "left":                 lbls.Position = 2    ' xlLabelPositionLeft
                Case "right":                lbls.Position = 4    ' xlLabelPositionRight
            End Select
        End If
        If props.Exists("label_format") Then lbls.NumberFormat = CStr(props("label_format"))
        If props.Exists("label_size") Then
            Dim lcSz As Long: lcSz = modActions.ToLong(props("label_size"))
            lbls.Font.Size = lcSz
            lbls.Format.TextFrame2.TextRange.Font.Size = lcSz
        End If
        If props.Exists("label_bold") Then
            Dim lcBd As Boolean: lcBd = modActions.ToBool(props("label_bold"))
            lbls.Font.Bold = lcBd
            lbls.Format.TextFrame2.TextRange.Font.Bold = lcBd
        End If
        If props.Exists("label_italic") Then
            Dim lcIt As Boolean: lcIt = modActions.ToBool(props("label_italic"))
            lbls.Font.Italic = lcIt
            lbls.Format.TextFrame2.TextRange.Font.Italic = lcIt
        End If
        If props.Exists("label_color") Then
            Dim lcCl As Long: lcCl = modActions.HexToRgb(CStr(props("label_color")))
            lbls.Font.Color.RGB = lcCl
            lbls.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = lcCl
        End If
        ' Label background fill (for tile-style labels in tight stacked segments)
        If props.Exists("label_fill") Then
            lbls.Format.Fill.Solid
            lbls.Format.Fill.ForeColor.RGB = modActions.HexToRgb(CStr(props("label_fill")))
        End If
        If props.Exists("label_fill_visible") Then
            If modActions.ToBool(props("label_fill_visible")) Then
                lbls.Format.Fill.Visible = msoTrue
            Else
                lbls.Format.Fill.Visible = msoFalse
                lbls.Format.Fill.Transparency = 1
                lbls.Interior.ColorIndex = -4142   ' xlNone
            End If
        End If
        ' Label outline border — hide for clean floating-totals look.
        ' "Automatic" border in PowerPoint ignores Format.Line.Visible alone.
        ' Force xlNone via Border.LineStyle, AND fully transparent line as fallback.
        If props.Exists("label_line_visible") Then
            If modActions.ToBool(props("label_line_visible")) Then
                lbls.Format.Line.Visible = msoTrue
            Else
                lbls.Format.Line.Visible = msoFalse
            End If
        End If
    End If
    ' Hide this series' entry from the legend (series stays in chart)
    If props.Exists("hide_from_legend") Then
        If modActions.ToBool(props("hide_from_legend")) Then
            On Error Resume Next
            ch.Legend.LegendEntries(seriesIndex).Delete
            On Error GoTo 0
        End If
    End If
    ' Per-point custom label text override.
    ' Setting Points(p).DataLabel.Text resets the per-point font properties to chart
    ' defaults — re-apply font color/size/italic/fill from props after each set.
    If props.Exists("custom_labels") Then
        Dim customs As Object: Set customs = props("custom_labels")
        Dim p As Long
        For p = 1 To ser.Points.Count
            If p <= customs.Count Then
                Dim labText As String: labText = CStr(customs(p))
                If Len(labText) = 0 Then
                    ser.Points(p).HasDataLabel = False
                Else
                    ser.Points(p).HasDataLabel = True
                    ser.Points(p).DataLabel.Text = labText
                    Dim pl As Object: Set pl = ser.Points(p).DataLabel
                    If props.Exists("label_color") Then
                        Dim lbCol As Long: lbCol = modActions.HexToRgb(CStr(props("label_color")))
                        ' Legacy Font API (older Office versions)
                        pl.Font.Color.RGB = lbCol
                        ' Modern TextFrame2 API — more reliable in current Office
                        pl.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = lbCol
                    End If
                    If props.Exists("label_size") Then
                        Dim lbSize As Long: lbSize = modActions.ToLong(props("label_size"))
                        pl.Font.Size = lbSize
                        pl.Format.TextFrame2.TextRange.Font.Size = lbSize
                    End If
                    If props.Exists("label_italic") Then
                        Dim lbIt As Boolean: lbIt = modActions.ToBool(props("label_italic"))
                        pl.Font.Italic = lbIt
                        pl.Format.TextFrame2.TextRange.Font.Italic = lbIt
                    End If
                    If props.Exists("label_bold") Then
                        Dim lbBd As Boolean: lbBd = modActions.ToBool(props("label_bold"))
                        pl.Font.Bold = lbBd
                        pl.Format.TextFrame2.TextRange.Font.Bold = lbBd
                    End If
                    If props.Exists("label_fill") Then
                        pl.Format.Fill.Solid
                        pl.Format.Fill.ForeColor.RGB = modActions.HexToRgb(CStr(props("label_fill")))
                    End If
                    If props.Exists("label_fill_visible") Then
                        If modActions.ToBool(props("label_fill_visible")) Then
                            pl.Format.Fill.Visible = msoTrue
                        Else
                            ' Force no fill via multiple paths — PowerPoint's "Automatic"
                            ' fill ignores some of these individually
                            pl.Format.Fill.Visible = msoFalse
                            pl.Format.Fill.Transparency = 1
                            pl.Interior.ColorIndex = -4142   ' xlNone (Excel chart API)
                        End If
                    End If
                    If props.Exists("label_line_visible") Then
                        If modActions.ToBool(props("label_line_visible")) Then
                            pl.Format.Line.Visible = msoTrue
                        Else
                            pl.Format.Line.Visible = msoFalse
                        End If
                    End If
                    ' Per-point label color override (e.g. white text on navy bars,
                    ' navy text on light bars in same series)
                    If props.Exists("point_label_colors") Then
                        Dim plc As Object: Set plc = props("point_label_colors")
                        If p <= plc.Count Then
                            Dim plcHex As String: plcHex = CStr(plc(p))
                            If Len(plcHex) > 0 Then
                                Dim plcRgb As Long: plcRgb = modActions.HexToRgb(plcHex)
                                pl.Font.Color.RGB = plcRgb
                                pl.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = plcRgb
                            End If
                        End If
                    End If
                    ' Per-point label position override (positive bars above, negatives below, etc.)
                    If props.Exists("point_label_positions") Then
                        Dim plp As Object: Set plp = props("point_label_positions")
                        If p <= plp.Count Then
                            Dim plpStr As String: plpStr = CStr(plp(p))
                            If Len(plpStr) > 0 Then
                                Select Case LCase(plpStr)
                                    Case "outside_end", "above": pl.Position = 0
                                    Case "inside_end":           pl.Position = 3
                                    Case "inside_base":          pl.Position = 4
                                    Case "center":               pl.Position = -4108
                                    Case "below":                pl.Position = 1
                                    Case "left":                 pl.Position = 2
                                    Case "right":                pl.Position = 4
                                End Select
                            End If
                        End If
                    End If
                End If
            End If
        Next p
    End If
    ' Per-point fill colors — for highlighting individual bars in single-series charts
    If props.Exists("point_fills") Then
        Dim pfills As Object: Set pfills = props("point_fills")
        Dim pp As Long
        For pp = 1 To ser.Points.Count
            If pp <= pfills.Count Then
                Dim pfHex As String: pfHex = CStr(pfills(pp))
                If Len(pfHex) > 0 Then
                    ser.Points(pp).Format.Fill.Solid
                    ser.Points(pp).Format.Fill.ForeColor.RGB = modActions.HexToRgb(pfHex)
                End If
            End If
        Next pp
    End If
    ' Per-point line colors (for line charts) and marker fills
    If props.Exists("point_marker_fills") Then
        Dim pmf As Object: Set pmf = props("point_marker_fills")
        Dim pm As Long
        For pm = 1 To ser.Points.Count
            If pm <= pmf.Count Then
                ser.Points(pm).MarkerBackgroundColor = modActions.HexToRgb(CStr(pmf(pm)))
            End If
        Next pm
    End If
    On Error GoTo 0
End Sub

' Set the entire chart legend's properties.
'   props: visible, position ("right"/"left"/"top"/"bottom"/"corner"), font props
Public Sub Do_set_chart_legend(slideNum As Long, shapeId As Long, ByVal props As Object)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11040, "Do_set_chart_legend", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11041, "Do_set_chart_legend", "not a chart"
    Dim ch As Object: Set ch = sh.Chart
    On Error Resume Next
    If props.Exists("visible") Then ch.HasLegend = modActions.ToBool(props("visible"))
    If ch.HasLegend Then
        Dim lg As Object: Set lg = ch.Legend
        If props.Exists("position") Then
            Select Case LCase(CStr(props("position")))
                Case "right":  lg.Position = -4152
                Case "left":   lg.Position = -4131
                Case "top":    lg.Position = -4160
                Case "bottom": lg.Position = -4107
                Case "corner": lg.Position = 2
            End Select
        End If
        If props.Exists("font_size") Then lg.Font.Size = modActions.ToLong(props("font_size"))
        If props.Exists("font_color") Then
            lg.Font.Color.RGB = modActions.HexToRgb(CStr(props("font_color")))
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub Do_set_chart_type(slideNum As Long, shapeId As Long, value As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11001, "Do_set_chart_type", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11002, "Do_set_chart_type", "not_a_native_chart"
    Dim ch As Object: Set ch = sh.Chart
    ch.ChartType = ChartTypeFromName(value)
End Sub

Public Function ChartTypeFromName(chartName As String) As Long
    Select Case LCase(chartName)
        Case "columnclustered", "xlcolumnclustered": ChartTypeFromName = 51
        Case "columnstacked":                         ChartTypeFromName = 52
        Case "line", "xlline":                        ChartTypeFromName = 4
        Case "linemarkers", "line_markers":           ChartTypeFromName = 65
        Case "linestacked", "line_stacked":           ChartTypeFromName = 63
        Case "linestackedmarkers", "line_stacked_markers": ChartTypeFromName = 66
        Case "areastacked", "area_stacked":           ChartTypeFromName = 76
        Case "areapercent", "area_100pct":            ChartTypeFromName = 78
        Case "barstackedpercent", "bar_100pct":       ChartTypeFromName = 59
        Case "columnstackedpercent", "column_100pct": ChartTypeFromName = 53
        Case "waterfall":                             ChartTypeFromName = 119  ' xlWaterfall (Office 2016+)
        Case "pareto":                                ChartTypeFromName = 122  ' xlPareto
        Case "funnel":                                ChartTypeFromName = 123  ' xlFunnel (Office 2019+)
        Case "histogram":                             ChartTypeFromName = 118
        Case "boxwhisker", "box_whisker":             ChartTypeFromName = 121
        Case "treemap":                               ChartTypeFromName = 117
        Case "sunburst":                              ChartTypeFromName = 120
        Case "radar":                                 ChartTypeFromName = -4151
        Case "radarmarkers", "radar_markers":         ChartTypeFromName = 81
        Case "radarfilled", "radar_filled":           ChartTypeFromName = 82
        Case "pie", "xlpie":                          ChartTypeFromName = 5
        Case "barclustered", "xlbarclustered":        ChartTypeFromName = 57
        Case "barstacked":                            ChartTypeFromName = 58
        Case "area", "xlarea":                        ChartTypeFromName = 1
        Case "scatter", "xlxyscatter":                ChartTypeFromName = -4169
        Case "doughnut":                              ChartTypeFromName = -4120
        Case Else:
            Err.Raise vbObjectError + 11003, "ChartTypeFromName", "unknown chart type: " & chartName
    End Select
End Function

Public Sub Do_set_chart_title(slideNum As Long, shapeId As Long, _
                              value As String, enabled As Boolean, _
                              Optional ByVal props As Object = Nothing)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11001, "Do_set_chart_title", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11002, "Do_set_chart_title", "not_a_native_chart"
    Dim ch As Object: Set ch = sh.Chart
    If Not enabled Then
        ch.HasTitle = False
        Exit Sub
    End If
    ch.HasTitle = True
    ch.ChartTitle.Text = value
    If props Is Nothing Then Exit Sub
    On Error Resume Next
    Dim t As Object: Set t = ch.ChartTitle
    If props.Exists("font_size") Then
        t.Font.Size = modActions.ToLong(props("font_size"))
        t.Format.TextFrame2.TextRange.Font.Size = modActions.ToLong(props("font_size"))
    End If
    If props.Exists("font_color") Then
        Dim tc As Long: tc = modActions.HexToRgb(CStr(props("font_color")))
        t.Font.Color.RGB = tc
        t.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = tc
    End If
    If props.Exists("font_bold") Then
        t.Font.Bold = modActions.ToBool(props("font_bold"))
        t.Format.TextFrame2.TextRange.Font.Bold = modActions.ToBool(props("font_bold"))
    End If
    If props.Exists("font_italic") Then
        t.Font.Italic = modActions.ToBool(props("font_italic"))
        t.Format.TextFrame2.TextRange.Font.Italic = modActions.ToBool(props("font_italic"))
    End If
    If props.Exists("position") Then
        Select Case LCase(CStr(props("position")))
            Case "above":   t.Position = -4160     ' xlChartTitlePositionAbove
            Case "overlay": t.Position = 0          ' overlaid on plot
            Case "left":    t.Position = -4131
            Case "right":   t.Position = -4152
        End Select
    End If
    On Error GoTo 0
End Sub

Public Sub Do_set_chart_axis_title(slideNum As Long, shapeId As Long, _
                                   axis As String, value As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11001, "Do_set_chart_axis_title", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11002, "Do_set_chart_axis_title", "not_a_native_chart"
    Dim ch As Object: Set ch = sh.Chart
    Dim axNum As Long
    Select Case LCase(axis)
        Case "x": axNum = 1
        Case "y": axNum = 2
        Case Else: Err.Raise vbObjectError + 11004, "Do_set_chart_axis_title", "axis must be 'x' or 'y'"
    End Select
    ch.Axes(axNum).HasTitle = True
    ch.Axes(axNum).AxisTitle.Text = value
End Sub

Public Sub Do_set_chart_legend_position(slideNum As Long, shapeId As Long, value As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11001, "Do_set_chart_legend_position", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11002, "Do_set_chart_legend_position", "not_a_native_chart"
    Dim ch As Object: Set ch = sh.Chart
    If LCase(value) = "none" Then
        ch.HasLegend = False
        Exit Sub
    End If
    ch.HasLegend = True
    Select Case LCase(value)
        Case "left":   ch.Legend.Position = -4131
        Case "right":  ch.Legend.Position = -4152
        Case "top":    ch.Legend.Position = -4160
        Case "bottom": ch.Legend.Position = -4107
        Case "corner": ch.Legend.Position = 2
        Case Else: Err.Raise vbObjectError + 11005, "Do_set_chart_legend_position", "unknown position: " & value
    End Select
End Sub

Public Sub Do_set_series_color(slideNum As Long, shapeId As Long, _
                               seriesIndex As Long, hexValue As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11001, "Do_set_series_color", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11002, "Do_set_series_color", "not_a_native_chart"
    Dim ch As Object: Set ch = sh.Chart
    If seriesIndex < 1 Or seriesIndex > ch.SeriesCollection.Count Then
        Err.Raise vbObjectError + 11006, "Do_set_series_color", "series_index out of range"
    End If
    Dim ser As Object: Set ser = ch.SeriesCollection(seriesIndex)
    ser.Format.Fill.ForeColor.RGB = modActions.HexToRgb(hexValue)
End Sub

Public Sub Do_set_series_values(slideNum As Long, shapeId As Long, _
                                 seriesIndex As Long, values As Variant)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11001, "Do_set_series_values", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11002, "Do_set_series_values", "not_a_native_chart"
    Dim ch As Object: Set ch = sh.Chart
    If seriesIndex < 1 Or seriesIndex > ch.SeriesCollection.Count Then
        Err.Raise vbObjectError + 11006, "Do_set_series_values", "series_index out of range"
    End If
    Dim arr() As Double
    Dim n As Long: n = NormalizeDoubleArray(values, arr)
    If n < 1 Then Err.Raise vbObjectError + 11007, "Do_set_series_values", "values: empty"
    ch.SeriesCollection(seriesIndex).Values = arr
End Sub

Public Sub Do_set_chart_categories(slideNum As Long, shapeId As Long, categories As Variant)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11001, "Do_set_chart_categories", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11002, "Do_set_chart_categories", "not_a_native_chart"
    Dim ch As Object: Set ch = sh.Chart
    Dim arr() As String
    Dim n As Long: n = NormalizeStringArray(categories, arr)
    If n < 1 Then Err.Raise vbObjectError + 11008, "Do_set_chart_categories", "categories: empty"
    If ch.SeriesCollection.Count < 1 Then Err.Raise vbObjectError + 11009, "Do_set_chart_categories", "no series"
    ch.SeriesCollection(1).XValues = arr
End Sub

Public Sub Do_set_series_name(slideNum As Long, shapeId As Long, _
                               seriesIndex As Long, name As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11001, "Do_set_series_name", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11002, "Do_set_series_name", "not_a_native_chart"
    Dim ch As Object: Set ch = sh.Chart
    If seriesIndex < 1 Or seriesIndex > ch.SeriesCollection.Count Then
        Err.Raise vbObjectError + 11006, "Do_set_series_name", "series_index out of range"
    End If
    ch.SeriesCollection(seriesIndex).Name = name
End Sub

Private Function NormalizeDoubleArray(v As Variant, ByRef out() As Double) As Long
    Dim col As Object
    If TypeName(v) = "Collection" Then
        Set col = v
        If col.Count = 0 Then NormalizeDoubleArray = 0: Exit Function
        ReDim out(0 To col.Count - 1)
        Dim i As Long
        For i = 1 To col.Count
            out(i - 1) = CDbl(col(i))
        Next i
        NormalizeDoubleArray = col.Count
    ElseIf IsArray(v) Then
        Dim lo As Long, hi As Long
        On Error Resume Next
        lo = LBound(v): hi = UBound(v)
        On Error GoTo 0
        If hi < lo Then NormalizeDoubleArray = 0: Exit Function
        ReDim out(0 To hi - lo)
        For i = lo To hi
            out(i - lo) = CDbl(v(i))
        Next i
        NormalizeDoubleArray = hi - lo + 1
    Else
        ReDim out(0 To 0)
        out(0) = CDbl(v)
        NormalizeDoubleArray = 1
    End If
End Function

Private Function NormalizeStringArray(v As Variant, ByRef out() As String) As Long
    Dim col As Object
    If TypeName(v) = "Collection" Then
        Set col = v
        If col.Count = 0 Then NormalizeStringArray = 0: Exit Function
        ReDim out(0 To col.Count - 1)
        Dim i As Long
        For i = 1 To col.Count
            out(i - 1) = CStr(col(i))
        Next i
        NormalizeStringArray = col.Count
    ElseIf IsArray(v) Then
        Dim lo As Long, hi As Long
        On Error Resume Next
        lo = LBound(v): hi = UBound(v)
        On Error GoTo 0
        If hi < lo Then NormalizeStringArray = 0: Exit Function
        ReDim out(0 To hi - lo)
        For i = lo To hi
            out(i - lo) = CStr(v(i))
        Next i
        NormalizeStringArray = hi - lo + 1
    Else
        ReDim out(0 To 0)
        out(0) = CStr(v)
        NormalizeStringArray = 1
    End If
End Function
