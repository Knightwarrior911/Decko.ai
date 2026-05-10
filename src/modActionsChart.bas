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

    ' Adjust SeriesCollection size to match requested series count.
    ' AddChart2 creates a default chart with N existing series (varies by type).
    Dim existingCount As Long: existingCount = ch.SeriesCollection.Count
    Dim s As Long
    ' Delete extras
    Do While ch.SeriesCollection.Count > seriesCount
        ch.SeriesCollection(ch.SeriesCollection.Count).Delete
    Loop
    ' Add missing
    Do While ch.SeriesCollection.Count < seriesCount
        ch.SeriesCollection.NewSeries
    Loop

    ' Populate each series
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

    ' Clean style: hide y-axis, hide gridlines, hide chart/plot borders & fills
    If cleanStyle Then
        On Error Resume Next
        ch.HasAxis(2, 1) = False                ' xlValue=2, xlPrimary=1
        ch.Axes(1).MajorGridlines.Delete         ' xlCategory=1
        ch.Axes(2).MajorGridlines.Delete         ' xlValue=2
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
    Select Case LCase(Trim(axis))
        Case "x", "category", "xlcategory":  axNum = 1   ' xlCategory
        Case "y", "value", "xlvalue":        axNum = 2   ' xlValue
        Case Else: Err.Raise vbObjectError + 11022, "Do_set_chart_axis", _
                              "axis must be x/y/category/value, got: " & axis
    End Select

    On Error Resume Next
    If props.Exists("visible") Then
        ch.HasAxis(axNum, 1) = modActions.ToBool(props("visible"))
    End If
    Dim ax As Object: Set ax = ch.Axes(axNum)
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
    If props.Exists("line_color") Then
        ser.Format.Line.ForeColor.RGB = modActions.HexToRgb(CStr(props("line_color")))
    End If
    If props.Exists("line_weight") Then
        ser.Format.Line.Weight = CDbl(props("line_weight"))
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
        If props.Exists("label_size") Then lbls.Font.Size = modActions.ToLong(props("label_size"))
        If props.Exists("label_bold") Then lbls.Font.Bold = modActions.ToBool(props("label_bold"))
        If props.Exists("label_italic") Then lbls.Font.Italic = modActions.ToBool(props("label_italic"))
        If props.Exists("label_color") Then
            lbls.Font.Color.RGB = modActions.HexToRgb(CStr(props("label_color")))
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
    ' Per-point custom label text override
    If props.Exists("custom_labels") Then
        Dim customs As Object: Set customs = props("custom_labels")
        Dim p As Long
        For p = 1 To ser.Points.Count
            If p <= customs.Count Then
                ser.Points(p).HasDataLabel = True
                ser.Points(p).DataLabel.Text = CStr(customs(p))
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
                              value As String, enabled As Boolean)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11001, "Do_set_chart_title", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11002, "Do_set_chart_title", "not_a_native_chart"
    Dim ch As Object: Set ch = sh.Chart
    If enabled Then
        ch.HasTitle = True
        ch.ChartTitle.Text = value
    Else
        ch.HasTitle = False
    End If
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
