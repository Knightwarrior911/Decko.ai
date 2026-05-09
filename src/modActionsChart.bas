Attribute VB_Name = "modActionsChart"
Option Explicit

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
