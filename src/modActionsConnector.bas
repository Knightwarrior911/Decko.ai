Attribute VB_Name = "modActionsConnector"
Option Explicit

Public Sub Do_add_connector(slideNum As Long, fromId As Long, toId As Long, _
                            kind As String, _
                            Optional arrowEnd As String = "filled", _
                            Optional hexColor As String = "#000000", _
                            Optional weightPt As Single = 1.0, _
                            Optional arrowStart As String = "none", _
                            Optional arrowSize As String = "medium", _
                            Optional fromPoint As String = "auto", _
                            Optional toPoint As String = "auto", _
                            Optional dashStyle As String = "solid")
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 10001, "Do_add_connector", "slide_out_of_range"
    End If
    Dim sl As Slide: Set sl = pres.Slides(slideNum)
    Dim shFrom As Shape, shTo As Shape
    Set shFrom = modActions.FindShape(slideNum, fromId)
    Set shTo = modActions.FindShape(slideNum, toId)
    If shFrom Is Nothing Or shTo Is Nothing Then
        Err.Raise vbObjectError + 10002, "Do_add_connector", "endpoint shape not found"
    End If

    Dim ctype As Long
    Select Case LCase(kind)
        Case "straight": ctype = 1
        Case "elbow":    ctype = 2
        Case "curved":   ctype = 3
        Case Else:       ctype = 1
    End Select

    Dim conn As Shape
    Set conn = sl.Shapes.AddConnector(ctype, 0, 0, 100, 100)

    Dim fpIdx As Long: fpIdx = ResolveConnPoint(fromPoint)
    Dim tpIdx As Long: tpIdx = ResolveConnPoint(toPoint)
    Dim autoRoute As Boolean: autoRoute = (fpIdx = 0 And tpIdx = 0)

    If fpIdx > 0 Then
        conn.ConnectorFormat.BeginConnect shFrom, fpIdx
    Else
        conn.ConnectorFormat.BeginConnect shFrom, 1
    End If
    If tpIdx > 0 Then
        conn.ConnectorFormat.EndConnect shTo, tpIdx
    Else
        conn.ConnectorFormat.EndConnect shTo, 1
    End If
    If autoRoute Then conn.RerouteConnections

    conn.Line.ForeColor.RGB = modActions.HexToRgb(hexColor)
    conn.Line.Weight = weightPt

    ' Arrow size: 1=small 2=medium 3=large
    Dim aSize As Long
    Select Case LCase(arrowSize)
        Case "small":  aSize = 1
        Case "large":  aSize = 3
        Case Else:     aSize = 2
    End Select

    conn.Line.EndArrowheadStyle = ResolveArrowStyle(arrowEnd)
    conn.Line.EndArrowheadLength = aSize
    conn.Line.EndArrowheadWidth = aSize
    conn.Line.BeginArrowheadStyle = ResolveArrowStyle(arrowStart)
    conn.Line.BeginArrowheadLength = aSize
    conn.Line.BeginArrowheadWidth = aSize

    Select Case LCase(dashStyle)
        Case "dash":     conn.Line.DashStyle = msoLineDash
        Case "dot":      conn.Line.DashStyle = msoLineSquareDot
        Case "round_dot": conn.Line.DashStyle = msoLineRoundDot
        Case "dash_dot": conn.Line.DashStyle = msoLineDashDot
        Case "long_dash": conn.Line.DashStyle = msoLineLongDash
        Case "long_dash_dot": conn.Line.DashStyle = msoLineLongDashDot
        Case Else:       conn.Line.DashStyle = msoLineSolid
    End Select
End Sub

Private Function ResolveArrowStyle(s As String) As Long
    ' msoArrowheadStyle: None=1, Triangle=2, Open=3, Stealth=4, Diamond=5, Oval=6
    Select Case LCase(Trim(s))
        Case "filled", "triangle": ResolveArrowStyle = 2
        Case "open":               ResolveArrowStyle = 3
        Case "stealth":            ResolveArrowStyle = 4
        Case "diamond":            ResolveArrowStyle = 5
        Case "oval":               ResolveArrowStyle = 6
        Case "none", "":           ResolveArrowStyle = 1
        Case Else:                 ResolveArrowStyle = 1
    End Select
End Function

Private Function ResolveConnPoint(pt As String) As Long
    Select Case LCase(Trim(pt))
        Case "top":    ResolveConnPoint = 1
        Case "right":  ResolveConnPoint = 2
        Case "bottom": ResolveConnPoint = 3
        Case "left":   ResolveConnPoint = 4
        Case "auto", "": ResolveConnPoint = 0
        Case Else
            If IsNumeric(pt) Then
                ResolveConnPoint = CLng(pt)
            Else
                ResolveConnPoint = 0
            End If
    End Select
End Function
