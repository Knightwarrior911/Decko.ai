Attribute VB_Name = "modActionsConnector"
Option Explicit

Public Sub Do_add_connector(slideNum As Long, fromId As Long, toId As Long, _
                            kind As String, arrowEnd As String, _
                            hexColor As String, weightPt As Single)
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
    conn.ConnectorFormat.BeginConnect shFrom, 1
    conn.ConnectorFormat.EndConnect shTo, 1
    conn.RerouteConnections

    conn.Line.ForeColor.RGB = modActions.HexToRgb(hexColor)
    conn.Line.Weight = weightPt

    Select Case LCase(arrowEnd)
        Case "filled":  conn.Line.EndArrowheadStyle = 5
        Case "open":    conn.Line.EndArrowheadStyle = 2
        Case "none":    conn.Line.EndArrowheadStyle = 1
        Case Else:      conn.Line.EndArrowheadStyle = 5
    End Select
End Sub
