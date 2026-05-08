Attribute VB_Name = "modActionsLayout"
Option Explicit

Public Sub Do_align_shapes(slideNum As Long, shapeIds As Variant, anchor As String)
    Dim shapes() As Shape
    Dim n As Long: n = ShapesByIds(slideNum, shapeIds, shapes)
    If n < 2 Then Err.Raise vbObjectError + 4001, "Do_align_shapes", "need >=2 shapes"
    Dim ref As Shape: Set ref = shapes(0)
    Dim i As Long
    For i = 1 To n - 1
        Select Case LCase(anchor)
            Case "left":    shapes(i).Left = ref.Left
            Case "right":   shapes(i).Left = ref.Left + ref.Width - shapes(i).Width
            Case "top":     shapes(i).Top = ref.Top
            Case "bottom":  shapes(i).Top = ref.Top + ref.Height - shapes(i).Height
            Case "hcenter": shapes(i).Left = ref.Left + (ref.Width - shapes(i).Width) / 2
            Case "vcenter": shapes(i).Top = ref.Top + (ref.Height - shapes(i).Height) / 2
            Case Else: Err.Raise vbObjectError + 4002, "Do_align_shapes", "unknown anchor: " & anchor
        End Select
    Next i
End Sub

Public Function ShapesByIds(slideNum As Long, shapeIds As Variant, ByRef out() As Shape) As Long
    Dim ids() As Long
    Dim cnt As Long: cnt = NormalizeIdsArray(shapeIds, ids)
    ReDim out(0 To cnt - 1)
    Dim i As Long, found As Long: found = 0
    For i = 0 To cnt - 1
        Dim sh As Shape: Set sh = modActions.FindShape(slideNum, ids(i))
        If Not sh Is Nothing Then
            Set out(found) = sh
            found = found + 1
        End If
    Next i
    If found < cnt Then ReDim Preserve out(0 To found - 1)
    ShapesByIds = found
End Function

Public Function NormalizeIdsArray(v As Variant, ByRef out() As Long) As Long
    Dim col As Object
    If TypeName(v) = "Collection" Then
        Set col = v
        ReDim out(0 To col.Count - 1)
        Dim i As Long
        For i = 1 To col.Count
            out(i - 1) = CLng(col(i))
        Next i
        NormalizeIdsArray = col.Count
    ElseIf IsArray(v) Then
        Dim lo As Long, hi As Long
        lo = LBound(v): hi = UBound(v)
        ReDim out(0 To hi - lo)
        For i = lo To hi
            out(i - lo) = CLng(v(i))
        Next i
        NormalizeIdsArray = hi - lo + 1
    Else
        ReDim out(0 To 0)
        out(0) = CLng(v)
        NormalizeIdsArray = 1
    End If
End Function

Public Sub Do_distribute_horizontal(slideNum As Long, shapeIds As Variant)
    Dim shapes() As Shape
    Dim n As Long: n = ShapesByIds(slideNum, shapeIds, shapes)
    If n < 3 Then Err.Raise vbObjectError + 4003, "Do_distribute_horizontal", "need >=3 shapes"
    SortShapesByLeft shapes
    Dim minLeft As Single: minLeft = shapes(0).Left
    Dim maxLeft As Single: maxLeft = shapes(n - 1).Left
    Dim gapStep As Single: gapStep = (maxLeft - minLeft) / (n - 1)
    Dim i As Long
    For i = 1 To n - 2
        shapes(i).Left = minLeft + gapStep * i
    Next i
End Sub

Public Sub Do_distribute_vertical(slideNum As Long, shapeIds As Variant)
    Dim shapes() As Shape
    Dim n As Long: n = ShapesByIds(slideNum, shapeIds, shapes)
    If n < 3 Then Err.Raise vbObjectError + 4003, "Do_distribute_vertical", "need >=3 shapes"
    SortShapesByTop shapes
    Dim minTop As Single: minTop = shapes(0).Top
    Dim maxTop As Single: maxTop = shapes(n - 1).Top
    Dim gapStep As Single: gapStep = (maxTop - minTop) / (n - 1)
    Dim i As Long
    For i = 1 To n - 2
        shapes(i).Top = minTop + gapStep * i
    Next i
End Sub

Private Sub SortShapesByLeft(ByRef arr() As Shape)
    Dim i As Long, j As Long, tmp As Shape
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j).Left < arr(i).Left Then
                Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Sub SortShapesByTop(ByRef arr() As Shape)
    Dim i As Long, j As Long, tmp As Shape
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j).Top < arr(i).Top Then
                Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next j
    Next i
End Sub
