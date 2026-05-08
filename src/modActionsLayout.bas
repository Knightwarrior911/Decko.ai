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
        If col.Count = 0 Then
            ReDim out(0 To 0)
            NormalizeIdsArray = 0
            Exit Function
        End If
        ReDim out(0 To col.Count - 1)
        Dim i As Long
        For i = 1 To col.Count
            out(i - 1) = CLng(col(i))
        Next i
        NormalizeIdsArray = col.Count
    ElseIf IsArray(v) Then
        Dim lo As Long, hi As Long
        On Error Resume Next
        lo = LBound(v)
        hi = UBound(v)
        On Error GoTo 0
        If hi < lo Then
            ReDim out(0 To 0)
            NormalizeIdsArray = 0
            Exit Function
        End If
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

Public Sub Do_tile_grid(slideNum As Long, shapeIds As Variant, cols As Long, gapPt As Single)
    If cols < 1 Then Err.Raise vbObjectError + 4004, "Do_tile_grid", "cols must be >=1"
    Dim shapes() As Shape
    Dim n As Long: n = ShapesByIds(slideNum, shapeIds, shapes)
    If n < 1 Then Err.Raise vbObjectError + 4001, "Do_tile_grid", "no shapes"
    Dim originX As Single: originX = shapes(0).Left
    Dim originY As Single: originY = shapes(0).Top
    Dim cellW As Single: cellW = shapes(0).Width + gapPt
    Dim cellH As Single: cellH = shapes(0).Height + gapPt
    Dim i As Long
    For i = 0 To n - 1
        Dim row As Long: row = i \ cols
        Dim col As Long: col = i Mod cols
        shapes(i).Left = originX + col * cellW
        shapes(i).Top = originY + row * cellH
    Next i
End Sub

Public Sub Do_fit_to_slide_margins(slideNum As Long, shapeId As Long, marginPt As Single)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 4001, "Do_fit_to_slide_margins", "shape not found"
    Dim sw As Single: sw = ActivePresentation.PageSetup.SlideWidth
    Dim shgt As Single: shgt = ActivePresentation.PageSetup.SlideHeight
    sh.Left = marginPt
    sh.Top = marginPt
    sh.LockAspectRatio = msoFalse
    sh.Width = sw - 2 * marginPt
    sh.Height = shgt - 2 * marginPt
End Sub

Public Sub Do_add_line(slideNum As Long, x1 As Single, y1 As Single, _
                       x2 As Single, y2 As Single, _
                       hexColor As String, weightPt As Single)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 4005, "Do_add_line", "slide_out_of_range"
    End If
    Dim ln As Shape
    Set ln = pres.Slides(slideNum).Shapes.AddLine(x1, y1, x2, y2)
    ln.Line.ForeColor.RGB = modActions.HexToRgb(hexColor)
    ln.Line.Weight = weightPt
End Sub

Public Sub Do_add_shape(slideNum As Long, kind As String, _
                        leftPt As Single, topPt As Single, _
                        widthPt As Single, heightPt As Single, _
                        fillHex As String, strokeHex As String, _
                        strokeWeight As Single)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 4005, "Do_add_shape", "slide_out_of_range"
    End If
    Dim msoKind As Long: msoKind = ResolveAutoShapeKind(kind)
    Dim sh As Shape
    Set sh = pres.Slides(slideNum).Shapes.AddShape(msoKind, leftPt, topPt, widthPt, heightPt)
    If Len(fillHex) > 0 Then
        sh.Fill.Visible = msoTrue
        sh.Fill.Solid
        sh.Fill.ForeColor.RGB = modActions.HexToRgb(fillHex)
    Else
        sh.Fill.Visible = msoFalse
    End If
    If Len(strokeHex) > 0 Then
        sh.Line.Visible = msoTrue
        sh.Line.ForeColor.RGB = modActions.HexToRgb(strokeHex)
        sh.Line.Weight = strokeWeight
    Else
        sh.Line.Visible = msoFalse
    End If
End Sub

Public Function ResolveAutoShapeKind(kind As String) As Long
    Select Case LCase(kind)
        Case "rect", "rectangle":   ResolveAutoShapeKind = 1
        Case "rrect", "round_rect": ResolveAutoShapeKind = 5
        Case "oval", "ellipse":     ResolveAutoShapeKind = 9
        Case "circle":              ResolveAutoShapeKind = 9
        Case "capsule":             ResolveAutoShapeKind = 73
        Case "arrow", "right_arrow": ResolveAutoShapeKind = 13
        Case "diamond":             ResolveAutoShapeKind = 4
        Case "triangle":            ResolveAutoShapeKind = 7
        Case Else: Err.Raise vbObjectError + 4006, "ResolveAutoShapeKind", "unknown kind: " & kind
    End Select
End Function

Public Sub Do_set_shape_kind(slideNum As Long, shapeId As Long, kind As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 4001, "Do_set_shape_kind", "shape not found"
    On Error Resume Next
    sh.AutoShapeType = ResolveAutoShapeKind(kind)
    If Err.Number <> 0 Then
        Dim msg As String: msg = "not_an_autoshape: " & Err.Description
        Err.Clear
        Err.Raise vbObjectError + 4007, "Do_set_shape_kind", msg
    End If
    On Error GoTo 0
End Sub

Public Sub Do_clear_slide(slideNum As Long, keepShapeIds As Variant)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 4005, "Do_clear_slide", "slide_out_of_range"
    End If
    Dim sl As Slide: Set sl = pres.Slides(slideNum)

    Dim keepIds() As Long
    Dim keepCount As Long: keepCount = NormalizeIdsArray(keepShapeIds, keepIds)

    Dim toDelete As New Collection
    Dim sh As Shape, i As Long, keep As Boolean
    For Each sh In sl.Shapes
        keep = False
        For i = 0 To keepCount - 1
            If sh.Id = keepIds(i) Then keep = True: Exit For
        Next i
        If Not keep Then toDelete.Add sh.Id
    Next sh

    For i = 1 To toDelete.Count
        Dim victim As Shape: Set victim = modActions.FindShape(slideNum, CLng(toDelete(i)))
        If Not victim Is Nothing Then victim.Delete
    Next i
End Sub

Public Sub Do_move_shape_relative(slideNum As Long, shapeId As Long, dxPt As Single, dyPt As Single)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 4001, "Do_move_shape_relative", "shape not found"
    sh.Left = sh.Left + dxPt
    sh.Top = sh.Top + dyPt
End Sub

Public Sub Do_recolor_fill_match(scope As String, fromHex As String, toHex As String)
    ApplyByScope scope, "fill", fromHex, toHex
End Sub

Public Sub Do_recolor_font_match(scope As String, fromHex As String, toHex As String)
    ApplyByScope scope, "font", fromHex, toHex
End Sub

Public Sub Do_delete_shapes_match(scope As String, kindFilter As String, _
                                  fillFilter As String, textContains As String)
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim slideFilter As Long: slideFilter = ParseScope(scope)
    Dim i As Long
    For i = 1 To pres.Slides.Count
        If slideFilter = 0 Or slideFilter = i Then
            Dim toDelete As New Collection
            Dim sh As Shape
            For Each sh In pres.Slides(i).Shapes
                If MatchesShape(sh, kindFilter, fillFilter, textContains) Then
                    toDelete.Add sh.Id
                End If
            Next sh
            Dim j As Long
            For j = 1 To toDelete.Count
                Dim victim As Shape: Set victim = modActions.FindShape(i, CLng(toDelete(j)))
                If Not victim Is Nothing Then victim.Delete
            Next j
        End If
    Next i
End Sub

Private Function ParseScope(scope As String) As Long
    If LCase(Left(scope, 6)) = "slide:" Then
        ParseScope = CLng(Mid(scope, 7))
    ElseIf LCase(scope) = "deck" Then
        ParseScope = 0
    Else
        Err.Raise vbObjectError + 4008, "ParseScope", "scope must be 'deck' or 'slide:N'"
    End If
End Function

Private Function MatchesShape(sh As Shape, kindFilter As String, _
                              fillFilter As String, textContains As String) As Boolean
    If Len(kindFilter) > 0 Then
        On Error Resume Next
        Dim k As Long: k = ResolveAutoShapeKind(kindFilter)
        If sh.AutoShapeType <> k Then
            MatchesShape = False
            Err.Clear
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If
    If Len(fillFilter) > 0 Then
        On Error Resume Next
        If sh.Fill.Type <> msoFillSolid Then
            MatchesShape = False
            Err.Clear
            Exit Function
        End If
        Dim hexVal As String: hexVal = modExportSnapshot.RgbToHex(sh.Fill.ForeColor.RGB)
        If LCase(hexVal) <> LCase(fillFilter) Then
            MatchesShape = False
            Err.Clear
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If
    If Len(textContains) > 0 Then
        Dim hasText As Boolean: hasText = False
        On Error Resume Next
        If sh.HasTextFrame Then
            If sh.TextFrame.HasText Then
                If InStr(sh.TextFrame.TextRange.Text, textContains) > 0 Then hasText = True
            End If
        End If
        Err.Clear
        On Error GoTo 0
        If Not hasText Then
            MatchesShape = False
            Exit Function
        End If
    End If
    MatchesShape = True
End Function

Private Sub ApplyByScope(scope As String, propertyKind As String, _
                         fromHex As String, toHex As String)
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim slideFilter As Long: slideFilter = ParseScope(scope)
    Dim i As Long
    For i = 1 To pres.Slides.Count
        If slideFilter = 0 Or slideFilter = i Then
            Dim sh As Shape
            For Each sh In pres.Slides(i).Shapes
                ApplyToShape sh, propertyKind, fromHex, toHex
            Next sh
        End If
    Next i
End Sub

Private Sub ApplyToShape(sh As Shape, propertyKind As String, _
                         fromHex As String, toHex As String)
    On Error Resume Next
    If propertyKind = "fill" Then
        If sh.Fill.Type = msoFillSolid Then
            Dim curHex As String: curHex = modExportSnapshot.RgbToHex(sh.Fill.ForeColor.RGB)
            If LCase(curHex) = LCase(fromHex) Then
                sh.Fill.ForeColor.RGB = modActions.HexToRgb(toHex)
            End If
        End If
    ElseIf propertyKind = "font" Then
        If sh.HasTextFrame Then
            If sh.TextFrame.HasText Then
                Dim fc As String: fc = modExportSnapshot.RgbToHex(sh.TextFrame.TextRange.Font.Color.RGB)
                If LCase(fc) = LCase(fromHex) Then
                    sh.TextFrame.TextRange.Font.Color.RGB = modActions.HexToRgb(toHex)
                End If
            End If
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Sub
