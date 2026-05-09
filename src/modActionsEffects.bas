Attribute VB_Name = "modActionsEffects"
Option Explicit

' Visual polish actions: rotation/flip/lines, shadow/glow/reflection,
' gradient/3D/preset, picture crop/recolor/brightness/contrast.

' ---- Shape geometry ----

Public Sub Do_rotate_shape(slideNum As Long, shapeId As Long, degrees As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_rotate_shape", "shape not found"
    sh.Rotation = degrees
End Sub

Public Sub Do_flip_shape(slideNum As Long, shapeId As Long, axis As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8002, "Do_flip_shape", "shape not found"
    Select Case LCase(axis)
        Case "h": sh.Flip msoFlipHorizontal
        Case "v": sh.Flip msoFlipVertical
        Case Else: Err.Raise vbObjectError + 8002, "Do_flip_shape", "axis must be h or v"
    End Select
End Sub

Public Sub Do_set_line_color(slideNum As Long, shapeId As Long, hexValue As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8003, "Do_set_line_color", "shape not found"
    sh.Line.Visible = msoTrue
    sh.Line.ForeColor.RGB = modActions.HexToRgb(hexValue)
End Sub

Public Sub Do_set_line_weight(slideNum As Long, shapeId As Long, weightPt As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8004, "Do_set_line_weight", "shape not found"
    If weightPt <= 0 Then Err.Raise vbObjectError + 8004, "Do_set_line_weight", "weight_pt must be > 0"
    sh.Line.Visible = msoTrue
    sh.Line.Weight = weightPt
End Sub

Public Sub Do_set_line_style(slideNum As Long, shapeId As Long, style As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8005, "Do_set_line_style", "shape not found"
    sh.Line.Visible = msoTrue
    Select Case LCase(style)
        Case "solid":   sh.Line.DashStyle = msoLineSolid
        Case "dash":    sh.Line.DashStyle = msoLineDash
        Case "dot":     sh.Line.DashStyle = msoLineRoundDot
        Case "dashdot": sh.Line.DashStyle = msoLineDashDot
        Case Else: Err.Raise vbObjectError + 8005, "Do_set_line_style", "style invalid: " & style
    End Select
End Sub

' ---- Shape effects ----

Public Sub Do_set_shadow(slideNum As Long, shapeId As Long, _
                          offsetX As Double, offsetY As Double, _
                          blur As Double, colorHex As String, transparency As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8006, "Do_set_shadow", "shape not found"
    sh.Shadow.Visible = msoTrue
    sh.Shadow.Type = msoShadow21   ' generic outer shadow
    sh.Shadow.OffsetX = offsetX
    sh.Shadow.OffsetY = offsetY
    sh.Shadow.Blur = blur
    sh.Shadow.ForeColor.RGB = modActions.HexToRgb(colorHex)
    If transparency >= 0 And transparency <= 1 Then sh.Shadow.Transparency = transparency
End Sub

Public Sub Do_set_glow(slideNum As Long, shapeId As Long, _
                        colorHex As String, radius As Double, transparency As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8007, "Do_set_glow", "shape not found"
    sh.Glow.Color.RGB = modActions.HexToRgb(colorHex)
    sh.Glow.Radius = radius
    If transparency >= 0 And transparency <= 1 Then sh.Glow.Transparency = transparency
End Sub

Public Sub Do_set_reflection(slideNum As Long, shapeId As Long, _
                              sizeFrac As Double, transparency As Double, distance As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8008, "Do_set_reflection", "shape not found"
    sh.Reflection.Type = msoReflectionType5   ' tight reflection
    sh.Reflection.Size = sizeFrac
    If transparency >= 0 And transparency <= 1 Then sh.Reflection.Transparency = transparency
    sh.Reflection.Offset = distance
    sh.Reflection.Blur = 4
End Sub

Public Sub Do_set_transparency(slideNum As Long, shapeId As Long, value As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8009, "Do_set_transparency", "shape not found"
    If value < 0 Or value > 1 Then Err.Raise vbObjectError + 8009, "Do_set_transparency", "value must be 0..1"
    On Error Resume Next
    sh.Fill.Transparency = value
    On Error GoTo 0
End Sub

' ---- Gradient + 3D ----

Public Sub Do_set_gradient_fill(slideNum As Long, shapeId As Long, _
                                 color1Hex As String, color2Hex As String, angleDeg As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8010, "Do_set_gradient_fill", "shape not found"
    sh.Fill.TwoColorGradient msoGradientHorizontal, 1
    sh.Fill.ForeColor.RGB = modActions.HexToRgb(color1Hex)
    sh.Fill.BackColor.RGB = modActions.HexToRgb(color2Hex)
    sh.Fill.GradientAngle = angleDeg
End Sub

Public Sub Do_set_3d_bevel(slideNum As Long, shapeId As Long, bevelType As String, depthPt As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8011, "Do_set_3d_bevel", "shape not found"
    Dim bt As Long
    Select Case LCase(bevelType)
        Case "circle":   bt = msoBevelCircle
        Case "slope":    bt = msoBevelSlope
        Case "cross":    bt = msoBevelCross
        Case "angle":    bt = msoBevelAngle
        Case "softround": bt = msoBevelSoftRound
        Case Else: Err.Raise vbObjectError + 8011, "Do_set_3d_bevel", "type invalid: " & bevelType
    End Select
    sh.ThreeD.BevelTopType = bt
    sh.ThreeD.BevelTopDepth = depthPt
End Sub

Public Sub Do_apply_preset_effect(slideNum As Long, shapeId As Long, presetIndex As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8012, "Do_apply_preset_effect", "shape not found"
    If presetIndex < 1 Or presetIndex > 24 Then _
        Err.Raise vbObjectError + 8012, "Do_apply_preset_effect", "preset 1..24"
    sh.Fill.PresetTextured presetIndex   ' MsoPresetTexture enum 1..24
End Sub

' ---- Picture-specific ----

Public Sub Do_crop_picture(slideNum As Long, shapeId As Long, _
                            leftPt As Double, rightPt As Double, _
                            topPt As Double, bottomPt As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8013, "Do_crop_picture", "shape not found"
    If sh.Type <> msoPicture And sh.Type <> msoLinkedPicture Then _
        Err.Raise vbObjectError + 8013, "Do_crop_picture", "shape is not a picture"
    sh.PictureFormat.CropLeft = leftPt
    sh.PictureFormat.CropRight = rightPt
    sh.PictureFormat.CropTop = topPt
    sh.PictureFormat.CropBottom = bottomPt
End Sub

Public Sub Do_recolor_picture(slideNum As Long, shapeId As Long, colorType As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8014, "Do_recolor_picture", "shape not found"
    If sh.Type <> msoPicture And sh.Type <> msoLinkedPicture Then _
        Err.Raise vbObjectError + 8014, "Do_recolor_picture", "shape is not a picture"
    Select Case LCase(colorType)
        Case "grayscale": sh.PictureFormat.ColorType = msoPictureGrayscale
        Case "sepia":     sh.PictureFormat.ColorType = msoPictureWatermark   ' approximate
        Case "washout":   sh.PictureFormat.ColorType = msoPictureWatermark
        Case "bw":        sh.PictureFormat.ColorType = msoPictureBlackAndWhite
        Case "auto":      sh.PictureFormat.ColorType = msoPictureAutomatic
        Case Else: Err.Raise vbObjectError + 8014, "Do_recolor_picture", "colorType invalid: " & colorType
    End Select
End Sub

Public Sub Do_set_brightness(slideNum As Long, shapeId As Long, value As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8015, "Do_set_brightness", "shape not found"
    If sh.Type <> msoPicture And sh.Type <> msoLinkedPicture Then _
        Err.Raise vbObjectError + 8015, "Do_set_brightness", "shape is not a picture"
    If value < -1 Or value > 1 Then Err.Raise vbObjectError + 8015, "Do_set_brightness", "value must be -1..1"
    sh.PictureFormat.Brightness = value
End Sub

Public Sub Do_set_contrast(slideNum As Long, shapeId As Long, value As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8016, "Do_set_contrast", "shape not found"
    If sh.Type <> msoPicture And sh.Type <> msoLinkedPicture Then _
        Err.Raise vbObjectError + 8016, "Do_set_contrast", "shape is not a picture"
    If value < -1 Or value > 1 Then Err.Raise vbObjectError + 8016, "Do_set_contrast", "value must be -1..1"
    sh.PictureFormat.Contrast = value
End Sub
