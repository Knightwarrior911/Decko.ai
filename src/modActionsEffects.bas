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

' Set a single Adjustment value on a shape. Used to control parameters of
' parametric shapes — pie wedge start/end angles, callout pointer position,
' rounded-rect corner radius, arrow stem width, etc. Index is 1-based.
' Value units depend on the shape (degrees for pie wedges, fraction 0..1 for
' most others). Caller must know the schema for the shape being adjusted.
Public Sub Do_set_shape_adjustment(slideNum As Long, shapeId As Long, _
                                    idx As Long, value As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8003, "Do_set_shape_adjustment", "shape not found"
    On Error Resume Next
    Dim cnt As Long: cnt = sh.Adjustments.Count
    If Err.Number <> 0 Then
        Err.Clear
        Err.Raise vbObjectError + 8004, "Do_set_shape_adjustment", _
                  "shape has no adjustments collection (kind=" & sh.AutoShapeType & ")"
    End If
    On Error GoTo 0
    If idx < 1 Or idx > cnt Then
        Err.Raise vbObjectError + 8005, "Do_set_shape_adjustment", _
                  "adjustment index " & idx & " out of range 1.." & cnt
    End If
    sh.Adjustments(idx) = value
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

' ---- Effect clearers ----
' Each set_shadow/set_glow/set_reflection action ENABLES that effect; there
' was no symmetric disable. These clearers remove the effect cleanly.

Public Sub Do_clear_shadow(slideNum As Long, shapeId As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8020, "Do_clear_shadow", "shape not found"
    On Error Resume Next
    sh.Shadow.Visible = msoFalse
    On Error GoTo 0
End Sub

Public Sub Do_clear_glow(slideNum As Long, shapeId As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8021, "Do_clear_glow", "shape not found"
    On Error Resume Next
    sh.Glow.Radius = 0           ' radius 0 removes the glow visually
    sh.Glow.Transparency = 1     ' belt-and-braces: full transparency
    On Error GoTo 0
End Sub

Public Sub Do_clear_reflection(slideNum As Long, shapeId As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8022, "Do_clear_reflection", "shape not found"
    On Error Resume Next
    sh.Reflection.Type = 0       ' msoReflectionTypeMixed=-2; 0 = none
    sh.Reflection.Size = 0
    On Error GoTo 0
End Sub

' Strip every visual effect on a shape: shadow, glow, reflection, 3D, soft edges.
' Useful as a "go back to flat" reset before re-styling.
Public Sub Do_clear_all_effects(slideNum As Long, shapeId As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8023, "Do_clear_all_effects", "shape not found"
    On Error Resume Next
    sh.Shadow.Visible = msoFalse
    sh.Glow.Radius = 0
    sh.Glow.Transparency = 1
    sh.Reflection.Type = 0
    sh.Reflection.Size = 0
    sh.ThreeD.BevelTopType = 0      ' msoBevelNone
    sh.ThreeD.BevelTopDepth = 0
    sh.ThreeD.Depth = 0
    sh.SoftEdge.Type = 0            ' msoSoftEdgeNone
    On Error GoTo 0
End Sub

' Soft edge (feathered border). Pass radius 0 to clear.
Public Sub Do_set_soft_edge(slideNum As Long, shapeId As Long, radiusPt As Double)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8024, "Do_set_soft_edge", "shape not found"
    If radiusPt < 0 Then radiusPt = 0
    On Error Resume Next
    If radiusPt = 0 Then
        sh.SoftEdge.Type = 0
    Else
        sh.SoftEdge.Radius = radiusPt
    End If
    On Error GoTo 0
End Sub

' 3D rotation around X/Y/Z axes (degrees). Useful for pseudo-isometric layouts.
' Any axis you omit stays at its current value.
Public Sub Do_set_3d_rotation(slideNum As Long, shapeId As Long, _
                               rotX As Double, rotY As Double, rotZ As Double, _
                               hasX As Boolean, hasY As Boolean, hasZ As Boolean)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8025, "Do_set_3d_rotation", "shape not found"
    On Error Resume Next
    If hasX Then sh.ThreeD.RotationX = rotX
    If hasY Then sh.ThreeD.RotationY = rotY
    If hasZ Then sh.ThreeD.RotationZ = rotZ
    On Error GoTo 0
End Sub

' Apply a Photoshop-style artistic effect to a picture. effect:
'   "none", "marker", "pencil_grayscale", "pencil_sketch", "line_drawing",
'   "chalk_sketch", "paint_strokes", "paint_brush", "glow_diffused",
'   "blur", "light_screen", "watercolor", "film_grain", "mosaic_bubbles",
'   "glass", "cement", "texturizer", "crisscross", "pastels_smooth", "plastic_wrap",
'   "cutout", "photocopy", "glow_edges"
' intensity: 0..100 (optional, default 50)
Public Sub Do_apply_picture_artistic_effect(slideNum As Long, shapeId As Long, _
                                             effect As String, intensity As Long, _
                                             hasIntensity As Boolean)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8030, "Do_apply_picture_artistic_effect", "shape not found"
    If sh.Type <> msoPicture And sh.Type <> msoLinkedPicture Then _
        Err.Raise vbObjectError + 8030, "Do_apply_picture_artistic_effect", "shape is not a picture"
    Dim eff As Long
    Select Case LCase(Trim(effect))
        Case "none":              eff = 0
        Case "marker":            eff = 1
        Case "pencil_grayscale":  eff = 2
        Case "pencil_sketch":     eff = 3
        Case "line_drawing":      eff = 4
        Case "chalk_sketch":      eff = 5
        Case "paint_strokes":     eff = 6
        Case "paint_brush":       eff = 7
        Case "glow_diffused":     eff = 8
        Case "blur":              eff = 9
        Case "light_screen":      eff = 10
        Case "watercolor":        eff = 11
        Case "film_grain":        eff = 12
        Case "mosaic_bubbles":    eff = 13
        Case "glass":             eff = 14
        Case "cement":            eff = 15
        Case "texturizer":        eff = 16
        Case "crisscross":        eff = 17
        Case "pastels_smooth":    eff = 18
        Case "plastic_wrap":      eff = 19
        Case "cutout":            eff = 20
        Case "photocopy":         eff = 21
        Case "glow_edges":        eff = 22
        Case Else: Err.Raise vbObjectError + 8030, "Do_apply_picture_artistic_effect", "unknown effect: " & effect
    End Select
    On Error Resume Next
    sh.PictureFormat.ArtisticEffect = eff
    If hasIntensity Then
        If intensity < 0 Then intensity = 0
        If intensity > 100 Then intensity = 100
        sh.PictureFormat.ArtisticEffectIntensity = intensity
    End If
    On Error GoTo 0
End Sub

' Reset a picture to its original state — undo all corrections (brightness,
' contrast, crop, artistic effects, recolor). Useful after over-editing.
Public Sub Do_reset_picture(slideNum As Long, shapeId As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8031, "Do_reset_picture", "shape not found"
    If sh.Type <> msoPicture And sh.Type <> msoLinkedPicture Then _
        Err.Raise vbObjectError + 8031, "Do_reset_picture", "shape is not a picture"
    On Error Resume Next
    With sh.PictureFormat
        .Brightness = 0
        .Contrast = 0
        .ColorType = msoPictureAutomatic
        .CropLeft = 0: .CropRight = 0: .CropTop = 0: .CropBottom = 0
        .ArtisticEffect = 0
    End With
    On Error GoTo 0
End Sub

' Toggle shape visibility (hide without deleting). Useful for staging /
' conditional layouts.
Public Sub Do_set_shape_visible(slideNum As Long, shapeId As Long, value As Boolean)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8032, "Do_set_shape_visible", "shape not found"
    sh.Visible = IIf(value, msoTrue, msoFalse)
End Sub

' Reconnect an existing connector to different endpoint shapes. Faster than
' delete + add_connector when only endpoints change.
Public Sub Do_reconnect_connector(slideNum As Long, connectorShapeId As Long, _
                                   fromShapeId As Long, toShapeId As Long, _
                                   fromConnectionSite As Long, toConnectionSite As Long, _
                                   hasFromSite As Boolean, hasToSite As Boolean)
    Dim conn As Shape: Set conn = modActions.FindShape(slideNum, connectorShapeId)
    If conn Is Nothing Then Err.Raise vbObjectError + 8033, "Do_reconnect_connector", "connector not found"
    If Not conn.Connector Then Err.Raise vbObjectError + 8033, "Do_reconnect_connector", "shape is not a connector"
    Dim fromSh As Shape: Set fromSh = modActions.FindShape(slideNum, fromShapeId)
    Dim toSh As Shape: Set toSh = modActions.FindShape(slideNum, toShapeId)
    If fromSh Is Nothing Or toSh Is Nothing Then _
        Err.Raise vbObjectError + 8033, "Do_reconnect_connector", "from/to shape not found"
    Dim fSite As Long: fSite = 1
    Dim tSite As Long: tSite = 1
    If hasFromSite Then fSite = fromConnectionSite
    If hasToSite Then tSite = toConnectionSite
    On Error Resume Next
    conn.ConnectorFormat.BeginConnect fromSh, fSite
    conn.ConnectorFormat.EndConnect toSh, tSite
    conn.RerouteConnections
    On Error GoTo 0
End Sub
