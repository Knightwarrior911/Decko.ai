Attribute VB_Name = "modActionsTemplate"
Option Explicit

' =============================================================================
' Template library — apply_template builds a whole on-brand slide skeleton in
' ONE action by composing existing modActionsLayout primitives. The LLM picks
' a template + fills slots instead of emitting ~15 primitive actions (fewer
' primitives -> less intent drift).
'
' Geometry follows the VISUAL DESIGN PRINCIPLES baked into
' frmExport.PromptTemplate: 3-tier hierarchy, clustering, >=36pt margins,
' a single accent colour.
' =============================================================================

Private Const M As Single = 36#               ' slide margin (pt)
Private Const GAP As Single = 18#             ' inter-cluster gap (pt)
Private Const C_DARK As String = "#15283C"    ' dominant / headings
Private Const C_ACCENT As String = "#2E75B6"  ' single accent
Private Const C_BODY As String = "#333333"    ' supporting body
Private Const C_NEUTRAL As String = "#F2F2F2" ' sibling cluster fill
Private Const PPLAYOUT_BLANK As Long = 12

Public Function TemplateNames() As String
    TemplateNames = "title,section,bullets,two_col,comparison,kpi_dashboard,quote"
End Function

' "" if every required slot for the template is present in content, else a
' specific "content.<slot>: required" reason. Called by ValidateAction.
Public Function ValidateTemplateSlots(tpl As String, content As Object) As String
    Dim need As String
    Select Case tpl
        Case "title":         need = "title,subtitle"
        Case "section":       need = "section_number,section_title"
        Case "bullets":       need = "heading,bullets"
        Case "two_col":       need = "heading,left_body,right_body"
        Case "comparison":    need = "heading,left_label,left_body,right_label,right_body"
        Case "kpi_dashboard": need = "heading,tiles"
        Case "quote":         need = "quote_text,attribution"
        Case Else
            ValidateTemplateSlots = "template: must be one of " & TemplateNames()
            Exit Function
    End Select
    Dim parts() As String: parts = Split(need, ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If content Is Nothing Then
            ValidateTemplateSlots = "content." & parts(i) & ": required"
            Exit Function
        ElseIf Not content.Exists(parts(i)) Then
            ValidateTemplateSlots = "content." & parts(i) & ": required"
            Exit Function
        End If
    Next i
End Function

Private Function SW() As Single
    SW = ActivePresentation.PageSetup.SlideWidth
End Function
Private Function SH() As Single
    SH = ActivePresentation.PageSetup.SlideHeight
End Function

' Resolve the target slide: explicit slide=N, else append a fresh blank slide.
Private Function TargetSlide(act As Object) As Long
    Dim pres As Presentation: Set pres = ActivePresentation
    If act.Exists("slide") Then
        TargetSlide = CLng(act("slide"))
        If TargetSlide < 1 Or TargetSlide > pres.Slides.Count Then
            Err.Raise vbObjectError + 9101, "apply_template", "slide_out_of_range"
        End If
    Else
        ' Append a new slide using the deck's own blank layout (deck-agnostic;
        ' the CustomLayouts index varies per template, so resolve by name).
        Dim layouts As Object
        Set layouts = pres.SlideMaster.CustomLayouts
        Dim lay As Object, blankLay As Object
        For Each lay In layouts
            If InStr(LCase(lay.Name), "blank") > 0 Then
                Set blankLay = lay
                Exit For
            End If
        Next lay
        If blankLay Is Nothing Then Set blankLay = layouts(layouts.Count)
        Dim newSlide As Object
        Set newSlide = pres.Slides.AddSlide(pres.Slides.Count + 1, blankLay)
        TargetSlide = newSlide.SlideIndex
    End If
End Function

Private Function CStrSlot(c As Object, key As String) As String
    If Not c Is Nothing Then
        If c.Exists(key) Then CStrSlot = CStr(c(key))
    End If
End Function

' Join a JSON array slot (Collection) into newline-bulleted text.
Private Function JoinList(c As Object, key As String, bulletPrefix As String) As String
    Dim out As String
    If c Is Nothing Then Exit Function
    If Not c.Exists(key) Then Exit Function
    Dim arr As Object: Set arr = c(key)
    Dim i As Long
    For i = 1 To arr.Count
        If i > 1 Then out = out & vbCrLf
        out = out & bulletPrefix & CStr(arr(i))
    Next i
    JoinList = out
End Function

Public Sub Do_apply_template_act(act As Object)
    Dim tpl As String
    If act.Exists("template") Then tpl = LCase(CStr(act("template")))
    Dim slideNum As Long: slideNum = TargetSlide(act)
    Dim c As Object
    If act.Exists("content") Then Set c = act("content")

    Select Case tpl
        Case "title":         BuildTitle slideNum, c
        Case "section":       BuildSection slideNum, c
        Case "bullets":       BuildBullets slideNum, c
        Case "two_col":       BuildTwoCol slideNum, c
        Case "comparison":    BuildComparison slideNum, c
        Case "kpi_dashboard": BuildKpi slideNum, c
        Case "quote":         BuildQuote slideNum, c
        Case Else
            Err.Raise vbObjectError + 9102, "apply_template", _
                "unknown template: " & tpl
    End Select
End Sub

Private Sub BuildTitle(sl As Long, c As Object)
    Dim w As Single: w = SW() - 2 * M
    modActionsLayout.Do_add_text_box sl, CStrSlot(c, "title"), _
        M, SH() * 0.34, w, SH() * 0.2, "tpl_title", C_DARK, 40, True, False, "left"
    modActionsLayout.Do_add_text_box sl, CStrSlot(c, "subtitle"), _
        M, SH() * 0.56, w, SH() * 0.12, "tpl_subtitle", C_ACCENT, 18, False, False, "left"
End Sub

Private Sub BuildSection(sl As Long, c As Object)
    Dim w As Single: w = SW() - 2 * M
    modActionsLayout.Do_add_text_box sl, CStrSlot(c, "section_number"), _
        M, SH() * 0.28, w, SH() * 0.16, "tpl_section_number", C_ACCENT, 32, True, False, "left"
    modActionsLayout.Do_add_text_box sl, CStrSlot(c, "section_title"), _
        M, SH() * 0.46, w, SH() * 0.2, "tpl_section_title", C_DARK, 36, True, False, "left"
End Sub

Private Sub BuildBullets(sl As Long, c As Object)
    Dim w As Single: w = SW() - 2 * M
    modActionsLayout.Do_add_text_box sl, CStrSlot(c, "heading"), _
        M, M, w, SH() * 0.16, "tpl_heading", C_DARK, 28, True, False, "left"
    modActionsLayout.Do_add_text_box sl, JoinList(c, "bullets", ChrW(8226) & "  "), _
        M, M + SH() * 0.16 + GAP, w, SH() - (M + SH() * 0.16 + GAP) - M, _
        "tpl_body", C_BODY, 14, False, False, "left"
End Sub

Private Sub BuildTwoCol(sl As Long, c As Object)
    Dim w As Single: w = SW() - 2 * M
    modActionsLayout.Do_add_text_box sl, CStrSlot(c, "heading"), _
        M, M, w, SH() * 0.16, "tpl_heading", C_DARK, 28, True, False, "left"
    Dim colW As Single: colW = (w - GAP) / 2
    Dim y As Single: y = M + SH() * 0.16 + GAP
    Dim hgt As Single: hgt = SH() - y - M
    modActionsLayout.Do_add_text_box sl, CStrSlot(c, "left_body"), _
        M, y, colW, hgt, "tpl_left", C_BODY, 14, False, False, "left"
    modActionsLayout.Do_add_text_box sl, CStrSlot(c, "right_body"), _
        M + colW + GAP, y, colW, hgt, "tpl_right", C_BODY, 14, False, False, "left"
End Sub

Private Sub BuildComparison(sl As Long, c As Object)
    Dim w As Single: w = SW() - 2 * M
    modActionsLayout.Do_add_text_box sl, CStrSlot(c, "heading"), _
        M, M, w, SH() * 0.16, "tpl_heading", C_DARK, 28, True, False, "left"
    Dim colW As Single: colW = (w - GAP) / 2
    Dim y As Single: y = M + SH() * 0.16 + GAP
    Dim hgt As Single: hgt = SH() - y - M
    ' left panel
    modActionsLayout.Do_add_shape sl, "rrect", _
        M, y, colW, hgt, C_NEUTRAL, "", 0, "tpl_left_panel"
    modActionsLayout.Do_add_text_box sl, CStrSlot(c, "left_label"), _
        M + 10, y + 10, colW - 20, 28, "tpl_left_label", C_ACCENT, 16, True, False, "left"
    modActionsLayout.Do_add_text_box sl, CStrSlot(c, "left_body"), _
        M + 10, y + 44, colW - 20, hgt - 54, "tpl_left_body", C_BODY, 13, False, False, "left"
    ' right panel
    Dim rx As Single: rx = M + colW + GAP
    modActionsLayout.Do_add_shape sl, "rrect", _
        rx, y, colW, hgt, C_NEUTRAL, "", 0, "tpl_right_panel"
    modActionsLayout.Do_add_text_box sl, CStrSlot(c, "right_label"), _
        rx + 10, y + 10, colW - 20, 28, "tpl_right_label", C_ACCENT, 16, True, False, "left"
    modActionsLayout.Do_add_text_box sl, CStrSlot(c, "right_body"), _
        rx + 10, y + 44, colW - 20, hgt - 54, "tpl_right_body", C_BODY, 13, False, False, "left"
End Sub

Private Sub BuildKpi(sl As Long, c As Object)
    Dim w As Single: w = SW() - 2 * M
    modActionsLayout.Do_add_text_box sl, CStrSlot(c, "heading"), _
        M, M, w, SH() * 0.16, "tpl_heading", C_DARK, 28, True, False, "left"
    Dim tiles As Object
    If Not c Is Nothing Then
        If c.Exists("tiles") Then Set tiles = c("tiles")
    End If
    If tiles Is Nothing Then Exit Sub
    Dim n As Long: n = tiles.Count
    If n < 1 Then Exit Sub
    Dim y As Single: y = M + SH() * 0.16 + GAP
    Dim hgt As Single: hgt = SH() - y - M
    Dim tileW As Single: tileW = (w - GAP * (n - 1)) / n
    Dim i As Long
    For i = 1 To n
        Dim t As Object: Set t = tiles(i)
        Dim x As Single: x = M + (i - 1) * (tileW + GAP)
        modActionsLayout.Do_add_shape sl, "rrect", _
            x, y, tileW, hgt, C_NEUTRAL, "", 0, "tpl_tile_" & i
        modActionsLayout.Do_add_text_box sl, CStr(t("stat")), _
            x + 8, y + hgt * 0.2, tileW - 16, hgt * 0.4, _
            "tpl_tile_stat_" & i, C_ACCENT, 36, True, False, "center"
        modActionsLayout.Do_add_text_box sl, CStr(t("label")), _
            x + 8, y + hgt * 0.62, tileW - 16, hgt * 0.3, _
            "tpl_tile_label_" & i, C_BODY, 12, False, False, "center"
    Next i
End Sub

Private Sub BuildQuote(sl As Long, c As Object)
    Dim w As Single: w = SW() - 2 * M
    modActionsLayout.Do_add_text_box sl, ChrW(8220) & CStrSlot(c, "quote_text") & ChrW(8221), _
        M, SH() * 0.3, w, SH() * 0.35, "tpl_quote", C_DARK, 30, True, False, "left"
    modActionsLayout.Do_add_text_box sl, ChrW(8212) & " " & CStrSlot(c, "attribution"), _
        M, SH() * 0.68, w, SH() * 0.12, "tpl_attribution", C_ACCENT, 16, False, True, "left"
End Sub
