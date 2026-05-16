Attribute VB_Name = "modActionsSpec"
Option Explicit

' =============================================================================
' Decks as code + variants. Three actions on top of apply_template:
'   build_deck_from_spec : spec JSON -> whole deck
'   extract_spec         : live deck -> spec JSON (reads tpl_* shape tags
'                          the template builders stamp; exact round-trip
'                          for template-built decks, best-effort otherwise)
'   generate_variants    : template+content+n -> n distinct candidate
'                          slides (fixed per-index preset table)
' Slide rendering is delegated to modActionsTemplate (no duplication).
' =============================================================================

Private Const M As Single = 36#
Private Const BUL As String = "" ' set at runtime (ChrW 8226)
Private Const C_ACC As String = "#2E75B6"
Private Const C_BODY As String = "#333333"
Private Const C_NEU As String = "#F2F2F2"

Private Function Bullet() As String
    Bullet = ChrW(8226) & "  "
End Function

' ---- build_deck_from_spec ---------------------------------------------------
Public Sub Do_build_deck_from_spec_act(act As Object)
    modActionsCapture.SetActiveRegistry act
    Dim spec As Object: Set spec = act("spec")
    Dim deck As Object: Set deck = spec("deck")
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim firstNew As Long: firstNew = pres.Slides.Count + 1

    Dim i As Long
    For i = 1 To deck.Count
        Dim e As Object: Set e = deck(i)
        Dim sl As Long: sl = modActionsTemplate.AppendBlankSlide()
        Dim c As Object
        If e.Exists("content") Then Set c = e("content")
        modActionsTemplate.RenderTemplate sl, CStr(e("template")), c
    Next i

    ' Optional: drop pre-existing slides so the spec defines the whole deck.
    If act.Exists("clear_existing") Then
        If modActions.ToBool(act("clear_existing")) Then
            Dim k As Long
            For k = firstNew - 1 To 1 Step -1
                pres.Slides(k).Delete
            Next k
        End If
    End If
End Sub

' ---- extract_spec -----------------------------------------------------------
Public Sub Do_extract_spec_act(act As Object)
    Dim js As String: js = ExtractDeckSpecJson()
    On Error Resume Next
    Dim path As String: path = ActivePresentation.FullName & ".spec.json"
    Dim fnum As Integer: fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, js
    Close #fnum
    On Error GoTo 0
End Sub

' Public so the harness can call it via Application.Run (like BuildSnapshotJson).
Public Function ExtractDeckSpecJson() As String
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim deck As Object: Set deck = New Collection
    Dim s As Long
    For s = 1 To pres.Slides.Count
        Dim entry As Object
        Set entry = ClassifySlide(pres.Slides(s))
        If Not entry Is Nothing Then deck.Add entry
    Next s
    Dim root As Object: Set root = CreateObject("Scripting.Dictionary")
    root.Add "deck", deck
    ExtractDeckSpecJson = modJSON.ConvertToJson(root)
End Function

' Read tpl_* shape tags stamped by the template builders, infer the
' template, rebuild the content dict. Returns Nothing for slides with no
' tpl_* shapes (non-template slides are skipped from the spec).
Private Function ClassifySlide(sl As Object) As Object
    Dim t As Object: Set t = CreateObject("Scripting.Dictionary") ' name -> text
    Dim sh As Object
    For Each sh In sl.Shapes
        Dim nm As String: nm = sh.Name
        If Left(nm, 4) = "tpl_" Then
            Dim tx As String: tx = ""
            On Error Resume Next
            If sh.HasTextFrame Then tx = sh.TextFrame.TextRange.Text
            On Error GoTo 0
            If Not t.Exists(nm) Then t.Add nm, tx
        End If
    Next sh
    If t.Count = 0 Then Exit Function

    Dim tpl As String, c As Object
    Set c = CreateObject("Scripting.Dictionary")

    If t.Exists("tpl_title") And t.Exists("tpl_subtitle") Then
        tpl = "title"
        c.Add "title", t("tpl_title")
        c.Add "subtitle", t("tpl_subtitle")
    ElseIf t.Exists("tpl_section_number") Then
        tpl = "section"
        c.Add "section_number", t("tpl_section_number")
        c.Add "section_title", t("tpl_section_title")
    ElseIf t.Exists("tpl_left_panel") Or t.Exists("tpl_left_label") Then
        tpl = "comparison"
        c.Add "heading", t("tpl_heading")
        c.Add "left_label", t("tpl_left_label")
        c.Add "left_body", t("tpl_left_body")
        c.Add "right_label", t("tpl_right_label")
        c.Add "right_body", t("tpl_right_body")
    ElseIf t.Exists("tpl_left") And t.Exists("tpl_right") Then
        tpl = "two_col"
        c.Add "heading", t("tpl_heading")
        c.Add "left_body", t("tpl_left")
        c.Add "right_body", t("tpl_right")
    ElseIf t.Exists("tpl_tile_stat_1") Then
        tpl = "kpi_dashboard"
        c.Add "heading", t("tpl_heading")
        Dim tiles As Object: Set tiles = New Collection
        Dim n As Long: n = 1
        Do While t.Exists("tpl_tile_stat_" & n)
            Dim ti As Object: Set ti = CreateObject("Scripting.Dictionary")
            ti.Add "stat", t("tpl_tile_stat_" & n)
            ti.Add "label", t("tpl_tile_label_" & n)
            tiles.Add ti
            n = n + 1
        Loop
        c.Add "tiles", tiles
    ElseIf t.Exists("tpl_quote") Then
        tpl = "quote"
        c.Add "quote_text", StripWrap(t("tpl_quote"), ChrW(8220), ChrW(8221))
        c.Add "attribution", LTrimStr(t("tpl_attribution"), ChrW(8212) & " ")
    ElseIf t.Exists("tpl_body") Then
        tpl = "bullets"
        c.Add "heading", t("tpl_heading")
        c.Add "bullets", SplitBullets(t("tpl_body"))
    Else
        Exit Function
    End If

    Dim entry As Object: Set entry = CreateObject("Scripting.Dictionary")
    entry.Add "template", tpl
    entry.Add "content", c
    Set ClassifySlide = entry
End Function

Private Function SplitBullets(body As String) As Object
    Dim col As Object: Set col = New Collection
    ' PowerPoint stores paragraph breaks as a bare CR (vbCr); normalise
    ' CRLF/CR/LF to LF before splitting so each bullet is its own line.
    Dim nb As String
    nb = Replace(Replace(body, vbCrLf, vbLf), vbCr, vbLf)
    Dim parts() As String: parts = Split(nb, vbLf)
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim ln As String: ln = parts(i)
        If Left(ln, Len(Bullet())) = Bullet() Then ln = Mid(ln, Len(Bullet()) + 1)
        If Len(Trim(ln)) > 0 Then col.Add ln
    Next i
    Set SplitBullets = col
End Function

Private Function StripWrap(s As String, a As String, b As String) As String
    Dim r As String: r = s
    If Left(r, 1) = a Then r = Mid(r, 2)
    If Right(r, 1) = b Then r = Left(r, Len(r) - 1)
    StripWrap = r
End Function

Private Function LTrimStr(s As String, pre As String) As String
    If Left(s, Len(pre)) = pre Then
        LTrimStr = Mid(s, Len(pre) + 1)
    Else
        LTrimStr = s
    End If
End Function

' ---- generate_variants ------------------------------------------------------
Public Sub Do_generate_variants_act(act As Object)
    modActionsCapture.SetActiveRegistry act
    Dim c As Object
    If act.Exists("content") Then Set c = act("content")

    ' Form 1: templates:[names] -> render the same content across each
    ' named template (builtin OR captured) so the user picks.
    If act.Exists("templates") Then
        Dim names As Object: Set names = act("templates")
        Dim j As Long
        For j = 1 To names.Count
            Dim slj As Long: slj = modActionsTemplate.AppendBlankSlide()
            modActionsTemplate.RenderTemplate slj, CStr(names(j)), c
        Next j
        Exit Sub
    End If

    ' Form 2: template + n -> n preset perturbations (builtin layouts).
    Dim tpl As String: tpl = LCase(CStr(act("template")))
    Dim n As Long: n = CLng(act("n"))
    Dim i As Long
    For i = 1 To n
        Dim sl As Long: sl = modActionsTemplate.AppendBlankSlide()
        RenderVariant sl, tpl, c, i
    Next i
End Sub

' A deliberately distinct layout per preset index (measurably different
' geometry + accent). Carries the same slot text so the user/LLM can pick.
' Principled layout archetypes (NOT cosmetic shuffles). Each is a
' genuinely different information architecture obeying the design rules:
'   >=M margins, ONE strictly-dominant (largest-font) element, <=2 fill
'   colours, single accent, clear size tiers. n cycles the set so the
'   user sees structurally different options of the SAME content.
Private Sub RenderVariant(sl As Long, tpl As String, c As Object, idx As Long)
    Dim sw As Single: sw = ActivePresentation.PageSetup.SlideWidth
    Dim sh As Single: sh = ActivePresentation.PageSetup.SlideHeight
    Dim heading As String: heading = SlotOr(c, Array("title", "heading", _
        "section_title", "quote_text"))
    Dim allv As String: allv = AllValues(c)

    Select Case ((idx - 1) Mod 5)
        Case 0: Arch_Hero sl, sw, sh, heading, allv, idx
        Case 1: Arch_Split sl, sw, sh, heading, allv, idx
        Case 2: Arch_Stack sl, sw, sh, heading, allv, idx
        Case 3: Arch_Quote sl, sw, sh, heading, allv, idx
        Case 4: Arch_Tiles sl, sw, sh, heading, allv, idx
    End Select
End Sub

Private Sub Arch_Hero(sl As Long, sw As Single, sh As Single, _
                      h As String, v As String, idx As Long)
    ' One dominant statement; everything else recedes.
    modActionsLayout.Do_add_text_box sl, h, M, sh * 0.32, sw - 2 * M, _
        sh * 0.26, "var_head_" & idx, "#15283C", 60, True, False, "center"
    modActionsLayout.Do_add_text_box sl, v, M, sh * 0.62, sw - 2 * M, _
        sh * 0.22, "var_sub_" & idx, C_BODY, 16, False, False, "center"
End Sub

Private Sub Arch_Split(sl As Long, sw As Single, sh As Single, _
                       h As String, v As String, idx As Long)
    modActionsLayout.Do_add_text_box sl, h, M, M, sw - 2 * M, sh * 0.14, _
        "var_head_" & idx, "#15283C", 28, True, False, "left"
    Dim y As Single: y = sh * 0.26
    Dim hgt As Single: hgt = sh - y - M
    Dim cw As Single: cw = (sw - 2 * M - 24) / 2
    modActionsLayout.Do_add_shape sl, "rrect", M, y, cw, hgt, _
        C_NEU, "", 0, "var_pL_" & idx
    modActionsLayout.Do_add_shape sl, "rrect", M + cw + 24, y, cw, hgt, _
        C_NEU, "", 0, "var_pR_" & idx
    modActionsLayout.Do_add_text_box sl, v, M + 12, y + 12, cw - 24, _
        hgt - 24, "var_bL_" & idx, C_BODY, 14, False, False, "left"
    modActionsLayout.Do_add_text_box sl, v, M + cw + 36, y + 12, cw - 24, _
        hgt - 24, "var_bR_" & idx, C_BODY, 14, False, False, "left"
End Sub

Private Sub Arch_Stack(sl As Long, sw As Single, sh As Single, _
                       h As String, v As String, idx As Long)
    modActionsLayout.Do_add_text_box sl, h, M, M, sw - 2 * M, sh * 0.13, _
        "var_head_" & idx, "#15283C", 28, True, False, "left"
    Dim parts() As String: parts = Split(v, vbCrLf)
    Dim y As Single: y = sh * 0.24
    Dim rowH As Single: rowH = (sh - y - M - 2 * 16) / 3
    Dim k As Long
    For k = 0 To 2
        Dim ry As Single: ry = y + k * (rowH + 16)
        modActionsLayout.Do_add_shape sl, "rrect", M, ry, sw - 2 * M, rowH, _
            C_NEU, "", 0, "var_row" & k & "_" & idx
        Dim t As String: t = ""
        If k <= UBound(parts) Then t = parts(k)
        If k = 2 Then t = v   ' last cluster carries the full content
        modActionsLayout.Do_add_text_box sl, t, M + 14, ry + 10, _
            sw - 2 * M - 28, rowH - 20, "var_c" & k & "_" & idx, _
            C_BODY, 14, False, False, "left"
    Next k
End Sub

Private Sub Arch_Quote(sl As Long, sw As Single, sh As Single, _
                       h As String, v As String, idx As Long)
    modActionsLayout.Do_add_text_box sl, ChrW(8220) & h & ChrW(8221), _
        M, sh * 0.26, sw - 2 * M, sh * 0.4, "var_q_" & idx, "#15283C", _
        44, True, False, "left"
    modActionsLayout.Do_add_text_box sl, v, M, sh * 0.7, sw - 2 * M, _
        sh * 0.18, "var_a_" & idx, C_ACC, 16, False, True, "left"
End Sub

Private Sub Arch_Tiles(sl As Long, sw As Single, sh As Single, _
                       h As String, v As String, idx As Long)
    modActionsLayout.Do_add_text_box sl, h, M, M, sw - 2 * M, sh * 0.13, _
        "var_head_" & idx, "#15283C", 30, True, False, "left"
    Dim parts() As String: parts = Split(v, vbCrLf)
    Dim n As Long: n = 4
    Dim y As Single: y = sh * 0.24
    Dim hgt As Single: hgt = sh - y - M
    Dim tw As Single: tw = (sw - 2 * M - 3 * 18) / n
    Dim i As Long
    For i = 0 To n - 1
        Dim x As Single: x = M + i * (tw + 18)
        modActionsLayout.Do_add_shape sl, "rrect", x, y, tw, hgt, _
            C_NEU, "", 0, "var_t" & i & "_" & idx
        Dim s As String: s = ""
        If i <= UBound(parts) Then s = parts(i)
        If i = 0 Then s = v
        modActionsLayout.Do_add_text_box sl, s, x + 8, y + hgt * 0.18, _
            tw - 16, hgt * 0.5, "var_ts" & i & "_" & idx, C_ACC, 20, _
            True, False, "center"
    Next i
End Sub

Private Function SlotOr(c As Object, keys As Variant) As String
    If c Is Nothing Then Exit Function
    Dim k As Variant
    For Each k In keys
        If c.Exists(CStr(k)) Then
            SlotOr = CStr(c(CStr(k)))
            Exit Function
        End If
    Next k
End Function

' Concatenate every scalar slot value (and array items) so all slot text
' is present on each variant for deterministic verification.
Private Function AllValues(c As Object) As String
    If c Is Nothing Then Exit Function
    Dim out As String
    Dim k As Variant
    For Each k In c.Keys
        Dim v As Variant
        If IsObject(c(k)) Then
            Dim it As Variant
            For Each it In c(k)
                If IsObject(it) Then
                    Dim kk As Variant
                    For Each kk In it.Keys
                        out = out & CStr(it(kk)) & vbCrLf
                    Next kk
                Else
                    out = out & CStr(it) & vbCrLf
                End If
            Next it
        Else
            out = out & CStr(c(k)) & vbCrLf
        End If
    Next k
    AllValues = out
End Function

' ---- shared validation helper (called by ValidateAction) --------------------
Public Function ValidateSpec(spec As Object) As String
    If spec Is Nothing Then ValidateSpec = "spec: required": Exit Function
    If Not spec.Exists("deck") Then ValidateSpec = "spec.deck: required": Exit Function
    Dim deck As Object: Set deck = spec("deck")
    If deck Is Nothing Then ValidateSpec = "spec.deck: must be a non-empty array": Exit Function
    If deck.Count = 0 Then ValidateSpec = "spec.deck: must be a non-empty array": Exit Function
    Dim i As Long
    For i = 1 To deck.Count
        Dim e As Object: Set e = deck(i)
        If Not e.Exists("template") Then
            ValidateSpec = "spec.deck[" & i & "].template: required"
            Exit Function
        End If
        ' Template name may be a builtin OR a user-captured name; existence
        ' is a runtime/registry concern (surfaced via FAILURES), not a
        ' validation error -- so no enum restriction here.
    Next i
End Function
