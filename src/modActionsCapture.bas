Attribute VB_Name = "modActionsCapture"
Option Explicit

' Module-level state (declarations section — must precede all procedures).
Private gActiveReg As String

' =============================================================================
' User-captured templates ("Deck DNA"). Templates are DATA, not code:
'   (a) registry = external JSON file (%APPDATA%\Decko\templates.json),
'       overridable per-call via act("registry_path") for tests.
'   (b) ONE generic renderer (RenderCapturedTemplate) replays a registry
'       entry's stored shapes + substitutes slot text — no per-template code.
'   (c) BuildCapturedManifest() is read live by frmExport.PromptTemplate so
'       the LLM is told about your templates without any code change.
' Capture auto-infers slots (every text shape) so you never tag anything.
' =============================================================================

Public Function DefaultRegistryPath() As String
    DefaultRegistryPath = Environ$("APPDATA") & "\Decko\templates.json"
End Function

Private Function RegistryPath(act As Object) As String
    If Not act Is Nothing Then
        If act.Exists("registry_path") Then
            RegistryPath = CStr(act("registry_path"))
            Exit Function
        End If
    End If
    RegistryPath = DefaultRegistryPath()
End Function

' Active registry path for the current action chain, so apply_template /
' build_deck_from_spec / generate_variants can resolve captured names
' without threading a path through every signature. Set at action entry.
Public Sub SetActiveRegistry(act As Object)
    gActiveReg = RegistryPath(act)
End Sub

Public Function GetActiveRegistry() As String
    If Len(gActiveReg) = 0 Then gActiveReg = DefaultRegistryPath()
    GetActiveRegistry = gActiveReg
End Function

' ---- registry IO ------------------------------------------------------------
Public Function LoadRegistry(path As String) As Object
    Dim reg As Object
    On Error Resume Next
    If Len(Dir(path)) > 0 Then
        Dim fnum As Integer: fnum = FreeFile
        Dim raw As String
        Open path For Input As #fnum
        If LOF(fnum) > 0 Then raw = Input$(LOF(fnum), fnum)
        Close #fnum
        Set reg = modJSON.ParseJson(raw)
    End If
    On Error GoTo 0
    If reg Is Nothing Then
        Set reg = CreateObject("Scripting.Dictionary")
    End If
    If Not reg.Exists("templates") Then
        reg.Add "templates", CreateObject("Scripting.Dictionary")
    End If
    Set LoadRegistry = reg
End Function

Private Sub SaveRegistry(path As String, reg As Object)
    Dim folder As String
    folder = Left$(path, InStrRev(path, "\") - 1)
    On Error Resume Next
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    On Error GoTo 0
    Dim fnum As Integer: fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, modJSON.ConvertToJson(reg)
    Close #fnum
End Sub

' ---- capture ----------------------------------------------------------------
Public Sub Do_capture_template_act(act As Object)
    Dim nm As String: nm = CStr(act("name"))
    Dim slideNum As Long
    If act.Exists("slide") Then
        slideNum = CLng(act("slide"))
    Else
        slideNum = Application.ActiveWindow.View.Slide.SlideIndex
    End If
    Dim sl As Object: Set sl = ActivePresentation.Slides(slideNum)

    Dim shapesCol As Object: Set shapesCol = New Collection
    Dim slotShapes As Object: Set slotShapes = New Collection
    Dim sh As Object
    For Each sh In sl.Shapes
        Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
        d.Add "left", CDbl(sh.Left)
        d.Add "top", CDbl(sh.Top)
        d.Add "width", CDbl(sh.Width)
        d.Add "height", CDbl(sh.Height)

        Dim hasText As Boolean: hasText = False
        Dim txt As String: txt = ""
        On Error Resume Next
        If sh.HasTextFrame Then
            txt = sh.TextFrame.TextRange.Text
            hasText = (Len(Trim(txt)) > 0)
        End If
        On Error GoTo 0
        d.Add "text", txt

        ' kind: textbox vs autoshape (mso_N), else generic box
        Dim kind As String: kind = "rect"
        On Error Resume Next
        If sh.Type = msoTextBox Then
            kind = "textbox"
        ElseIf sh.Type = msoAutoShape Then
            kind = "mso_" & CStr(sh.AutoShapeType)
        End If
        On Error GoTo 0
        d.Add "kind", kind

        Dim fSize As Double: fSize = 0
        Dim fBold As Boolean: fBold = False
        Dim fItalic As Boolean: fItalic = False
        Dim fColor As String: fColor = ""
        If hasText Then
            On Error Resume Next
            fSize = CDbl(sh.TextFrame.TextRange.Font.Size)
            fBold = (sh.TextFrame.TextRange.Font.Bold <> False)
            fItalic = (sh.TextFrame.TextRange.Font.Italic <> False)
            fColor = modVerify.RgbToHex(sh.TextFrame.TextRange.Font.Color.RGB)
            On Error GoTo 0
        End If
        d.Add "font_size", fSize
        d.Add "font_bold", fBold
        d.Add "font_italic", fItalic
        d.Add "font_color", fColor

        Dim fillHex As String: fillHex = ""
        On Error Resume Next
        If sh.Fill.Visible = msoTrue Then fillHex = modVerify.RgbToHex(sh.Fill.ForeColor.RGB)
        On Error GoTo 0
        d.Add "fill", fillHex

        If hasText Then
            d.Add "role", "slot"
            slotShapes.Add d
        Else
            d.Add "role", "deco"
        End If
        shapesCol.Add d
    Next sh

    ' Auto-name slots: largest font (tie -> topmost) = heading; largest
    ' remaining area = body; rest = slot_N in reading order.
    Dim slots As Object: Set slots = New Collection
    AssignSlotNames slotShapes, slots

    Dim tplDict As Object: Set tplDict = CreateObject("Scripting.Dictionary")
    tplDict.Add "slots", slots
    tplDict.Add "shapes", shapesCol

    Dim path As String: path = RegistryPath(act)
    Dim reg As Object: Set reg = LoadRegistry(path)
    Dim templates As Object: Set templates = reg("templates")
    If templates.Exists(nm) Then templates.Remove nm
    templates.Add nm, tplDict
    SaveRegistry path, reg
End Sub

Private Sub AssignSlotNames(slotShapes As Object, ByRef slots As Object)
    Dim n As Long: n = slotShapes.Count
    If n = 0 Then Exit Sub
    ' heading = max font size, tie-break topmost
    Dim headIdx As Long: headIdx = 1
    Dim i As Long
    For i = 2 To n
        Dim a As Object: Set a = slotShapes(i)
        Dim b As Object: Set b = slotShapes(headIdx)
        If a("font_size") > b("font_size") Or _
           (a("font_size") = b("font_size") And a("top") < b("top")) Then
            headIdx = i
        End If
    Next i
    ' body = max area among the rest
    Dim bodyIdx As Long: bodyIdx = 0
    Dim bestArea As Double: bestArea = -1
    For i = 1 To n
        If i <> headIdx Then
            Dim s As Object: Set s = slotShapes(i)
            Dim ar As Double: ar = s("width") * s("height")
            If ar > bestArea Then bestArea = ar: bodyIdx = i
        End If
    Next i
    Dim ctr As Long: ctr = 3
    For i = 1 To n
        Dim nm As String
        If i = headIdx Then
            nm = "heading"
        ElseIf i = bodyIdx Then
            nm = "body"
        Else
            nm = "slot_" & ctr
            ctr = ctr + 1
        End If
        slotShapes(i).Add "slot", nm
        slots.Add nm
    Next i
End Sub

' ---- generic renderer (ONE path for ANY captured template) ------------------
Public Sub RenderCapturedTemplate(slideNum As Long, nm As String, _
                                   content As Object, registryPath As String)
    Dim reg As Object: Set reg = LoadRegistry(registryPath)
    Dim templates As Object: Set templates = reg("templates")
    If Not templates.Exists(nm) Then
        Err.Raise vbObjectError + 9301, "capture", "no captured template: " & nm
    End If
    Dim tpl As Object: Set tpl = templates(nm)
    Dim shapesCol As Object: Set shapesCol = tpl("shapes")
    Dim i As Long
    For i = 1 To shapesCol.Count
        Dim s As Object: Set s = shapesCol(i)
        Dim txt As String: txt = CStr(s("text"))
        If CStr(s("role")) = "slot" And Not content Is Nothing Then
            Dim slotName As String: slotName = CStr(s("slot"))
            If content.Exists(slotName) Then txt = CStr(content(slotName))
        End If
        Dim kind As String: kind = CStr(s("kind"))
        If kind = "textbox" Then
            modActionsLayout.Do_add_text_box slideNum, txt, _
                CSng(s("left")), CSng(s("top")), CSng(s("width")), CSng(s("height")), _
                "", CStr(s("font_color")), CLng(s("font_size")), _
                CBool(s("font_bold")), CBool(s("font_italic")), "left", _
                CStr(s("fill"))
        Else
            modActionsLayout.Do_add_shape slideNum, kind, _
                CSng(s("left")), CSng(s("top")), CSng(s("width")), CSng(s("height")), _
                CStr(s("fill")), "", 0, "", txt, CStr(s("font_color")), _
                CLng(s("font_size")), CBool(s("font_bold"))
        End If
    Next i
End Sub

Public Function HasCaptured(nm As String, registryPath As String) As Boolean
    Dim reg As Object: Set reg = LoadRegistry(registryPath)
    HasCaptured = reg("templates").Exists(nm)
End Function

' ---- list / delete / rename -------------------------------------------------
Public Sub Do_list_templates_act(act As Object)
    ' Side-effect free for the deck; writes a readable list next to the deck.
    On Error Resume Next
    Dim p As String: p = ActivePresentation.FullName & ".templates_list.txt"
    Dim fnum As Integer: fnum = FreeFile
    Open p For Output As #fnum
    Print #fnum, BuildCapturedManifest(RegistryPath(act))
    Close #fnum
    On Error GoTo 0
End Sub

Public Sub Do_delete_template_act(act As Object)
    Dim path As String: path = RegistryPath(act)
    Dim reg As Object: Set reg = LoadRegistry(path)
    Dim nm As String: nm = CStr(act("name"))
    If reg("templates").Exists(nm) Then reg("templates").Remove nm
    SaveRegistry path, reg
End Sub

Public Sub Do_rename_template_act(act As Object)
    Dim path As String: path = RegistryPath(act)
    Dim reg As Object: Set reg = LoadRegistry(path)
    Dim templates As Object: Set templates = reg("templates")
    Dim fromN As String: fromN = CStr(act("from"))
    Dim toN As String: toN = CStr(act("to"))
    If templates.Exists(fromN) Then
        Dim v As Object: Set v = templates(fromN)
        templates.Remove fromN
        If templates.Exists(toN) Then templates.Remove toN
        templates.Add toN, v
        SaveRegistry path, reg
    End If
End Sub

' ---- live manifest (read by frmExport.PromptTemplate) -----------------------
Public Function BuildCapturedManifest(registryPath As String) As String
    Dim reg As Object: Set reg = LoadRegistry(registryPath)
    Dim templates As Object: Set templates = reg("templates")
    If templates.Count = 0 Then Exit Function
    Dim sb As String
    sb = vbCrLf & "=== YOUR CAPTURED TEMPLATES ===" & vbCrLf
    sb = sb & "Use these names with apply_template / build_deck_from_spec / " & _
         "generate_variants(""templates"":[..])." & vbCrLf
    Dim k As Variant
    For Each k In templates.Keys
        Dim tpl As Object: Set tpl = templates(k)
        Dim slots As Object: Set slots = tpl("slots")
        Dim line As String: line = "  " & CStr(k) & "  slots: "
        Dim i As Long
        For i = 1 To slots.Count
            If i > 1 Then line = line & ", "
            line = line & CStr(slots(i))
        Next i
        sb = sb & line & vbCrLf
    Next k
    sb = sb & "EXAMPLE: {""type"":""apply_template"",""template"":""<name>""," & _
         """content"":{""heading"":""..""}}" & vbCrLf
    BuildCapturedManifest = sb
End Function
