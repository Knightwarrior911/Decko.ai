Attribute VB_Name = "modActionsIcon"
Option Explicit

' Fetches Microsoft Fluent UI System Icons from unpkg CDN, recolors SVG text,
' caches to %TEMP%\decko_icons\, and inserts via AddPicture (PPT 2019+ renders SVG).
'
' Action schema:
'   {"type":"insert_icon",
'    "slide": 1,
'    "icon": "factory",          -- icon name, e.g. factory / airplane / people / globe
'    "style": "filled",          -- "filled" (default) or "regular"
'    "size": 48,                 -- 16|20|24|28|32|48 (default 48)
'    "color": "#15283C",         -- hex color applied to SVG fill (optional)
'    "left": 100, "top": 200,    -- position in points
'    "width": 48, "height": 48,  -- size in points
'    "ref_name": "icon_factory"} -- optional shape name

Private Const UNPKG_BASE As String = "https://unpkg.com/@fluentui/svg-icons/icons/"
Private Const CACHE_DIR_NAME As String = "decko_icons"

Public Sub Do_insert_icon(act As Object)
    Dim slideNum As Long: slideNum = CLng(act("slide"))
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then slideNum = pres.Slides.Count

    Dim iconName As String
    iconName = LCase(Replace(CStr(act("icon")), " ", "_"))

    Dim style As String: style = "filled"
    If act.Exists("style") Then style = LCase(CStr(act("style")))

    Dim sz As Long: sz = 48
    If act.Exists("size") Then sz = CLng(act("size"))

    Dim color As String: color = ""
    If act.Exists("color") Then color = CStr(act("color"))

    Dim leftPt As Single:  leftPt = CSng(act("left"))
    Dim topPt As Single:   topPt  = CSng(act("top"))
    Dim wPt As Single:     wPt    = CSng(act("width"))
    Dim hPt As Single:     hPt    = CSng(act("height"))

    Dim svgPath As String: svgPath = GetOrDownloadIcon(iconName, sz, style, color)

    Dim sl As Slide: Set sl = pres.Slides(slideNum)
    Dim pic As Shape
    On Error Resume Next
    Set pic = sl.Shapes.AddPicture(FileName:=svgPath, _
        LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=leftPt, Top:=topPt, Width:=wPt, Height:=hPt)
    If Err.Number <> 0 Or pic Is Nothing Then
        Dim addErr As String: addErr = Err.Description
        Err.Clear
        On Error GoTo 0
        Err.Raise vbObjectError + 7002, "Do_insert_icon", _
            "AddPicture failed (" & addErr & "). PowerPoint 2019+ required for SVG."
    End If
    On Error GoTo 0

    pic.LockAspectRatio = msoFalse
    pic.Width = wPt
    pic.Height = hPt

    If act.Exists("ref_name") Then pic.Name = CStr(act("ref_name"))
End Sub

' Returns local path to (possibly recolored) cached SVG, downloading if needed.
Private Function GetOrDownloadIcon(iconName As String, sz As Long, _
                                    style As String, color As String) As String
    Dim cacheDir As String: cacheDir = Environ("TEMP") & "\" & CACHE_DIR_NAME
    EnsureDirExists cacheDir

    Dim baseName As String: baseName = iconName & "_" & sz & "_" & style
    Dim colorSuffix As String: colorSuffix = ""
    If Len(color) > 0 Then colorSuffix = "_" & Replace(color, "#", "")
    Dim cachedPath As String: cachedPath = cacheDir & "\" & baseName & colorSuffix & ".svg"

    If Dir(cachedPath) <> "" Then
        GetOrDownloadIcon = cachedPath
        Exit Function
    End If

    Dim url As String: url = UNPKG_BASE & baseName & ".svg"
    Dim svgText As String: svgText = modActionsWeb.HttpGetText(url)

    If Len(Trim(svgText)) = 0 Then
        Err.Raise vbObjectError + 7001, "Do_insert_icon", _
            "Icon not found on CDN: " & iconName & " (tried: " & url & ")" & vbCrLf & _
            "Check name at: https://icon.fluentui.dev"
    End If

    If Len(color) > 0 Then svgText = RecolorSvg(svgText, color)
    modActionsWeb.WriteTextFile cachedPath, svgText
    GetOrDownloadIcon = cachedPath
End Function

' Fluent icons use fill="currentColor". Replace with the target hex color.
Private Function RecolorSvg(svgText As String, hexColor As String) As String
    Dim s As String: s = svgText
    s = Replace(s, "fill=""currentColor""", "fill=""" & hexColor & """")
    s = Replace(s, "fill:currentColor", "fill:" & hexColor)
    s = Replace(s, "fill: currentColor", "fill: " & hexColor)
    RecolorSvg = s
End Function

Private Sub EnsureDirExists(path As String)
    If Dir(path, vbDirectory) = "" Then
        On Error Resume Next
        MkDir path
        On Error GoTo 0
    End If
End Sub
