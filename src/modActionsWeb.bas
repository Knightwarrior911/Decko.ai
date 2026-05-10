Attribute VB_Name = "modActionsWeb"
Option Explicit

' Web fetch + image download for Decko.
'
' Public actions (callable from JSON dispatcher):
'   Do_fetch_page_images(url, destFolder, refName)
'       Downloads the page HTML, scrapes <img src> + alt text, downloads
'       each image as a binary file under destFolder, writes manifest.json
'       listing {src_url, local_path, alt, width_px, height_px} for each.
'       Sets Public g_LastFetchFolder so the picker can find it.
'
'   Do_open_image_picker(folder)
'       Opens frmImagePicker over destFolder (defaults to g_LastFetchFolder).
'       Picker writes selection.json on OK; sets g_LastPickerSelection.
'
'   Do_download_image(url, destPath)
'       One-shot URL -> file. Used by build_image_grid_table when given a
'       URL instead of a local path.
'
' All HTTP via WinHttp.WinHttpRequest.5.1 (binary safe, follows redirects,
' modern TLS). Image regex is intentionally loose - the picker UI lets the
' user discard junk like sprites/logos/tracking pixels.

Public g_LastFetchFolder As String
Public g_LastPickerSelection As String   ' path to selection.json

Private Const USER_AGENT As String = _
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " & _
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

' ---------- Public dispatcher targets ----------------------------------------

Public Sub Do_fetch_page_images(url As String, _
                                Optional ByVal destFolder As String = "", _
                                Optional ByVal refName As String = "")
    If Len(Trim(url)) = 0 Then Err.Raise vbObjectError + 9001, "Do_fetch_page_images", "url required"

    If Len(destFolder) = 0 Then destFolder = MakeAssetFolder("page_" & SafeSlug(url))
    EnsureFolder destFolder

    Dim html As String
    html = HttpGetText(url)
    If Len(html) = 0 Then
        Err.Raise vbObjectError + 9002, "Do_fetch_page_images", _
            "page returned empty body (likely blocked, JS-only, or timeout): " & url
    End If

    ' Persist raw HTML for debugging / re-scraping
    WriteTextFile destFolder & "\page.html", html

    Dim imgs As Collection
    Set imgs = ExtractImageRefs(html, url)

    Dim manifest As New Collection
    Dim i As Long, downloaded As Long
    downloaded = 0
    For i = 1 To imgs.Count
        Dim ref As Object: Set ref = imgs(i)
        Dim absUrl As String: absUrl = CStr(ref("src"))
        If LooksLikeRealImage(absUrl) Then
            Dim ext As String: ext = GuessImageExt(absUrl)
            Dim localName As String
            localName = "img_" & Format(downloaded + 1, "000") & "." & ext
            Dim localPath As String: localPath = destFolder & "\" & localName
            If DownloadBinary(absUrl, localPath) Then
                downloaded = downloaded + 1
                Dim entry As Object: Set entry = CreateObject("Scripting.Dictionary")
                entry.Add "src_url", absUrl
                entry.Add "local_path", localPath
                entry.Add "alt", ref("alt")
                entry.Add "title", ref("title")
                entry.Add "context_text", ref("context")
                manifest.Add entry
            End If
        End If
    Next i

    Dim wrap As Object: Set wrap = CreateObject("Scripting.Dictionary")
    wrap.Add "source_url", url
    wrap.Add "folder", destFolder
    wrap.Add "image_count", downloaded
    wrap.Add "images", manifest
    WriteTextFile destFolder & "\manifest.json", modJSON.ConvertToJson(wrap)

    g_LastFetchFolder = destFolder
    Debug.Print "fetch_page_images: " & downloaded & " image(s) into " & destFolder
End Sub

Public Sub Do_open_image_picker(Optional ByVal folder As String = "")
    If Len(folder) = 0 Then folder = g_LastFetchFolder
    If Len(folder) = 0 Then
        Err.Raise vbObjectError + 9003, "Do_open_image_picker", "no folder given and g_LastFetchFolder unset"
    End If
    If Dir(folder, vbDirectory) = "" Then
        Err.Raise vbObjectError + 9004, "Do_open_image_picker", "folder does not exist: " & folder
    End If
    frmImagePicker.PickerFolder = folder
    frmImagePicker.Show
End Sub

Public Sub Do_download_image(url As String, destPath As String)
    EnsureFolder ParentFolder(destPath)
    If Not DownloadBinary(url, destPath) Then
        Err.Raise vbObjectError + 9005, "Do_download_image", "download failed: " & url
    End If
End Sub

' ---------- HTTP -------------------------------------------------------------

Public Function HttpGetText(url As String) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    On Error Resume Next
    http.Option(4) = 13056   ' SslErrorIgnoreFlags - tolerate misconfigured certs
    http.Option(6) = True    ' EnableRedirects
    ' Resolve / connect / send / receive timeouts in ms
    http.SetTimeouts 15000, 30000, 60000, 90000
    On Error GoTo 0
    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", USER_AGENT
    http.SetRequestHeader "Accept", "text/html,application/xhtml+xml,*/*"
    http.SetRequestHeader "Accept-Language", "en-US,en;q=0.9"
    http.SetRequestHeader "Accept-Encoding", "identity"   ' avoid compressed body
    On Error Resume Next
    http.Send
    If Err.Number <> 0 Then
        Debug.Print "HttpGetText send error: " & Err.Description
        HttpGetText = ""
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    Debug.Print "HttpGetText status=" & http.Status & " len=" & Len(http.ResponseText) & " url=" & url
    If http.Status >= 400 Then
        HttpGetText = ""
        Exit Function
    End If
    HttpGetText = http.ResponseText
End Function

Private Function DownloadBinary(url As String, destPath As String) As Boolean
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    On Error Resume Next
    http.Option(4) = 13056
    http.Option(6) = True
    On Error GoTo 0

    On Error Resume Next
    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", USER_AGENT
    http.SetRequestHeader "Accept", "image/*,*/*"
    http.Send
    If Err.Number <> 0 Then
        DownloadBinary = False
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    If http.Status >= 400 Then
        DownloadBinary = False
        Exit Function
    End If

    Dim body As Variant: body = http.ResponseBody
    If IsEmpty(body) Then
        DownloadBinary = False
        Exit Function
    End If

    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1     ' adTypeBinary
    stream.Open
    stream.Write body
    On Error Resume Next
    If Dir(destPath) <> "" Then Kill destPath
    On Error GoTo 0
    stream.SaveToFile destPath, 2   ' adSaveCreateOverWrite
    stream.Close

    DownloadBinary = (FileLen(destPath) > 200)   ' reject empties / 1px gifs
End Function

' ---------- HTML scraping ----------------------------------------------------

' Returns Collection of Dictionary{src, alt, title, context} with absolute URLs.
Private Function ExtractImageRefs(html As String, baseUrl As String) As Collection
    Dim out As New Collection
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    ' Capture entire <img ...> tag so we can inspect attributes.
    re.Pattern = "<img\b[^>]*>"

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")

    Dim matches As Object: Set matches = re.Execute(html)
    Dim m As Object
    For Each m In matches
        Dim tag As String: tag = m.Value
        Dim src As String: src = AttrFromTag(tag, "src")
        ' Some sites use lazy-load attrs. Fall back if src is a placeholder.
        If Len(src) = 0 Or LCase(src) Like "*data:image/svg*" Or InStr(LCase(src), "blank.gif") > 0 Then
            Dim alt2 As String
            alt2 = AttrFromTag(tag, "data-src")
            If Len(alt2) > 0 Then src = alt2
        End If
        If Len(src) = 0 Then
            Dim ds As String: ds = AttrFromTag(tag, "data-original")
            If Len(ds) > 0 Then src = ds
        End If
        If Len(src) > 0 Then
            Dim absUrl As String: absUrl = ResolveUrl(src, baseUrl)
            If Len(absUrl) > 0 And Not seen.Exists(absUrl) Then
                seen.Add absUrl, True
                Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
                d.Add "src", absUrl
                d.Add "alt", AttrFromTag(tag, "alt")
                d.Add "title", AttrFromTag(tag, "title")
                d.Add "context", ""
                out.Add d
            End If
        End If
    Next m

    ' Also capture CSS background-image: url(...) found in style attributes -
    ' modern card layouts often use bg-image instead of <img>.
    Dim re2 As Object: Set re2 = CreateObject("VBScript.RegExp")
    re2.Global = True
    re2.IgnoreCase = True
    re2.Pattern = "background(?:-image)?\s*:\s*url\(\s*['""]?([^'""\)]+)['""]?\s*\)"
    Dim matches2 As Object: Set matches2 = re2.Execute(html)
    For Each m In matches2
        Dim u As String: u = m.SubMatches(0)
        Dim au As String: au = ResolveUrl(u, baseUrl)
        If Len(au) > 0 And Not seen.Exists(au) Then
            seen.Add au, True
            Dim d2 As Object: Set d2 = CreateObject("Scripting.Dictionary")
            d2.Add "src", au
            d2.Add "alt", ""
            d2.Add "title", ""
            d2.Add "context", "css-background"
            out.Add d2
        End If
    Next m

    Set ExtractImageRefs = out
End Function

Private Function AttrFromTag(tag As String, attr As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "\b" & attr & "\s*=\s*(?:""([^""]*)""|'([^']*)'|([^\s>]+))"
    Dim m As Object: Set m = re.Execute(tag)
    If m.Count = 0 Then
        AttrFromTag = ""
        Exit Function
    End If
    Dim hit As Object: Set hit = m(0)
    Dim i As Long
    For i = 0 To 2
        If Len(hit.SubMatches(i)) > 0 Then
            AttrFromTag = hit.SubMatches(i)
            Exit Function
        End If
    Next i
    AttrFromTag = ""
End Function

Private Function ResolveUrl(href As String, baseUrl As String) As String
    Dim h As String: h = Trim(href)
    If Len(h) = 0 Then ResolveUrl = "": Exit Function
    If LCase(Left(h, 5)) = "data:" Then ResolveUrl = "": Exit Function
    If LCase(Left(h, 11)) = "javascript:" Then ResolveUrl = "": Exit Function
    If Left(h, 2) = "//" Then
        ResolveUrl = "https:" & h
        Exit Function
    End If
    If LCase(Left(h, 7)) = "http://" Or LCase(Left(h, 8)) = "https://" Then
        ResolveUrl = h
        Exit Function
    End If
    Dim scheme As String, host As String, path As String
    SplitUrl baseUrl, scheme, host, path
    If Left(h, 1) = "/" Then
        ResolveUrl = scheme & "://" & host & h
        Exit Function
    End If
    ' Relative to current path's directory
    Dim dir As String: dir = path
    Dim slash As Long: slash = InStrRev(dir, "/")
    If slash > 0 Then dir = Left(dir, slash) Else dir = "/"
    ResolveUrl = scheme & "://" & host & dir & h
End Function

Private Sub SplitUrl(url As String, ByRef scheme As String, _
                     ByRef host As String, ByRef path As String)
    scheme = "https": host = "": path = "/"
    Dim s As String: s = url
    Dim p As Long: p = InStr(s, "://")
    If p > 0 Then
        scheme = LCase(Left(s, p - 1))
        s = Mid(s, p + 3)
    End If
    Dim slash As Long: slash = InStr(s, "/")
    If slash = 0 Then
        host = s
    Else
        host = Left(s, slash - 1)
        path = Mid(s, slash)
    End If
    ' Strip query/fragment for the path-as-base purpose - keep only file portion
    Dim q As Long: q = InStr(path, "?")
    If q > 0 Then path = Left(path, q - 1)
    Dim f As Long: f = InStr(path, "#")
    If f > 0 Then path = Left(path, f - 1)
End Sub

Private Function LooksLikeRealImage(url As String) As Boolean
    Dim u As String: u = LCase(url)
    If Len(u) = 0 Then LooksLikeRealImage = False: Exit Function
    ' Strip query for extension detection
    Dim q As Long: q = InStr(u, "?")
    If q > 0 Then u = Left(u, q - 1)
    If u Like "*.jpg" Or u Like "*.jpeg" Or u Like "*.png" Or u Like "*.webp" _
       Or u Like "*.gif" Or u Like "*.bmp" Or u Like "*.tiff" Or u Like "*.tif" _
       Or u Like "*.svg" Then
        ' Skip 1x1 trackers and known sprite names
        If InStr(u, "spacer") > 0 Or InStr(u, "1x1") > 0 Or InStr(u, "pixel") > 0 _
           Or InStr(u, "tracking") > 0 Then
            LooksLikeRealImage = False
        Else
            LooksLikeRealImage = True
        End If
    Else
        ' Unknown extension - keep if URL looks like an image-ish endpoint
        If InStr(u, "/image") > 0 Or InStr(u, "/photo") > 0 Or InStr(u, "/media") > 0 _
           Or InStr(u, "format=jpg") > 0 Or InStr(u, "format=png") > 0 _
           Or InStr(u, "format=webp") > 0 Then
            LooksLikeRealImage = True
        Else
            LooksLikeRealImage = False
        End If
    End If
End Function

Private Function GuessImageExt(url As String) As String
    Dim u As String: u = LCase(url)
    Dim q As Long: q = InStr(u, "?")
    If q > 0 Then u = Left(u, q - 1)
    Dim hash As Long: hash = InStr(u, "#")
    If hash > 0 Then u = Left(u, hash - 1)
    Dim dot As Long: dot = InStrRev(u, ".")
    If dot > 0 And Len(u) - dot <= 5 Then
        Dim ext As String: ext = Mid(u, dot + 1)
        Select Case ext
            Case "jpg", "jpeg", "png", "gif", "webp", "bmp", "tif", "tiff", "svg"
                GuessImageExt = ext
                Exit Function
        End Select
    End If
    GuessImageExt = "jpg"
End Function

' ---------- Filesystem helpers ----------------------------------------------

Private Function MakeAssetFolder(slug As String) As String
    Dim base As String: base = AssetsRoot()
    EnsureFolder base
    Dim ts As String: ts = Format(Now, "yyyymmdd_hhnnss")
    Dim folder As String: folder = base & "\" & slug & "_" & ts
    EnsureFolder folder
    MakeAssetFolder = folder
End Function

Private Function AssetsRoot() As String
    Dim deck As String: deck = ActivePresentation.FullName
    Dim deckDir As String: deckDir = ParentFolder(deck)
    AssetsRoot = deckDir & "\assets"
End Function

Private Sub EnsureFolder(path As String)
    On Error Resume Next
    If Len(path) = 0 Then Exit Sub
    If Dir(path, vbDirectory) = "" Then
        ' MkDir cannot create nested folders in one call - walk parents.
        Dim parts() As String
        parts = Split(Replace(path, "/", "\"), "\")
        Dim acc As String: acc = parts(0)
        Dim i As Long
        For i = 1 To UBound(parts)
            acc = acc & "\" & parts(i)
            If Dir(acc, vbDirectory) = "" Then MkDir acc
        Next i
    End If
    On Error GoTo 0
End Sub

Private Function ParentFolder(path As String) As String
    Dim p As String: p = Replace(path, "/", "\")
    Dim slash As Long: slash = InStrRev(p, "\")
    If slash > 0 Then ParentFolder = Left(p, slash - 1) Else ParentFolder = ""
End Function

Public Sub WriteTextFile(path As String, contents As String)
    ' Write UTF-8 *without* the leading BOM. ADODB.Stream's text mode prepends
    ' a 3-byte BOM that breaks downstream JSON parsers (Python's json.loads,
    ' VBA modJSON.ParseJson). Strip it by re-staging through a binary stream.
    Dim text As Object: Set text = CreateObject("ADODB.Stream")
    text.Type = 2
    text.Charset = "utf-8"
    text.Open
    text.WriteText contents

    Dim bin As Object: Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1     ' binary
    bin.Open

    text.Position = 3       ' skip BOM
    text.CopyTo bin
    text.Close

    On Error Resume Next
    If Dir(path) <> "" Then Kill path
    On Error GoTo 0
    bin.SaveToFile path, 2
    bin.Close
End Sub

Public Function ReadTextFile(path As String) As String
    On Error Resume Next
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    ReadTextFile = stream.ReadText
    stream.Close
    On Error GoTo 0
End Function

Private Function SafeSlug(s As String) As String
    Dim out As String: out = ""
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch Like "[A-Za-z0-9]" Then
            out = out & ch
        ElseIf ch = "." Or ch = "-" Or ch = "_" Then
            out = out & ch
        End If
    Next i
    If Len(out) > 40 Then out = Left(out, 40)
    If Len(out) = 0 Then out = "page"
    SafeSlug = out
End Function
