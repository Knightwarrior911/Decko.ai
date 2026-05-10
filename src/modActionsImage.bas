Attribute VB_Name = "modActionsImage"
Option Explicit

Public Sub Do_insert_picture(slideNum As Long, picturePath As String, _
                             leftPt As Single, topPt As Single, _
                             widthPt As Single, heightPt As Single)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 6001, "Do_insert_picture", "slide_out_of_range"
    End If
    If Not FileExists(picturePath) Then
        Err.Raise vbObjectError + 6002, "Do_insert_picture", "file_not_found: " & picturePath
    End If
    Dim pic As Shape
    Set pic = pres.Slides(slideNum).Shapes.AddPicture( _
        FileName:=picturePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=leftPt, Top:=topPt, _
        Width:=widthPt, Height:=heightPt)
    ' Force exact dimensions — AddPicture preserves source aspect by default
    pic.LockAspectRatio = msoFalse
    pic.Width = widthPt
    pic.Height = heightPt
End Sub

Public Sub Do_replace_picture(slideNum As Long, shapeId As Long, picturePath As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 6001, "Do_replace_picture", "shape not found"
    If sh.Type <> msoPicture Then Err.Raise vbObjectError + 6003, "Do_replace_picture", "shape is not a picture"
    If Not FileExists(picturePath) Then Err.Raise vbObjectError + 6002, "Do_replace_picture", "file_not_found: " & picturePath

    Dim L As Single, T As Single, W As Single, H As Single
    L = sh.Left: T = sh.Top: W = sh.Width: H = sh.Height
    sh.Delete
    Dim pic As Shape
    Set pic = ActivePresentation.Slides(slideNum).Shapes.AddPicture( _
        FileName:=picturePath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=L, Top:=T, Width:=W, Height:=H)
    pic.LockAspectRatio = msoFalse
    pic.Width = W
    pic.Height = H
End Sub

Private Function FileExists(p As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(p)
End Function

' Drop every image in `folder` onto picker slide(s) as labeled grids.
' User uses these slides to visually pick which images to keep, then
' references the filenames (e.g. "img_003.jpg") in a follow-up
' build_image_grid_table. Spans multiple slides if N > maxPerSlide so each
' thumbnail stays large enough to recognize.
Public Sub Do_build_image_picker_slide(folder As String, _
                                        Optional ByVal cols As Long = 4, _
                                        Optional ByVal insertAt As Long = 0, _
                                        Optional ByVal maxPerSlide As Long = 24)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then
        Err.Raise vbObjectError + 6010, "Do_build_image_picker_slide", "folder not found: " & folder
    End If
    If cols < 1 Then cols = 4
    If maxPerSlide < 1 Then maxPerSlide = 24

    ' Use a Collection of paths instead of a passed array - cleaner across subs.
    Dim names As New Collection
    Dim f As Object
    For Each f In fso.GetFolder(folder).Files
        Dim ext As String: ext = LCase(fso.GetExtensionName(f.Name))
        Select Case ext
            Case "jpg", "jpeg", "png", "gif", "bmp", "webp", "tif", "tiff"
                names.Add f.Path
        End Select
    Next f
    Dim n As Long: n = names.Count
    If n = 0 Then
        Err.Raise vbObjectError + 6011, "Do_build_image_picker_slide", "no images in folder"
    End If
    SortCollection names

    Dim pres As Presentation: Set pres = ActivePresentation
    Dim layout As CustomLayout: Set layout = pres.SlideMaster.CustomLayouts(1)
    Dim slidesNeeded As Long: slidesNeeded = (n + maxPerSlide - 1) \ maxPerSlide

    Dim startPos As Long: startPos = insertAt
    If startPos < 1 Or startPos > pres.Slides.Count + 1 Then startPos = pres.Slides.Count + 1

    Dim slW As Single: slW = pres.PageSetup.SlideWidth
    Dim slH As Single: slH = pres.PageSetup.SlideHeight
    Dim margin As Single: margin = 20
    Dim gap As Single: gap = 8
    Dim labelH As Single: labelH = 14

    Dim slIdx As Long
    For slIdx = 1 To slidesNeeded
        Dim batchStart As Long: batchStart = (slIdx - 1) * maxPerSlide + 1
        Dim batchEnd As Long: batchEnd = slIdx * maxPerSlide
        If batchEnd > n Then batchEnd = n

        Dim sl As Slide: Set sl = pres.Slides.AddSlide(startPos + slIdx - 1, layout)

        Dim hdr As Shape
        Set hdr = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            20, 10, slW - 40, 28)
        hdr.TextFrame.TextRange.Text = "PICKER " & slIdx & "/" & slidesNeeded & _
            ": images " & batchStart & "-" & batchEnd & " from " & folder
        hdr.TextFrame.TextRange.Font.Size = 11
        hdr.TextFrame.TextRange.Font.Bold = msoTrue

        Dim count As Long: count = batchEnd - batchStart + 1
        Dim rows As Long: rows = (count + cols - 1) \ cols
        Dim availW As Single: availW = slW - 2 * margin
        Dim availH As Single: availH = slH - 50 - margin
        Dim cellW As Single: cellW = (availW - gap * (cols - 1)) / cols
        Dim cellH As Single: cellH = (availH - gap * (rows - 1)) / rows
        Dim picH As Single: picH = cellH - labelH - 2

        Dim i As Long
        For i = batchStart To batchEnd
            Dim idxInBatch As Long: idxInBatch = i - batchStart
            Dim r As Long: r = idxInBatch \ cols
            Dim c As Long: c = idxInBatch Mod cols
            Dim x As Single: x = margin + c * (cellW + gap)
            Dim y As Single: y = 50 + r * (cellH + gap)
            Dim picPath As String: picPath = CStr(names(i))

            On Error Resume Next
            AddPictureContain sl, picPath, x, y, cellW, picH
            Err.Clear
            On Error GoTo 0

            On Error Resume Next
            Dim lbl As Shape
            Set lbl = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                x, y + picH, cellW, labelH)
            If Err.Number = 0 And Not lbl Is Nothing Then
                lbl.TextFrame.TextRange.Text = fso.GetFileName(picPath)
                lbl.TextFrame.TextRange.Font.Size = 7
                lbl.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
                lbl.TextFrame.MarginLeft = 1
                lbl.TextFrame.MarginRight = 1
                lbl.TextFrame.MarginTop = 0
                lbl.TextFrame.MarginBottom = 0
            End If
            Err.Clear
            On Error GoTo 0
        Next i
    Next slIdx
End Sub

Private Sub SortCollection(col As Collection)
    Dim n As Long: n = col.Count
    If n < 2 Then Exit Sub
    Dim arr() As String: ReDim arr(1 To n)
    Dim i As Long
    For i = 1 To n: arr(i) = CStr(col(i)): Next i
    SortStrings arr
    ' Rebuild collection in sorted order.
    Do While col.Count > 0: col.Remove 1: Loop
    For i = 1 To n: col.Add arr(i): Next i
End Sub

' Place picture inside (left, top, w, h), preserving source aspect ratio.
' Picture is centered; uncovered area inside the rect stays transparent.
Public Sub AddPictureContain(sl As Slide, imgPath As String, _
                              leftPt As Single, topPt As Single, _
                              widthPt As Single, heightPt As Single)
    On Error GoTo failed
    Dim pic As Shape
    ' Pass -1 for Width/Height so PowerPoint uses the source pixel dimensions
    ' (in points). This gives us the natural aspect ratio to compute the fit.
    Set pic = sl.Shapes.AddPicture(FileName:=imgPath, _
        LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=leftPt, Top:=topPt, Width:=-1, Height:=-1)
    If pic Is Nothing Then GoTo failed

    Dim natW As Single: natW = pic.Width
    Dim natH As Single: natH = pic.Height
    If natW <= 0 Or natH <= 0 Then GoTo failed

    Dim scaleW As Single: scaleW = widthPt / natW
    Dim scaleH As Single: scaleH = heightPt / natH
    Dim factor As Single
    If scaleW < scaleH Then factor = scaleW Else factor = scaleH

    pic.LockAspectRatio = msoTrue
    pic.Width = natW * factor
    pic.Left = leftPt + (widthPt - pic.Width) / 2
    pic.Top = topPt + (heightPt - pic.Height) / 2
    Exit Sub
failed:
    Err.Clear
End Sub

Private Sub SortStrings(arr() As String)
    Dim lo As Long: lo = LBound(arr)
    Dim hi As Long: hi = UBound(arr)
    Dim i As Long, j As Long, tmp As String
    For i = lo To hi - 1
        For j = i + 1 To hi
            If LCase(arr(i)) > LCase(arr(j)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub
