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
