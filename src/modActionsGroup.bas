Attribute VB_Name = "modActionsGroup"
Option Explicit

Public Sub Do_group_shapes(slideNum As Long, shapeIds As Variant)
    Dim ids() As Long
    Dim cnt As Long: cnt = modActionsLayout.NormalizeIdsArray(shapeIds, ids)
    If cnt < 2 Then Err.Raise vbObjectError + 9001, "Do_group_shapes", "need >=2 shapes"
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim sl As Slide: Set sl = pres.Slides(slideNum)

    Dim names() As String
    ReDim names(0 To cnt - 1)
    Dim i As Long
    For i = 0 To cnt - 1
        Dim sh As Shape: Set sh = modActions.FindShape(slideNum, ids(i))
        If sh Is Nothing Then Err.Raise vbObjectError + 9002, "Do_group_shapes", "shape not found: " & ids(i)
        names(i) = sh.Name
    Next i
    Dim shapesObj As Object: Set shapesObj = sl.Shapes
    Dim rng As Object: Set rng = shapesObj.Range(names)
    rng.Group
End Sub

Public Sub Do_ungroup(slideNum As Long, shapeId As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 9003, "Do_ungroup", "shape not found"
    If sh.Type <> msoGroup Then Err.Raise vbObjectError + 9004, "Do_ungroup", "shape is not a group"
    sh.Ungroup
End Sub
