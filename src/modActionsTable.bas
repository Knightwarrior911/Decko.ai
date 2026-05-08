Attribute VB_Name = "modActionsTable"
Option Explicit

Public Sub Do_add_table_row(slideNum As Long, shapeId As Long, afterRow As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_add_table_row", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_add_table_row", "shape is not a table"
    Dim tbl As Object: Set tbl = sh.Table
    Dim curRows As Long: curRows = tbl.Rows.Count
    If afterRow < 0 Or afterRow > curRows Then
        Err.Raise vbObjectError + 8003, "Do_add_table_row", "after_row out of range"
    End If
    If afterRow >= curRows Then
        ' Append at end — call without arg
        tbl.Rows.Add
    Else
        Dim insertAt As Long
        insertAt = afterRow + 1
        tbl.Rows.Add insertAt
    End If
End Sub

Public Sub Do_delete_table_row(slideNum As Long, shapeId As Long, rowNum As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_delete_table_row", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_delete_table_row", "shape is not a table"
    Dim tbl As Object: Set tbl = sh.Table
    If rowNum < 1 Or rowNum > tbl.Rows.Count Then
        Err.Raise vbObjectError + 8004, "Do_delete_table_row", "row out of range"
    End If
    tbl.Rows(rowNum).Delete
End Sub

Public Sub Do_add_table_col(slideNum As Long, shapeId As Long, afterCol As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_add_table_col", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_add_table_col", "shape is not a table"
    Dim tbl As Object: Set tbl = sh.Table
    Dim curCols As Long: curCols = tbl.Columns.Count
    If afterCol < 0 Or afterCol > curCols Then
        Err.Raise vbObjectError + 8005, "Do_add_table_col", "after_col out of range"
    End If
    If afterCol >= curCols Then
        tbl.Columns.Add
    Else
        Dim insertAt As Long
        insertAt = afterCol + 1
        tbl.Columns.Add insertAt
    End If
End Sub

Public Sub Do_delete_table_col(slideNum As Long, shapeId As Long, colNum As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_delete_table_col", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_delete_table_col", "shape is not a table"
    Dim tbl As Object: Set tbl = sh.Table
    If colNum < 1 Or colNum > tbl.Columns.Count Then
        Err.Raise vbObjectError + 8006, "Do_delete_table_col", "col out of range"
    End If
    tbl.Columns(colNum).Delete
End Sub

Public Sub Do_merge_cells(slideNum As Long, shapeId As Long, _
                          rowA As Long, colA As Long, _
                          rowB As Long, colB As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_merge_cells", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_merge_cells", "shape is not a table"
    Dim tbl As Object: Set tbl = sh.Table
    tbl.Cell(rowA, colA).Merge tbl.Cell(rowB, colB)
End Sub
