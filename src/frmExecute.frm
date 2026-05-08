VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExecute 
   Caption         =   "PPT AI Editor — Execute Instructions"
   ClientHeight    =   9600
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   12000
   OleObjectBlob   =   "frmExecute.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mParsed As Object
Private mValid() As Boolean

Private Sub btnParse_Click()
    lblStatus.Caption = ""
    btnApply.Enabled = False
    lstActions.Clear

    On Error Resume Next
    Set mParsed = modJSON.ParseJson(txtInstructions.Text)
    If Err.Number <> 0 Then
        lblStatus.Caption = "Invalid JSON: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    If Not mParsed.Exists("actions") Then
        lblStatus.Caption = "Missing top-level 'actions' array."
        Exit Sub
    End If

    Dim actions As Object: Set actions = mParsed("actions")
    ReDim mValid(1 To actions.Count)

    Dim i As Long, anyValid As Boolean: anyValid = False
    For i = 1 To actions.Count
        Dim act As Object: Set act = actions(i)
        Dim reason As String
        reason = modExecuteInstructions.PreviewValidate(act)
        mValid(i) = (Len(reason) = 0)
        If mValid(i) Then anyValid = True

        Dim row As String
        row = i & ". " & GetStrSafe(act, "type")
        If act.Exists("slide") Then row = row & " | slide=" & act("slide")
        If act.Exists("shape_id") Then row = row & " | shape_id=" & act("shape_id")
        If mValid(i) Then
            row = row & " | OK"
        Else
            row = row & " | INVALID: " & reason
        End If
        lstActions.AddItem row
    Next i

    btnApply.Enabled = anyValid
    lblStatus.Caption = actions.Count & " actions parsed. " & _
                        IIf(anyValid, "Click Apply to run valid actions.", _
                                      "No valid actions; nothing to apply.")
End Sub

Private Sub btnApply_Click()
    If mParsed Is Nothing Then Exit Sub
    Dim summary As String
    summary = modExecuteInstructions.ExecuteFromString(txtInstructions.Text)
    lblStatus.Caption = summary
    btnApply.Enabled = False
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Function GetStrSafe(d As Object, key As String) As String
    If d.Exists(key) Then GetStrSafe = CStr(d(key))
End Function


