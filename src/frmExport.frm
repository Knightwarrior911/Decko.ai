VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExport 
   Caption         =   "PPT AI Editor — Export Snapshot"
   ClientHeight    =   7200
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   10800
   OleObjectBlob   =   "frmExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PROMPT_TEMPLATE As String = _
"You are editing a PowerPoint presentation. Below is the current state as JSON:" & vbCrLf & vbCrLf & _
"```json" & vbCrLf & "{snapshot}" & vbCrLf & "```" & vbCrLf & vbCrLf & _
"I want the following changes:" & vbCrLf & vbCrLf & _
"[REPLACE THIS LINE WITH YOUR REQUEST]" & vbCrLf & vbCrLf & _
"Return ONLY a valid instructions JSON in this exact format. No prose, no" & vbCrLf & _
"explanation, no markdown fences:" & vbCrLf & vbCrLf & _
"{" & vbCrLf & _
"  ""actions"": [" & vbCrLf & _
"    {""type"": ""<action_type>"", ""slide"": <int>, ""shape_id"": <int>, ...}" & vbCrLf & _
"  ]" & vbCrLf & _
"}" & vbCrLf & vbCrLf & _
"Rules:" & vbCrLf & _
"- Use only shape_ids that exist in the snapshot. Do not invent ids." & vbCrLf & _
"- Slide numbers are 1-based." & vbCrLf & _
"- Colors as #RRGGBB hex." & vbCrLf & _
"- Lengths in points." & vbCrLf & _
"- Allowed action types: set_text, set_font_size, set_font_bold," & vbCrLf & _
"  set_font_italic, set_font_color, set_fill_color, move_shape," & vbCrLf & _
"  resize_shape, delete_shape, add_slide, delete_slide, duplicate_slide," & vbCrLf & _
"  set_cell_text, swap_table_columns, swap_table_rows."

Private Sub UserForm_Initialize()
    On Error Resume Next
    txtSnapshot.Text = modExportSnapshot.BuildSnapshotJson()
    If Err.Number <> 0 Then
        txtSnapshot.Text = "ERROR: " & Err.Description
        Err.Clear
    End If
End Sub

Private Sub btnCopySnapshot_Click()
    CopyToClipboard txtSnapshot.Text
    lblStatus.Caption = "Snapshot copied to clipboard."
End Sub

Private Sub btnCopyWithTemplate_Click()
    Dim payload As String
    payload = Replace(PROMPT_TEMPLATE, "{snapshot}", txtSnapshot.Text)
    CopyToClipboard payload
    lblStatus.Caption = "Snapshot + prompt template copied to clipboard."
End Sub

Private Sub btnSaveTxt_Click()
    Dim deckPath As String: deckPath = ActivePresentation.FullName
    Dim ts As String: ts = Format(Now, "yyyy-mm-dd_hhnnss")
    Dim outPath As String
    outPath = deckPath & "_snapshot_" & ts & ".txt"

    Dim f As Integer: f = FreeFile
    Open outPath For Output As #f
    Print #f, txtSnapshot.Text
    Close #f
    lblStatus.Caption = "Saved: " & outPath
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub CopyToClipboard(s As String)
    Dim doObj As Object
    Set doObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    doObj.SetText s
    doObj.PutInClipboard
End Sub


