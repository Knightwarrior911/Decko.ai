VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExecute 
   Caption         =   "Decko.ai � Execute Instructions"
   ClientHeight    =   10400
   ClientLeft      =   90
   ClientTop       =   410
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
' When the user loads actions from a file, the full uncorrupted JSON is held
' here and takes precedence over the textbox. MSForms textboxes can mangle
' large pastes (whitespace injected into numbers/keys); the file path avoids
' the textbox entirely. Cleared the moment the user types in the textbox.
Private mLoadedJson As String
Private Sub btnFixThis_Click()
    ' Post-Apply quality warnings -> clipboard. See modVerify.
    Dim n As Long
    n = modVerify.CopyWarningsPromptToClipboard()
    If n = 0 Then
        lblStatus.Caption = "No warnings to copy. Either run Apply first, or the deck is clean."
    Else
        lblStatus.Caption = n & " warning(s) copied to clipboard as LLM prompt. " & _
                            "Paste into your chat and ask the model to fix."
    End If
End Sub

Private Sub btnFixErrors_Click()
    ' Pre-Apply action validation errors -> clipboard. Reads current JSON
    ' (textbox or loaded file), runs PreviewValidate on each action, builds
    ' an LLM-ready prompt with errors + canonical guidance for each failing
    ' action type. Eliminates the user having to read INVALID lines and
    ' hand-type a correction request.
    Dim json As String: json = CurrentJson()
    If Len(json) = 0 Then
        lblStatus.Caption = "No actions JSON to validate. Paste or Load a batch first."
        Exit Sub
    End If
    Dim prompt As String
    prompt = modExecuteInstructions.BuildErrorFixPrompt(json)
    If Len(prompt) = 0 Then
        ' No validation errors: the actions are well-formed but may still
        ' not match what the user meant. Offer an intent re-steer prompt
        ' (the plain-language plan + a revise template) instead of a
        ' dead-end message. Same clipboard pattern as the error path.
        Dim plan As String
        plan = modExecuteInstructions.BuildActionPlanSummary(json)
        Dim rprompt As String
        rprompt = "These PowerPoint actions are valid but may not match my intent." & vbCrLf & _
                  "Here is exactly what they will do:" & vbCrLf & vbCrLf & _
                  plan & vbCrLf & vbCrLf & _
                  "That is NOT what I meant. Revise the actions JSON so it does this instead: " & _
                  "<describe the correct intent here>. " & _
                  "Return only the corrected {""actions"":[...]} JSON, no prose."
        Dim dobjR As MSForms.DataObject
        Set dobjR = New MSForms.DataObject
        dobjR.SetText rprompt
        dobjR.PutInClipboard
        lblStatus.Caption = "Re-steer prompt copied. Edit the <...> part, paste into your LLM chat."
        Exit Sub
    End If
    Dim dobj As MSForms.DataObject
    Set dobj = New MSForms.DataObject
    dobj.SetText prompt
    dobj.PutInClipboard
    lblStatus.Caption = "Error-fix prompt copied to clipboard. Paste into your LLM chat."
End Sub

Private Sub UserForm_Initialize()
    lblStatus.Caption = ""
    mLoadedJson = ""
End Sub

' The JSON to parse/apply: a file loaded via btnLoadFile if one is held,
' otherwise the textbox contents.
Private Function CurrentJson() As String
    If Len(mLoadedJson) > 0 Then
        CurrentJson = mLoadedJson
    Else
        CurrentJson = txtInstructions.text
    End If
End Function

Private Sub txtInstructions_Change()
    ' Hand-editing the textbox supersedes any previously loaded file.
    mLoadedJson = ""
End Sub

Private Sub btnLoadFile_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select an actions JSON file"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "JSON / JSONL / text", "*.json;*.jsonl;*.txt"
    fd.Filters.Add "All files", "*.*"
    If fd.Show <> -1 Then Exit Sub

    Dim path As String: path = fd.SelectedItems(1)
    Dim s As String
    On Error GoTo IOFail
    Dim fnum As Integer: fnum = FreeFile
    Open path For Input As #fnum
    If LOF(fnum) > 0 Then s = Input$(LOF(fnum), fnum)
    Close #fnum
    On Error GoTo 0

    txtInstructions.text = s          ' raises txtInstructions_Change -> clears mLoadedJson
    mLoadedJson = s                   ' ...so set it AFTER the assignment above
    lstActions.Clear
    Set mParsed = Nothing
    lblStatus.Caption = "Loaded " & Len(s) & " chars from " & path & " -- click Parse."
    Exit Sub
IOFail:
    On Error GoTo 0
    lblStatus.Caption = "Could not read file: " & Err.Description
End Sub

Private Sub btnParse_Click()
    lblStatus.Caption = ""
    lstActions.Clear

    On Error Resume Next
    Set mParsed = modJSON.ParseJson(modExecuteInstructions.SanitizeJsonInput(CurrentJson()))
    If Err.Number <> 0 Then
        lblStatus.Caption = "Invalid JSON: " & Err.Description & _
            "  (large textbox pastes can corrupt -- use 'Load from file' for big batches)"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    If Not mParsed.Exists("actions") Then
        lblStatus.Caption = "Missing top-level 'actions' array."
        Exit Sub
    End If

    Dim actions As Object: Set actions = mParsed("actions")
    ReDim mValid(1 To actions.count)

    Dim i As Long, anyValid As Boolean: anyValid = False
    For i = 1 To actions.count
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

    ' Plain-language plan preview (read-only; no mutation). The executor has
    ' no undo/auto-backup, so this is the user's chance to catch intent
    ' drift BEFORE Apply. Appended below the validation rows (non-breaking).
    Dim planTxt As String
    planTxt = modExecuteInstructions.BuildActionPlanSummary(CurrentJson())
    lstActions.AddItem "----  WILL DO (preview)  ----"
    Dim plLines() As String, pli As Long
    plLines = Split(planTxt, vbCrLf)
    For pli = LBound(plLines) To UBound(plLines)
        lstActions.AddItem plLines(pli)
    Next pli

    lblStatus.Caption = actions.count & " actions parsed. Review the WILL DO preview below. " & _
                        IIf(anyValid, "Click Apply to run, or 'Fix Errors' to copy a re-steer prompt.", _
                                      "No valid actions; nothing to apply.")
End Sub

Private Sub btnApply_Click()
    If mParsed Is Nothing Then
        lblStatus.Caption = "Click Parse first."
        Exit Sub
    End If
    Dim summary As String
    summary = modExecuteInstructions.ExecuteFromString(CurrentJson())
    lblStatus.Caption = summary
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Function GetStrSafe(d As Object, key As String) As String
    If d.Exists(key) Then GetStrSafe = CStr(d(key))
End Function


