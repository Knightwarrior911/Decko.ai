Attribute VB_Name = "modUI"
Option Explicit

' Public entry points registered as macros (visible in Alt+F8).
Public Sub ExecuteInstructions()
    frmExecute.Show vbModeless
End Sub

Public Sub ImportSlides()
    frmImportSlides.Show vbModeless
End Sub

' Test hook: writes the LLM prompt (with icon allow-list appended) to outPath.
' Used by tests/run_smoke_icon_prompt.py.
Public Sub DumpPromptToFile(outPath As String)
    Dim s As String: s = frmExport.PromptTemplate()
    Dim f As Integer: f = FreeFile
    Open outPath For Output As #f
    Print #f, s
    Close #f
End Sub

' Alt+F8: capture the active slide as a reusable template (Deck DNA).
' Standard-module macro -- no .frx form edit.
Public Sub CaptureTemplate()
    Dim nm As String
    nm = InputBox( _
        "Name this captured template (saved to your Decko library)." & vbCrLf & _
        "The active slide's layout becomes a reusable stamp.", _
        "Decko - Capture Template")
    If Len(Trim(nm)) = 0 Then Exit Sub
    Dim act As Object: Set act = CreateObject("Scripting.Dictionary")
    act.Add "type", "capture_template"
    act.Add "name", Trim(nm)
    modActionsCapture.Do_capture_template_act act
    MsgBox "Captured '" & Trim(nm) & "'." & vbCrLf & _
           "It now appears in 'Copy snapshot + prompt template' so your " & _
           "LLM can use it.", vbInformation, "Decko"
End Sub

' Alt+F8: view captured templates (numbered, readable) and delete one.
Public Sub ManageTemplates()
    Dim lst As String
    lst = modActionsCapture.NumberedTemplateList( _
        modActionsCapture.DefaultRegistryPath())
    If Len(lst) = 0 Then
        MsgBox "No captured templates yet. Run the CaptureTemplate macro " & _
               "on a slide you like.", vbInformation, _
               "Decko - Manage Templates"
        Exit Sub
    End If
    Dim choice As String
    choice = InputBox( _
        "YOUR CAPTURED TEMPLATES:" & vbCrLf & vbCrLf & lst & vbCrLf & _
        "To DELETE one, type its NAME exactly and press OK." & vbCrLf & _
        "Leave blank or press Cancel to just view (deletes nothing).", _
        "Decko - Manage Templates")
    If Len(Trim(choice)) = 0 Then Exit Sub
    Dim act As Object: Set act = CreateObject("Scripting.Dictionary")
    act.Add "type", "delete_template"
    act.Add "name", Trim(choice)
    modActionsCapture.Do_delete_template_act act
    MsgBox "Deleted '" & Trim(choice) & "'. Slides already built from it " & _
           "are unchanged.", vbInformation, "Decko - Manage Templates"
End Sub
