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
    nm = InputBox("Template name (saved to your Decko library):", "Capture Template")
    If Len(Trim(nm)) = 0 Then Exit Sub
    Dim act As Object: Set act = CreateObject("Scripting.Dictionary")
    act.Add "type", "capture_template"
    act.Add "name", nm
    modActionsCapture.Do_capture_template_act act
    MsgBox "Captured '" & nm & "'. It appears in the next " & _
           "Copy snapshot + prompt template.", vbInformation
End Sub

' Alt+F8: list captured templates; type a name to delete it.
Public Sub ManageTemplates()
    Dim man As String
    man = modActionsCapture.BuildCapturedManifest( _
        modActionsCapture.DefaultRegistryPath())
    If Len(man) = 0 Then man = "(no captured templates yet)"
    Dim choice As String
    choice = InputBox(man & vbCrLf & vbCrLf & _
        "Type a template name to DELETE it, or Cancel.", "Manage Templates")
    If Len(Trim(choice)) = 0 Then Exit Sub
    Dim act As Object: Set act = CreateObject("Scripting.Dictionary")
    act.Add "type", "delete_template"
    act.Add "name", choice
    modActionsCapture.Do_delete_template_act act
    MsgBox "Deleted '" & choice & "'.", vbInformation
End Sub
