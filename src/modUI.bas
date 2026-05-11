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
