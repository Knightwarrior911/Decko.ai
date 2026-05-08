Attribute VB_Name = "modUI"
Option Explicit

' Public entry points registered as macros (visible in Alt+F8).
Public Sub ExportSnapshot()
    frmExport.Show vbModeless
End Sub

Public Sub ExecuteInstructions()
    frmExecute.Show vbModeless
End Sub

Public Sub ImportSlides()
    frmImportSlides.Show vbModeless
End Sub
