VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImportSlides 
   Caption         =   "PPT AI Editor — Import Slides"
   ClientHeight    =   6400
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9600
   OleObjectBlob   =   "frmImportSlides.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImportSlides"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    txtPath.Text = ""
    txtRange.Text = ""
    txtPosition.Text = "1"
    btnImport.enabled = False
    lblStatus.Caption = ""
End Sub

Private Sub btnBrowse_Click()
    Dim picked As String
    On Error Resume Next
    Dim fd As Object
    Set fd = Application.FileDialog(3)
    If Not fd Is Nothing Then
        fd.Filters.Clear
        fd.Filters.Add "PowerPoint Files", "*.pptx; *.pptm"
        If fd.Show = -1 Then
            picked = fd.SelectedItems(1)
        End If
    End If
    On Error GoTo 0

    If Len(picked) = 0 Then
        picked = InputBox("Path to source deck:", "Source deck")
    End If

    If Len(picked) > 0 Then
        txtPath.Text = picked
        UpdateImportButton
    End If
End Sub

Private Sub txtRange_Change()
    UpdateImportButton
End Sub

Private Sub txtPosition_Change()
    UpdateImportButton
End Sub

Private Sub UpdateImportButton()
    btnImport.enabled = (Len(txtPath.Text) > 0 And Len(txtRange.Text) > 0 And Len(txtPosition.Text) > 0)
End Sub

Private Sub btnImport_Click()
    On Error GoTo Failure
    Dim ids As Variant
    ids = ParseRange(txtRange.Text)
    Dim pos As Long: pos = CLng(txtPosition.Text)

    Dim before As Long: before = ActivePresentation.Slides.Count
    modActionsSlide.Do_import_slides_from_deck txtPath.Text, ids, pos
    Dim afterCount As Long: afterCount = ActivePresentation.Slides.Count
    lblStatus.Caption = "Imported " & (afterCount - before) & " slide(s) at position " & pos
    Exit Sub
Failure:
    lblStatus.Caption = "ERROR: " & Err.Description
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Function ParseRange(s As String) As Variant
    Dim parts() As String
    parts = Split(s, ",")
    Dim col As New Collection
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim p As String: p = Trim(parts(i))
        If InStr(p, "-") > 0 Then
            Dim ab() As String: ab = Split(p, "-")
            Dim a As Long: a = CLng(Trim(ab(0)))
            Dim b As Long: b = CLng(Trim(ab(1)))
            Dim k As Long
            For k = a To b
                col.Add k
            Next k
        Else
            col.Add CLng(p)
        End If
    Next i
    Dim arr() As Long
    ReDim arr(0 To col.Count - 1)
    Dim j As Long
    For j = 1 To col.Count
        arr(j - 1) = col(j)
    Next j
    ParseRange = arr
End Function

