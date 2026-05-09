"""Rebuild frmImportSlides from scratch via Add(VB_FORM) with correct
InsideWidth/Height. Bypasses name-reservation by temporarily removing
modUI's reference to frmImportSlides during the Add+rename step.

This avoids the auto-shrunk InsideWidth=214 problem caused by importing
the stale .frx blob.
"""
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
SRC = REPO_ROOT / "src"

VB_FORM = 3
TEXTBOX = "Forms.TextBox.1"
CMDBTN  = "Forms.CommandButton.1"
LABEL   = "Forms.Label.1"

DARK_BG  = 0                                  # #000000
DARK_BTN = 51 + 51 * 256 + 51 * 65536          # #333333
MID_BTN  = 87 + 87 * 256 + 87 * 65536          # #575757
WHITE    = 16777215                            # #FFFFFF
BLACK    = 0
FONT_NAME = "Cascadia Code"


def setp(c, **kw):
    for k, v in kw.items():
        try:
            setattr(c, k, v)
        except Exception as e:
            print(f"    [warn] {k}={v!r}: {e}")


def style_input(c):
    c.BackColor = WHITE
    c.ForeColor = BLACK
    c.SpecialEffect = 2
    c.Font.Name = FONT_NAME
    c.Font.Size = 10


def style_label(c):
    c.BackColor = DARK_BG
    c.ForeColor = WHITE
    c.Font.Name = FONT_NAME
    c.Font.Size = 10
    try:
        c.AutoSize = False
        c.WordWrap = False
    except Exception:
        pass


def style_btn_primary(c):
    c.BackColor = DARK_BTN
    c.ForeColor = WHITE
    c.Font.Name = FONT_NAME
    c.Font.Size = 10
    c.Font.Bold = True


def style_btn_secondary(c):
    c.BackColor = MID_BTN
    c.ForeColor = WHITE
    c.Font.Name = FONT_NAME
    c.Font.Size = 10
    c.Font.Bold = True


FRM_CODE = '''\
Option Explicit

Private Sub UserForm_Initialize()
    txtPath.Text = ""
    txtRange.Text = ""
    txtPosition.Text = "1"
    btnImport.Enabled = False
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
    btnImport.Enabled = (Len(txtPath.Text) > 0 And Len(txtRange.Text) > 0 And Len(txtPosition.Text) > 0)
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
'''

MOD_UI_NO_REF = '''\
Attribute VB_Name = "modUI"
Option Explicit

' Public entry points registered as macros (visible in Alt+F8).
Public Sub ExportSnapshot()
    frmExport.Show vbModeless
End Sub

Public Sub ExecuteInstructions()
    frmExecute.Show vbModeless
End Sub
'''

MOD_UI_FULL = '''\
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
'''


def remove_by_name(components, name):
    for c in components:
        try:
            if c.Name == name:
                components.Remove(c)
                return True
        except Exception:
            pass
    return False


def main():
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.Visible = True

    # Stage 1: build the form in a TEMP .pptm where no name conflict exists,
    # then export the .frm/.frx files. Stage 2: Import them into the carrier.
    TEMP = REPO_ROOT / "tools" / "_temp_form_stage.pptm"
    if TEMP.exists():
        TEMP.unlink()

    print(f"[stage 1] Build form in temp project: {TEMP}")
    pres = app.Presentations.Add(WithWindow=True)
    time.sleep(2)
    try:
        # Save as macro-enabled so VBProject persists
        pres.SaveAs(str(TEMP), 25)  # ppSaveAsOpenXMLPresentationMacroEnabled
        time.sleep(1)
        components = pres.VBProject.VBComponents
        comp = components.Add(VB_FORM)
        comp.Name = "frmImportSlides"
        designer = comp.Designer
        designer.Caption = "Decko.ai • Import Slides"
        designer.BackColor = DARK_BG
        designer.Font.Name = FONT_NAME
        designer.Font.Size = 10

        # Resize the form. Phase 2 used designer.Width/Height; some versions
        # accept it, others reject. Try every approach and keep what works.
        for accessor, val in [
            ("designer.Width = 480",     lambda: setattr(designer, "Width", 480)),
            ("designer.Height = 320",    lambda: setattr(designer, "Height", 320)),
            ("comp.Properties Width",    lambda: comp.Properties("Width").__setattr__("Value", 480)),
            ("comp.Properties Height",   lambda: comp.Properties("Height").__setattr__("Value", 320)),
        ]:
            try:
                val()
                print(f"  ok: {accessor}")
            except Exception as e:
                print(f"  fail: {accessor}: {e}")
        try:
            print(f"  inside now: {designer.InsideWidth} x {designer.InsideHeight}")
        except Exception:
            pass

        ctrls = designer.Controls

        c = ctrls.Add(LABEL, "lblPath", True);     setp(c, Caption="Source deck",
            Top=12, Left=12, Width=100, Height=20); style_label(c)
        c = ctrls.Add(TEXTBOX, "txtPath", True);   setp(c, Top=12, Left=120, Width=280, Height=22, Locked=True); style_input(c)
        c = ctrls.Add(CMDBTN, "btnBrowse", True);  setp(c, Caption="Browse...",
            Top=12, Left=408, Width=60, Height=22); style_btn_secondary(c)

        c = ctrls.Add(LABEL, "lblRange", True);    setp(c, Caption="Slide range (e.g. 1-3,5,7-9)",
            Top=50, Left=12, Width=240, Height=20); style_label(c)
        c = ctrls.Add(TEXTBOX, "txtRange", True);  setp(c, Top=50, Left=256, Width=212, Height=22); style_input(c)

        c = ctrls.Add(LABEL, "lblPosition", True); setp(c, Caption="Insert at position",
            Top=90, Left=12, Width=240, Height=20); style_label(c)
        c = ctrls.Add(TEXTBOX, "txtPosition", True); setp(c, Top=90, Left=256, Width=60, Height=22); style_input(c)

        c = ctrls.Add(CMDBTN, "btnImport", True);  setp(c, Caption="Import", Enabled=False,
            Top=140, Left=12, Width=80, Height=28); style_btn_primary(c)
        c = ctrls.Add(CMDBTN, "btnCancel", True);  setp(c, Caption="Cancel",
            Top=140, Left=100, Width=80, Height=28); style_btn_secondary(c)

        c = ctrls.Add(LABEL, "lblStatus", True);   setp(c, Caption="",
            Top=180, Left=12, Width=456, Height=110); style_label(c)

        # VBA code
        cm = comp.CodeModule
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(FRM_CODE)

        pres.Save()
        # Export to src/ for next update_macros sync
        comp.Export(str(SRC / "frmImportSlides.frm"))
        print(f"  exported -> {SRC / 'frmImportSlides.frm'}")
        pres.Close()
        time.sleep(1)
        # Wipe temp
        if TEMP.exists():
            TEMP.unlink()

        # Stage 2: Import the fresh .frm into the real carrier
        print(f"[stage 2] Import fresh .frm into carrier")
        pres = app.Presentations.Open(str(CARRIER), WithWindow=True)
        time.sleep(2)
        components = pres.VBProject.VBComponents
        if remove_by_name(components, "frmImportSlides"):
            print("  removed old frmImportSlides")
        components.Import(str(SRC / "frmImportSlides.frm"))
        print("  imported frmImportSlides")
        pres.Save()
        print("[done]")
    finally:
        try:
            pres.Close()
        except Exception:
            pass
        app.Quit()
        time.sleep(1)


if __name__ == "__main__":
    sys.exit(main())
