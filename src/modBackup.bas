Attribute VB_Name = "modBackup"
Option Explicit

' Copy ActivePresentation file to a sibling backup with timestamp.
' Returns the backup file path. Raises if presentation is unsaved.
Public Function BackupActiveDeck() As String
    Dim pres As Presentation
    Set pres = ActivePresentation

    If Len(pres.Path) = 0 Then
        Err.Raise vbObjectError + 1001, _
                  "modBackup", _
                  "Cannot back up unsaved presentation. Save the deck first."
    End If

    Dim base As String, ext As String, dirPath As String
    dirPath = pres.Path
    Dim full As String: full = pres.FullName
    Dim dotPos As Long
    dotPos = InStrRev(full, ".")
    base = Mid(full, Len(dirPath) + 2, dotPos - Len(dirPath) - 2)
    ext = Mid(full, dotPos + 1)

    Dim ts As String
    ts = Format(Now, "yyyy-mm-dd_hhnnss")

    Dim dest As String
    dest = dirPath & "\" & base & "_backup_" & ts & "." & ext

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile pres.FullName, dest, True

    BackupActiveDeck = dest
End Function

' Append one JSONL line to <deckPath>.action_log.jsonl
Public Sub LogAction(deckPath As String, op As String, slideNum As Variant, _
                     shapeId As Variant, paramsJson As String, status As String, _
                     reason As String)
    Dim logPath As String
    logPath = deckPath & ".action_log.jsonl"

    Dim line As String
    line = "{""ts"":""" & FormatIsoUtc(Now) & """" & _
           ",""op"":""" & op & """" & _
           ",""slide"":" & ToJsonNumOrNull(slideNum) & _
           ",""shape_id"":" & ToJsonNumOrNull(shapeId) & _
           ",""params"":" & paramsJson & _
           ",""status"":""" & status & """"
    If Len(reason) > 0 Then
        line = line & ",""reason"":""" & EscapeJsonString(reason) & """"
    End If
    line = line & "}"

    Dim fnum As Integer
    fnum = FreeFile
    Open logPath For Append As #fnum
    Print #fnum, line
    Close #fnum
End Sub

Private Function ToJsonNumOrNull(v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        ToJsonNumOrNull = "null"
    Else
        ToJsonNumOrNull = CStr(v)
    End If
End Function

Private Function FormatIsoUtc(dt As Date) As String
    ' Local time, ISO 8601-style. "Z" suffix removed (timestamps are local, not UTC).
    ' HH = 24-hour clock; lowercase hh would produce 12-hour values for afternoon times.
    FormatIsoUtc = Format(dt, "yyyy-mm-dd") & "T" & Format(dt, "HH:nn:ss")
End Function

Public Function EscapeJsonString(s As String) As String
    Dim r As String: r = s
    r = Replace(r, "\", "\\")
    r = Replace(r, """", "\""")
    r = Replace(r, vbCrLf, "\n")
    r = Replace(r, vbCr, "\n")
    r = Replace(r, vbLf, "\n")
    r = Replace(r, vbTab, "\t")
    EscapeJsonString = r
End Function
