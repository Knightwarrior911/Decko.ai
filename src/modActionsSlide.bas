Attribute VB_Name = "modActionsSlide"
Option Explicit

Public Sub Do_set_slide_background_color(slideNum As Long, hexColor As String)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 7050, "Do_set_slide_background_color", "slide_out_of_range"
    End If
    Dim sl As Slide: Set sl = pres.Slides(slideNum)
    sl.FollowMasterBackground = msoFalse
    sl.Background.Fill.Visible = msoTrue
    sl.Background.Fill.Solid
    sl.Background.Fill.ForeColor.RGB = modActions.HexToRgb(hexColor)
End Sub

Public Sub Do_insert_slide_number(slideNum As Long, _
                                   leftPt As Single, topPt As Single, _
                                   widthPt As Single, heightPt As Single, _
                                   Optional refName As String = "", _
                                   Optional fontColor As String = "", _
                                   Optional fontSize As Long = 0)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 7051, "Do_insert_slide_number", "slide_out_of_range"
    End If
    Dim sh As Shape
    Set sh = pres.Slides(slideNum).Shapes.AddTextbox( _
        msoTextOrientationHorizontal, leftPt, topPt, widthPt, heightPt)
    If Len(refName) > 0 Then sh.Name = refName
    sh.TextFrame.AutoSize = ppAutoSizeNone
    sh.TextFrame.TextRange.InsertSlideNumber
    With sh.TextFrame.TextRange.Font
        If Len(fontColor) > 0 Then .Color.RGB = modActions.HexToRgb(fontColor)
        If fontSize > 0 Then .Size = fontSize
    End With
End Sub

Public Sub Do_move_slide(fromIdx As Long, toIdx As Long)
    Dim pres As Presentation: Set pres = ActivePresentation
    If fromIdx < 1 Or fromIdx > pres.Slides.Count Then
        Err.Raise vbObjectError + 7001, "Do_move_slide", "from out of range"
    End If
    If toIdx < 1 Or toIdx > pres.Slides.Count Then
        Err.Raise vbObjectError + 7002, "Do_move_slide", "to out of range"
    End If
    pres.Slides(fromIdx).MoveTo toIdx
End Sub

Public Sub Do_extract_slides(slideIndices As Variant, outputPath As String)
    Dim ids() As Long
    Dim cnt As Long: cnt = modActionsLayout.NormalizeIdsArray(slideIndices, ids)
    If cnt < 1 Then Err.Raise vbObjectError + 7003, "Do_extract_slides", "no slides specified"

    Dim src As Presentation: Set src = ActivePresentation
    Dim total As Long: total = src.Slides.Count
    Dim i As Long
    For i = 0 To cnt - 1
        If ids(i) < 1 Or ids(i) > total Then
            Err.Raise vbObjectError + 7004, "Do_extract_slides", _
                "slide index out of range: " & ids(i)
        End If
    Next i

    ' Clipboard-FREE extraction. PowerPoint slide Copy/Paste in
    ' automation is unreliable (it needs an active window + clipboard
    ' ownership; outPres is windowless), failing with -2147188160
    ' "Clipboard is empty...". Instead: SaveCopyAs the source to a temp
    ' file (does NOT change the source's path/dirty state), then pull
    ' each requested slide by index via Slides.InsertFromFile — no
    ' clipboard, no window, deterministic.
    Dim tmpSrc As String
    tmpSrc = Environ$("TEMP") & "\decko_extract_" & _
             Format(Now, "yyyymmddhhnnss") & Int(Rnd * 100000) & ".pptx"
    src.SaveCopyAs tmpSrc

    Dim outPres As Presentation
    Set outPres = Application.Presentations.Add(WithWindow:=msoFalse)
    For i = 0 To cnt - 1
        outPres.Slides.InsertFromFile tmpSrc, outPres.Slides.Count, _
                                      ids(i), ids(i)
    Next i

    outPres.SaveAs outputPath
    outPres.Close

    On Error Resume Next
    Kill tmpSrc
    On Error GoTo 0
End Sub

Public Sub Do_import_slides_from_deck(sourcePath As String, slideIndices As Variant, _
                                      targetPosition As Long)
    Dim ids() As Long
    Dim cnt As Long: cnt = modActionsLayout.NormalizeIdsArray(slideIndices, ids)
    If cnt < 1 Then Err.Raise vbObjectError + 7005, "Do_import_slides_from_deck", "no slide indices"
    If Not FileExistsLocal(sourcePath) Then
        Err.Raise vbObjectError + 7006, "Do_import_slides_from_deck", "source_not_found: " & sourcePath
    End If

    Dim pres As Presentation: Set pres = ActivePresentation
    If targetPosition < 1 Then targetPosition = 1
    If targetPosition > pres.Slides.Count + 1 Then targetPosition = pres.Slides.Count + 1

    Dim i As Long
    Dim insertedSoFar As Long: insertedSoFar = 0
    i = 0
    Do While i <= cnt - 1
        Dim startIdx As Long: startIdx = ids(i)
        Dim endIdx As Long: endIdx = startIdx
        Do While i + 1 <= cnt - 1
            If ids(i + 1) = ids(i) + 1 Then
                i = i + 1
                endIdx = ids(i)
            Else
                Exit Do
            End If
        Loop
        pres.Slides.InsertFromFile sourcePath, _
                                   targetPosition - 1 + insertedSoFar, _
                                   startIdx, endIdx
        insertedSoFar = insertedSoFar + (endIdx - startIdx + 1)
        i = i + 1
    Loop
End Sub

Private Function FileExistsLocal(p As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExistsLocal = fso.FileExists(p)
End Function

' Hide a slide from the slideshow (still visible in editor, skipped during play).
' Useful for backup-detail slides or speaker-only references.
Public Sub Do_set_slide_hidden(slideNum As Long, value As Boolean)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 7060, "Do_set_slide_hidden", "slide_out_of_range"
    End If
    pres.Slides(slideNum).SlideShowTransition.Hidden = IIf(value, msoTrue, msoFalse)
End Sub

' Clear speaker notes on a slide (set body placeholder text to empty).
Public Sub Do_clear_speaker_notes(slideNum As Long)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 7061, "Do_clear_speaker_notes", "slide_out_of_range"
    End If
    Dim sl As Slide: Set sl = pres.Slides(slideNum)
    Dim ph As Object
    Dim i As Long
    For i = 1 To sl.NotesPage.Shapes.Placeholders.Count
        Set ph = sl.NotesPage.Shapes.Placeholders(i)
        If ph.PlaceholderFormat.Type = ppPlaceholderBody Then
            ph.TextFrame.TextRange.Text = ""
            Exit Sub
        End If
    Next i
End Sub

' Rename a slide. PowerPoint exposes Slide.Name; visible in the slide-sorter
' tooltip and useful for the snapshot to label slides semantically.
Public Sub Do_set_slide_name(slideNum As Long, newName As String)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 7062, "Do_set_slide_name", "slide_out_of_range"
    End If
    If Len(Trim(newName)) = 0 Then Err.Raise vbObjectError + 7062, "Do_set_slide_name", "name empty"
    pres.Slides(slideNum).Name = newName
End Sub

' Slide transition effect (entry animation when this slide appears in slideshow).
' effect: "none"/"fade"/"push"/"wipe"/"split"/"reveal"/"cut"/"dissolve"/"checkerboard"/
'         "blinds"/"random_bars"/"box"/"comb"/"zoom"/"morph"
' speed:  "slow" / "medium" / "fast"  (optional, default medium)
' advance_on_click:        bool (default true)
' advance_after_seconds:   num >=0 (default 0 = no auto-advance)
Public Sub Do_set_slide_transition(slideNum As Long, effect As String, _
                                    speed As String, _
                                    advanceOnClick As Boolean, _
                                    advanceAfterSec As Double, _
                                    hasSpeed As Boolean, _
                                    hasOnClick As Boolean, _
                                    hasAdvSec As Boolean)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 7070, "Do_set_slide_transition", "slide_out_of_range"
    End If
    Dim t As Object: Set t = pres.Slides(slideNum).SlideShowTransition
    Dim ef As Long: ef = ResolveTransitionEffect(effect)
    On Error Resume Next
    t.EntryEffect = ef
    If hasSpeed Then
        Select Case LCase(speed)
            Case "slow":   t.Speed = 1   ' ppTransitionSpeedSlow
            Case "medium": t.Speed = 2   ' ppTransitionSpeedMedium
            Case "fast":   t.Speed = 3   ' ppTransitionSpeedFast
        End Select
    End If
    If hasOnClick Then t.AdvanceOnClick = IIf(advanceOnClick, msoTrue, msoFalse)
    If hasAdvSec Then
        If advanceAfterSec > 0 Then
            t.AdvanceOnTime = msoTrue
            t.AdvanceTime = advanceAfterSec
        Else
            t.AdvanceOnTime = msoFalse
        End If
    End If
    On Error GoTo 0
End Sub

Private Function ResolveTransitionEffect(effect As String) As Long
    Select Case LCase(Trim(effect))
        Case "none":            ResolveTransitionEffect = 0
        Case "cut":             ResolveTransitionEffect = 257
        Case "fade":            ResolveTransitionEffect = 1793
        Case "push":            ResolveTransitionEffect = 1796
        Case "wipe":            ResolveTransitionEffect = 1539
        Case "split":           ResolveTransitionEffect = 257
        Case "reveal":          ResolveTransitionEffect = 1795
        Case "dissolve":        ResolveTransitionEffect = 1281
        Case "checkerboard":    ResolveTransitionEffect = 769
        Case "blinds":          ResolveTransitionEffect = 513
        Case "random_bars":     ResolveTransitionEffect = 1025
        Case "box":             ResolveTransitionEffect = 3329
        Case "comb":            ResolveTransitionEffect = 3585
        Case "zoom":            ResolveTransitionEffect = 4097
        Case "morph":           ResolveTransitionEffect = 4353
        Case Else: Err.Raise vbObjectError + 7071, "ResolveTransitionEffect", "unknown effect: " & effect
    End Select
End Function

' Change slide layout. Single-slide shortcut for apply_layout_to_slides.
Public Sub Do_change_slide_layout(slideNum As Long, layoutIndex As Long)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 7072, "Do_change_slide_layout", "slide_out_of_range"
    End If
    If layoutIndex < 0 Or layoutIndex >= pres.SlideMaster.CustomLayouts.Count Then
        Err.Raise vbObjectError + 7072, "Do_change_slide_layout", _
                  "layout_index out of range (0.." & (pres.SlideMaster.CustomLayouts.Count - 1) & ")"
    End If
    pres.Slides(slideNum).CustomLayout = pres.SlideMaster.CustomLayouts(layoutIndex + 1)
End Sub

' --- Sections -------------------------------------------------------------
' PowerPoint deck section management. Sections group consecutive slides.

Public Sub Do_add_section(beforeSlide As Long, sectionName As String)
    Dim pres As Presentation: Set pres = ActivePresentation
    If beforeSlide < 1 Or beforeSlide > pres.Slides.Count Then
        Err.Raise vbObjectError + 7073, "Do_add_section", "slide_out_of_range"
    End If
    Dim secProps As Object: Set secProps = pres.SectionProperties
    secProps.AddSection beforeSlide, sectionName
End Sub

Public Sub Do_delete_section(sectionIndex As Long, deleteSlides As Boolean)
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim secProps As Object: Set secProps = pres.SectionProperties
    If sectionIndex < 1 Or sectionIndex > secProps.Count Then
        Err.Raise vbObjectError + 7074, "Do_delete_section", "section_index out of range"
    End If
    secProps.Delete sectionIndex, deleteSlides
End Sub

Public Sub Do_rename_section(sectionIndex As Long, newName As String)
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim secProps As Object: Set secProps = pres.SectionProperties
    If sectionIndex < 1 Or sectionIndex > secProps.Count Then
        Err.Raise vbObjectError + 7075, "Do_rename_section", "section_index out of range"
    End If
    secProps.Rename sectionIndex, newName
End Sub

Public Sub Do_move_section(sectionIndex As Long, toPosition As Long)
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim secProps As Object: Set secProps = pres.SectionProperties
    If sectionIndex < 1 Or sectionIndex > secProps.Count Then
        Err.Raise vbObjectError + 7076, "Do_move_section", "section_index out of range"
    End If
    secProps.Move sectionIndex, toPosition
End Sub
