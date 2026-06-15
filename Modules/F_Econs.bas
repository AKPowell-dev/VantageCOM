Attribute VB_Name = "F_Econs"
Option Explicit
Option Private Module

' Econs PowerPoint export.
'
' Entry point (mapped to the <cmd>econs command) shows a single unified
' configuration dialog, then runs the cross-product paste (each case rendered
' for each output) into the active PowerPoint presentation. The last-used
' configuration is remembered as a preset and only updated after a successful
' paste; cancelling leaves the saved preset unchanged.

Private Const LAYOUT_NAME As String = "content no text"
Private Const TARGET_SHEET As String = "Inputs"

' Standard centered-figure geometry (points): a freshly placed figure is scaled
' to FIG_WIDTH_PTS wide and dropped FIG_TOP_PTS below the slide top.
Private Const FIG_WIDTH_PTS As Single = 9.5 * 72
Private Const FIG_TOP_PTS As Single = 0.74 * 72

' Stamped onto every figure this tool pastes (via AlternativeText) so replace
' mode can find the figure it owns rather than guessing by size.
Private Const ECONS_TAG As String = "econs"

' History: last-N successful runs, stored as serialized cls_EconsConfig records
' in the per-user registry. Item1/Time1 is the most recent.
Private Const HIST_APP As String = "Vantage"
Private Const HIST_SECTION As String = "EconsHistory"
Private Const HIST_MAX As Long = 15

Public Sub Econs_Output_PPT_V2()
    Dim cfg As cls_EconsConfig
    Dim confirmed As Boolean

    Set cfg = New cls_EconsConfig
    cfg.LoadFromRegistry

    ' Open the dialog pre-populated with the last-used preset.
    Set UF_EconsConfig.Config = cfg
    UF_EconsConfig.PopulateFromConfig
    UF_EconsConfig.Show

    confirmed = UF_EconsConfig.Confirmed
    Set cfg = UF_EconsConfig.Config      ' mutated in place by the dialog
    Unload UF_EconsConfig

    ' Cancel: no paste, saved preset untouched.
    If Not confirmed Then Exit Sub

    ' Persist the preset and record the run only after the paste succeeds.
    If EconsRunPaste(cfg) Then
        cfg.SaveToRegistry
        Call EconsHistoryRecord(cfg)
    End If
End Sub

' Return the dropdown (list data-validation) options for the case input cell as
' a 0-based string array, or an empty Array() when the cell has no list
' validation or cannot be resolved. Used by the dialog so cases are picked from
' valid options rather than typed (an invalid case name breaks the model).
Public Function EconsCaseOptions(ByVal caseRangeName As String) As Variant
    Dim wb As Workbook, wsInputs As Worksheet, ws As Worksheet
    Dim caseCell As Range, srcRange As Range, cell As Range
    Dim formula As String
    Dim result() As String
    Dim count As Long
    Dim parts As Variant, i As Long, token As String
    Dim vType As Long

    EconsCaseOptions = Array()

    On Error GoTo Done
    If Trim$(caseRangeName) = "" Then GoTo Done
    Set wb = ActiveWorkbook
    If wb Is Nothing Then GoTo Done

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, TARGET_SHEET, vbTextCompare) = 0 Then
            Set wsInputs = ws
            Exit For
        End If
    Next ws

    Set caseCell = Nothing
    On Error Resume Next
    If Not wsInputs Is Nothing Then Set caseCell = wsInputs.Range(caseRangeName)
    If caseCell Is Nothing Then Set caseCell = wb.Names(caseRangeName).RefersToRange
    On Error GoTo Done
    If caseCell Is Nothing Then GoTo Done

    ' .Validation.Type raises when the cell has no validation rule.
    vType = -1
    On Error Resume Next
    vType = caseCell.Validation.Type
    formula = caseCell.Validation.Formula1
    On Error GoTo Done
    If vType <> xlValidateList Then GoTo Done
    If Trim$(formula) = "" Then GoTo Done

    ReDim result(0 To 1000)
    count = 0

    If Left$(formula, 1) = "=" Then
        ' Reference to a range or named range (resolved in the cell's sheet).
        Set srcRange = Nothing
        On Error Resume Next
        Set srcRange = caseCell.Parent.Evaluate(Mid$(formula, 2))
        On Error GoTo Done
        If srcRange Is Nothing Then GoTo Done
        For Each cell In srcRange.Cells
            token = Trim$(CStr(cell.Value))
            If token <> "" And count <= UBound(result) Then
                result(count) = token
                count = count + 1
            End If
        Next cell
    Else
        ' Literal list typed into the validation rule, separated by the locale's
        ' list separator (comma on US systems, semicolon on many others).
        Dim sep As String
        sep = Application.International(xlListSeparator)
        If sep = "" Then sep = ","
        parts = Split(formula, sep)
        For i = LBound(parts) To UBound(parts)
            token = Trim$(parts(i))
            If token <> "" And count <= UBound(result) Then
                result(count) = token
                count = count + 1
            End If
        Next i
    End If

    If count > 0 Then
        ReDim Preserve result(0 To count - 1)
        EconsCaseOptions = result
    End If

Done:
    On Error GoTo 0
End Function

' Run the PowerPoint paste using the supplied configuration. Returns True only
' when the export completes; returns False (after showing a message) on any
' precondition failure or error so the caller can skip saving the preset.
Private Function EconsRunPaste(ByVal cfg As cls_EconsConfig) As Boolean
    Dim originalCalc As XlCalculation
    Dim originalEvents As Boolean
    Dim originalScreen As Boolean

    Dim pptApp As Object, pptPres As Object
    Dim slide As Object, customLayout As Object
    Dim wb As Workbook, wsInputs As Worksheet, pic As Object
    Dim slideWidth As Single
    Dim caseCell As Range
    Dim ws As Worksheet, d As Object, cl As Object
    Dim cases As Variant
    Dim ci As Long, oi As Long
    Dim caseName As String, outName As String, rngName As String, title As String
    Dim rng As Range
    Dim originalCaseValue As Variant
    Dim missing As String
    Dim errorList As Object
    Dim summary As String, k As Long

    EconsRunPaste = False

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No active workbook found. Please open your Excel file and try again.", vbExclamation
        Exit Function
    End If

    Set wsInputs = Nothing
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, TARGET_SHEET, vbTextCompare) = 0 Then
            Set wsInputs = ws
            Exit For
        End If
    Next ws
    If wsInputs Is Nothing Then
        MsgBox "Worksheet '" & TARGET_SHEET & "' was not found in " & wb.Name & ".", vbExclamation
        Exit Function
    End If

    Set caseCell = Nothing
    On Error Resume Next
    Set caseCell = wsInputs.Range(cfg.CaseRangeName)
    If caseCell Is Nothing Then Set caseCell = wb.Names(cfg.CaseRangeName).RefersToRange
    On Error GoTo 0
    If caseCell Is Nothing Then
        MsgBox "Named cell or range '" & cfg.CaseRangeName & "' was not found on the Inputs sheet.", vbExclamation
        Exit Function
    End If

    ' Pre-flight: confirm every named range we will paste actually resolves, so a
    ' typo is reported once, up front, instead of once per case mid-run (and
    ' before any slides are created or PowerPoint is launched).
    missing = EconsMissingRanges(wb, cfg)
    If missing <> "" Then
        MsgBox "These named ranges were not found in " & wb.Name & ":" & vbCrLf & _
               missing & vbCrLf & vbCrLf & _
               "Check the output ranges and try again.", vbExclamation, "Econs"
        Exit Function
    End If

    On Error Resume Next
    Set pptApp = GetObject(Class:="PowerPoint.Application")
    If pptApp Is Nothing Then Set pptApp = CreateObject(Class:="PowerPoint.Application")
    On Error GoTo 0
    If pptApp Is Nothing Then
        MsgBox "Unable to start PowerPoint.", vbExclamation
        Exit Function
    End If
    pptApp.Visible = True

    If pptApp.Presentations.Count = 0 Then
        MsgBox "No PowerPoint presentations are open. Please open one and try again.", vbExclamation
        Exit Function
    End If
    Set pptPres = pptApp.ActivePresentation

    Set customLayout = Nothing
    For Each d In pptPres.Designs
        For Each cl In d.SlideMaster.CustomLayouts
            If LCase$(cl.Name) = LAYOUT_NAME Then
                Set customLayout = cl
                Exit For
            End If
        Next cl
        If Not customLayout Is Nothing Then Exit For
    Next d
    If customLayout Is Nothing Then
        MsgBox "Custom layout '" & LAYOUT_NAME & "' not found in the active presentation.", vbExclamation
        Exit Function
    End If
    slideWidth = pptPres.PageSetup.SlideWidth

    cases = cfg.CasesArray()

    ' Tracks slides already written to this run so two outputs never land on the
    ' same page when several slides share a title.
    Dim usedSlides As Object
    Set usedSlides = CreateObject("Scripting.Dictionary")

    ' Per-output problems are collected here and reported once at the end rather
    ' than popping a modal dialog mid-run (which would block, once per case).
    Set errorList = New Collection

    ' Heavy work begins; suspend recalculation/UI and restore on exit. The case
    ' input cell is also captured so the model is left exactly as the user had
    ' it, not stuck on the last case printed.
    originalCalc = Application.Calculation
    originalEvents = Application.EnableEvents
    originalScreen = Application.ScreenUpdating
    originalCaseValue = caseCell.Value

    On Error GoTo CleanFail
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For ci = LBound(cases) To UBound(cases)
        caseName = Trim$(cases(ci))
        If caseName <> "" Then
            caseCell.Value = caseName
            Application.CalculateFull
            DoEvents

            ' Standard outputs (apply to every selected case).
            For oi = 1 To cfg.NumOutputs
                Call EconsEmitOutput(pptPres, customLayout, slideWidth, wb, _
                                     cfg.OutputName(oi), cfg.OutputRange(oi), caseName, _
                                     cfg.ReplaceExisting, cfg.TitleCaseFirst, usedSlides, errorList)
            Next oi

            ' Advanced outputs (only for the cases they were assigned to).
            For oi = 1 To cfg.MaxAdvanced
                If cfg.AdvAppliesTo(oi, caseName) Then
                    Call EconsEmitOutput(pptPres, customLayout, slideWidth, wb, _
                                         cfg.AdvName(oi), cfg.AdvRange(oi), caseName, _
                                         cfg.ReplaceExisting, cfg.TitleCaseFirst, usedSlides, errorList)
                End If
            Next oi
        End If
    Next ci

    EconsRunPaste = True

CleanExit:
    ' Restore the model to its pre-run state: put the case cell back, restore the
    ' application settings, then recalculate so the workbook reflects the cell we
    ' just restored rather than the last case printed.
    On Error Resume Next
    If Not caseCell Is Nothing Then caseCell.Value = originalCaseValue
    Application.Calculation = originalCalc
    Application.ScreenUpdating = originalScreen
    Application.EnableEvents = originalEvents
    If Not caseCell Is Nothing Then Application.CalculateFull
    On Error GoTo 0

    ' Report any per-output problems once, after restoring the model.
    If Not errorList Is Nothing Then
        If errorList.Count > 0 Then
            summary = "The export finished, but some outputs had problems:" & vbCrLf
            For k = 1 To errorList.Count
                summary = summary & vbCrLf & "  - " & errorList(k)
            Next k
            MsgBox summary, vbExclamation, "Econs"
        End If
    End If
    Exit Function

CleanFail:
    MsgBox "Error running econs export: " & Err.Description, vbCritical
    EconsRunPaste = False
    Resume CleanExit
End Function

' Produce one output for one case: replace the matching slide's picture (replace
' mode) or append a new slide (append mode). A newly appended slide is titled in
' the chosen order; a replaced slide keeps its existing title (which may use the
' opposite order). Errors for a single output are reported but do not abort the
' whole run.
Private Sub EconsEmitOutput(ByVal pptPres As Object, ByVal customLayout As Object, _
    ByVal slideWidth As Single, ByVal wb As Workbook, _
    ByVal outName As String, ByVal rngName As String, ByVal caseName As String, _
    ByVal replaceMode As Boolean, ByVal caseFirst As Boolean, ByVal usedSlides As Object, _
    ByVal errorList As Object)

    Dim rng As Range, slide As Object, pic As Object

    Set rng = Nothing
    On Error Resume Next
    Set rng = wb.Names(rngName).RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then
        ' Pre-flight should already have caught this; record it and move on.
        errorList.Add "Named range '" & rngName & "' not found (output '" & outName & "')."
        Exit Sub
    End If

    On Error GoTo EmitFail

    Set slide = Nothing
    If replaceMode Then
        ' Match a slide titled with the output name and case in EITHER order; when
        ' several share a title, pick the one whose existing output best matches
        ' the shape about to be pasted.
        Set slide = EconsPickSlide(pptPres, outName, caseName, rng, usedSlides)
    End If

    If slide Is Nothing Then
        ' Append a new slide with the standard centered sizing and a fresh title.
        Set slide = pptPres.Slides.AddSlide(pptPres.Slides.Count + 1, customLayout)
        Set pic = EconsPasteRange(rng, slide)
        If pic Is Nothing Then
            slide.Delete
            errorList.Add "Could not paste output '" & outName & "' for case '" & caseName & "'."
            Exit Sub
        End If
        Call EconsApplyDefaultSizing(pic, slideWidth)
        Call EconsTagPicture(pic)

        On Error Resume Next
        If Not slide.Shapes.Title Is Nothing Then
            slide.Shapes.Title.TextFrame.TextRange.Text = EconsBuildTitle(outName, caseName, caseFirst)
        End If
        On Error GoTo EmitFail
    Else
        ' Replace the existing output picture; keep the slide's existing title
        ' (which may be in the opposite order).
        Call EconsReplaceOnSlide(slide, rng, slideWidth)
    End If

    ' Don't reuse this slide for another output in this run.
    On Error Resume Next
    usedSlides(CStr(slide.SlideID)) = True
    On Error GoTo 0
    Exit Sub

EmitFail:
    errorList.Add "Error creating output '" & outName & "' for case '" & caseName & "': " & Err.Description
End Sub

' Resolve every output range the run will paste (standard + complete advanced
' slots) and return a newline-bulleted list of those that do not resolve, or ""
' when all are present. Used for the up-front pre-flight check.
Private Function EconsMissingRanges(ByVal wb As Workbook, ByVal cfg As cls_EconsConfig) As String
    Dim oi As Long, missing As String

    For oi = 1 To cfg.NumOutputs
        If Not EconsRangeResolves(wb, cfg.OutputRange(oi)) Then
            missing = missing & vbCrLf & "  - Output " & oi & " (" & cfg.OutputName(oi) & _
                      "): '" & cfg.OutputRange(oi) & "'"
        End If
    Next oi

    For oi = 1 To cfg.MaxAdvanced
        If cfg.AdvIsComplete(oi) Then
            If Not EconsRangeResolves(wb, cfg.AdvRange(oi)) Then
                missing = missing & vbCrLf & "  - Extra output " & (oi + 2) & " (" & cfg.AdvName(oi) & _
                          "): '" & cfg.AdvRange(oi) & "'"
            End If
        End If
    Next oi

    EconsMissingRanges = missing
End Function

' True if rngName resolves to a range in wb. The named range address is stable
' across cases, so resolving it once up front is valid.
Private Function EconsRangeResolves(ByVal wb As Workbook, ByVal rngName As String) As Boolean
    Dim rng As Range

    On Error Resume Next
    Set rng = wb.Names(rngName).RefersToRange
    On Error GoTo 0
    EconsRangeResolves = Not (rng Is Nothing)
End Function

' Apply the standard centered figure geometry: scale to the target width, then
' center horizontally FIG_TOP_PTS below the slide top. Shared by append and by
' replace-onto-a-blank-slide so both produce identically sized figures.
Private Sub EconsApplyDefaultSizing(ByVal pic As Object, ByVal slideWidth As Single)
    With pic
        .LockAspectRatio = msoTrue
        If .Width > 0 Then .ScaleWidth FIG_WIDTH_PTS / .Width, msoFalse, msoScaleFromTopLeft
        .Left = (slideWidth - .Width) / 2
        .Top = FIG_TOP_PTS
    End With
End Sub

' Stamp the tool's tag onto a freshly pasted figure so replace mode can later
' identify it unambiguously. Best-effort; tagging failure is not fatal.
Private Sub EconsTagPicture(ByVal pic As Object)
    On Error Resume Next
    pic.AlternativeText = ECONS_TAG
    On Error GoTo 0
End Sub

' True if shp carries the econs tag (i.e. this tool pasted it).
Private Function EconsIsTagged(ByVal shp As Object) As Boolean
    Dim t As String

    t = ""
    On Error Resume Next
    t = shp.AlternativeText
    On Error GoTo 0
    EconsIsTagged = (StrComp(Trim$(t), ECONS_TAG, vbTextCompare) = 0)
End Function

' Build the slide title for an appended output in the configured order.
Private Function EconsBuildTitle(ByVal outName As String, ByVal caseName As String, _
    ByVal caseFirst As Boolean) As String

    Dim disp As String
    disp = caseName & " Case"
    If caseFirst Then
        EconsBuildTitle = disp & " | " & outName
    Else
        EconsBuildTitle = outName & " | " & disp
    End If
End Function

' Choose the slide to replace for this output. Considers only slides whose title
' contains the output name and case (in EITHER order, "Name | Case" or
' "Case | Name") and that have not already been written to this run. When more
' than one candidate remains, the one whose existing output picture's aspect
' ratio is closest to the range about to be pasted wins; ties fall back to
' document order (so output 1 takes the first such slide, output 2 the next).
' Returns Nothing if there is no match.
Private Function EconsPickSlide( _
    ByVal pptPres As Object, ByVal outName As String, ByVal caseName As String, _
    ByVal rng As Range, ByVal usedSlides As Object) As Object

    Dim s As Object, t As String
    Dim newAspect As Double, oldAspect As Double, diff As Double, bestDiff As Double
    Dim pic As Object, best As Object

    Set EconsPickSlide = Nothing

    newAspect = 0
    On Error Resume Next
    If rng.Height > 0 Then newAspect = rng.Width / rng.Height
    On Error GoTo 0

    bestDiff = 1E+99
    For Each s In pptPres.Slides
        If Not EconsSlideUsed(usedSlides, s) Then
            t = ""
            On Error Resume Next
            If s.Shapes.HasTitle Then t = s.Shapes.Title.TextFrame.TextRange.Text
            On Error GoTo 0

            If EconsTitleMatches(t, outName, caseName) Then
                ' Larger default so a slide with no readable picture loses to one
                ' that has a measurable, closely-matching output.
                diff = 1E+18
                If newAspect > 0 Then
                    Set pic = EconsMainPicture(s)
                    If Not pic Is Nothing Then
                        oldAspect = 0
                        On Error Resume Next
                        If pic.Height > 0 Then oldAspect = pic.Width / pic.Height
                        On Error GoTo 0
                        If oldAspect > 0 Then diff = Abs(newAspect - oldAspect)
                    End If
                End If

                If diff < bestDiff Then
                    bestDiff = diff
                    Set best = s
                End If
            End If
        End If
    Next s

    Set EconsPickSlide = best
End Function

' True if the slide has already been written to this run.
Private Function EconsSlideUsed(ByVal usedSlides As Object, ByVal s As Object) As Boolean
    Dim id As String

    EconsSlideUsed = False
    If usedSlides Is Nothing Then Exit Function

    id = ""
    On Error Resume Next
    id = CStr(s.SlideID)
    On Error GoTo 0
    If id <> "" Then EconsSlideUsed = usedSlides.Exists(id)
End Function

' True if slideTitle is made of the output name and the case (with or without
' the " Case" suffix) separated by "|", in either order. This lets replace mode
' match both "Name | Case Case" and "Case Case | Name".
Private Function EconsTitleMatches(ByVal slideTitle As String, _
    ByVal outName As String, ByVal caseName As String) As Boolean

    Dim key As String, parts As Variant
    Dim hA As String, hB As String
    Dim nOut As String, nCaseWith As String, nCaseWithout As String

    EconsTitleMatches = False
    key = EconsNormalizeTitle(slideTitle)
    If key = "" Then Exit Function
    If InStr(key, "|") = 0 Then Exit Function

    parts = Split(key, "|")
    If (UBound(parts) - LBound(parts)) <> 1 Then Exit Function

    hA = Trim$(parts(LBound(parts)))
    hB = Trim$(parts(LBound(parts) + 1))

    nOut = EconsNormalizeTitle(outName)
    nCaseWith = EconsNormalizeTitle(caseName & " Case")
    nCaseWithout = EconsNormalizeTitle(caseName)

    If hA = nOut And (hB = nCaseWith Or hB = nCaseWithout) Then
        EconsTitleMatches = True
    ElseIf hB = nOut And (hA = nCaseWith Or hA = nCaseWithout) Then
        EconsTitleMatches = True
    End If
End Function

' Normalise a slide title for comparison: strip CR/LF/VT/control characters,
' collapse internal whitespace, trim, and lower-case.
Private Function EconsNormalizeTitle(ByVal s As String) As String
    Dim r As String

    r = s
    r = Replace(r, vbCr, " ")
    r = Replace(r, vbLf, " ")
    r = Replace(r, Chr$(11), " ")   ' vertical tab (PowerPoint line break)
    r = Replace(r, Chr$(9), " ")    ' tab
    r = Replace(r, Chr$(7), " ")    ' cell/column separator
    r = Replace(r, Chr$(160), " ")  ' non-breaking space
    Do While InStr(r, "  ") > 0
        r = Replace(r, "  ", " ")
    Loop
    EconsNormalizeTitle = LCase$(Trim$(r))
End Function

' Replace the main output picture on a slide with a fresh paste of rng. The new
' figure mirrors the existing picture's WIDTH (shrunk to it only if wider, never
' enlarged), is anchored to the existing picture's top-left corner, inherits its
' z-order, and then the old picture is removed. If the slide has no picture yet,
' the new one is centered with the standard sizing.
Private Sub EconsReplaceOnSlide(ByVal slide As Object, ByVal rng As Range, ByVal slideWidth As Single)
    Dim oldPic As Object, newPic As Object
    Dim z As Long, k As Long

    Set oldPic = EconsMainPicture(slide)

    Set newPic = EconsPasteRange(rng, slide)
    If newPic Is Nothing Then Exit Sub      ' paste failed; leave the slide as-is
    newPic.LockAspectRatio = msoTrue
    Call EconsTagPicture(newPic)

    If oldPic Is Nothing Then
        Call EconsApplyDefaultSizing(newPic, slideWidth)
    Else
        z = oldPic.ZOrderPosition
        ' Mirror the existing image's width (top-left aligned, never enlarged).
        Call FitShapeToTargetWidth(newPic, oldPic.Left, oldPic.Top, oldPic.Width)
        oldPic.Delete
        If z > 0 Then
            newPic.ZOrder msoSendToBack
            For k = 1 To z - 1
                newPic.ZOrder msoBringForward
            Next k
        End If
    End If
End Sub

' Copy a range as a picture and paste it onto a slide, retrying on transient
' clipboard "not enough memory" errors, then releasing the clipboard. Returns
' the pasted shape, or Nothing if it could not be pasted.
Private Function EconsPasteRange(ByVal rng As Range, ByVal slide As Object) As Object
    Dim shp As Object, attempt As Long

    For attempt = 1 To 3
        On Error Resume Next
        rng.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
        On Error GoTo 0

        Set shp = PptTryPastePicture(slide)
        If Not shp Is Nothing Then
            Application.CutCopyMode = False
            Call ClearWindowsClipboard
            Set EconsPasteRange = shp
            Exit Function
        End If

        DoEvents
    Next attempt

    Set EconsPasteRange = Nothing
End Function

' The slide's main output shape (the largest non-placeholder shape, i.e. the
' pasted econs picture), or Nothing. Skipping placeholders avoids replacing the
' title/body text boxes; covers pictures pasted as image or OLE object.
Private Function EconsMainPicture(ByVal slide As Object) As Object
    Dim shp As Object, best As Object, taggedBest As Object
    Dim bestArea As Double, taggedArea As Double, area As Double

    bestArea = -1
    taggedArea = -1
    For Each shp In slide.Shapes
        If shp.Type <> msoPlaceholder Then
            area = 0
            On Error Resume Next
            area = shp.Width * shp.Height
            On Error GoTo 0

            If EconsIsTagged(shp) Then
                If area > taggedArea Then
                    taggedArea = area
                    Set taggedBest = shp
                End If
            End If

            If area > bestArea Then
                bestArea = area
                Set best = shp
            End If
        End If
    Next shp

    ' Prefer a figure this tool stamped; fall back to the largest non-placeholder
    ' shape for slides created before tagging existed.
    If Not taggedBest Is Nothing Then
        Set EconsMainPicture = taggedBest
    Else
        Set EconsMainPicture = best
    End If
End Function

' --- Run history ----------------------------------------------------------

' Record cfg as the most recent run. An exact duplicate already in the history
' is moved to the top rather than added again, so repeating the same run does
' not grow the list. The list is capped at HIST_MAX entries.
Public Sub EconsHistoryRecord(ByVal cfg As cls_EconsConfig)
    Dim newSer As String, s As String
    Dim sers() As String, times() As String
    Dim n As Long, m As Long, i As Long

    newSer = cfg.Serialize()
    n = EconsHistoryCount()

    ReDim sers(1 To HIST_MAX)
    ReDim times(1 To HIST_MAX)

    ' The new entry goes first.
    m = 1
    sers(1) = newSer
    times(1) = Format$(Now, "yyyy-mm-dd  hh:nn")

    ' Carry forward existing entries, skipping any exact duplicate of the new one.
    For i = 1 To n
        s = EconsHistorySerialized(i)
        If s <> "" And StrComp(s, newSer, vbBinaryCompare) <> 0 Then
            If m < HIST_MAX Then
                m = m + 1
                sers(m) = s
                times(m) = EconsHistoryTimestamp(i)
            End If
        End If
    Next i

    On Error Resume Next
    SaveSetting HIST_APP, HIST_SECTION, "Count", CStr(m)
    For i = 1 To m
        SaveSetting HIST_APP, HIST_SECTION, "Item" & i, sers(i)
        SaveSetting HIST_APP, HIST_SECTION, "Time" & i, times(i)
    Next i
    ' Remove any stale trailing entries left over from a previously longer list.
    For i = m + 1 To HIST_MAX
        DeleteSetting HIST_APP, HIST_SECTION, "Item" & i
        DeleteSetting HIST_APP, HIST_SECTION, "Time" & i
    Next i
    On Error GoTo 0
End Sub

' Number of stored history entries (0..HIST_MAX).
Public Function EconsHistoryCount() As Long
    Dim raw As String

    raw = GetSetting(HIST_APP, HIST_SECTION, "Count", "0")
    If IsNumeric(raw) Then EconsHistoryCount = CLng(raw)
    If EconsHistoryCount < 0 Then EconsHistoryCount = 0
    If EconsHistoryCount > HIST_MAX Then EconsHistoryCount = HIST_MAX
End Function

' The serialized config for history entry idx (1 = most recent), or "".
Public Function EconsHistorySerialized(ByVal idx As Long) As String
    EconsHistorySerialized = GetSetting(HIST_APP, HIST_SECTION, "Item" & idx, "")
End Function

' The display timestamp for history entry idx (1 = most recent), or "".
Public Function EconsHistoryTimestamp(ByVal idx As Long) As String
    EconsHistoryTimestamp = GetSetting(HIST_APP, HIST_SECTION, "Time" & idx, "")
End Function

' Prune a configuration loaded from history so it only references cases the
' current model offers. Unsupported cases (main and advanced) are removed in
' place; the return value is a user-facing note of what was dropped, or "" when
' nothing was dropped.
Public Function EconsPruneToModel(ByVal cfg As cls_EconsConfig) As String
    Dim options As Variant
    Dim kept As String, removed As String
    Dim dropped As String
    Dim i As Long

    options = EconsCaseOptions(cfg.CaseRangeName)
    If Not IsArray(options) Then options = Array()
    If UBound(options) < 0 Then
        EconsPruneToModel = "The case cell '" & cfg.CaseRangeName & _
            "' has no selectable cases in this model, so the saved cases could not be verified."
        Exit Function
    End If

    ' Main cases.
    Call EconsFilterCsv(cfg.CasesText, options, kept, removed)
    cfg.CasesText = kept
    If removed <> "" Then dropped = dropped & vbCrLf & "  - Cases: " & removed

    ' Advanced slots.
    For i = 1 To cfg.MaxAdvanced
        If Trim$(cfg.AdvName(i)) <> "" Or Trim$(cfg.AdvRange(i)) <> "" Or Trim$(cfg.AdvCasesText(i)) <> "" Then
            Call EconsFilterCsv(cfg.AdvCasesText(i), options, kept, removed)
            cfg.AdvCasesText(i) = kept
            If removed <> "" Then _
                dropped = dropped & vbCrLf & "  - " & cfg.AdvName(i) & " (extra output " & (i + 2) & "): " & removed
        End If
    Next i

    If dropped <> "" Then
        EconsPruneToModel = "Some saved cases aren't available in this model and were left unchecked:" & dropped
    End If
End Function

' Split a comma-separated list into the tokens present in options (keptCsv) and
' those that are not (removedCsv), both as comma-separated strings.
Private Sub EconsFilterCsv(ByVal csv As String, ByVal options As Variant, _
    ByRef keptCsv As String, ByRef removedCsv As String)

    Dim parts As Variant, i As Long, j As Long, tok As String, found As Boolean

    keptCsv = ""
    removedCsv = ""
    If Trim$(csv) = "" Then Exit Sub

    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        tok = Trim$(parts(i))
        If tok <> "" Then
            found = False
            For j = LBound(options) To UBound(options)
                If StrComp(tok, CStr(options(j)), vbTextCompare) = 0 Then
                    found = True
                    Exit For
                End If
            Next j
            If found Then
                If keptCsv <> "" Then keptCsv = keptCsv & ", "
                keptCsv = keptCsv & tok
            Else
                If removedCsv <> "" Then removedCsv = removedCsv & ", "
                removedCsv = removedCsv & tok
            End If
        End If
    Next i
End Sub
