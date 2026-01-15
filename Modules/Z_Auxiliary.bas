Attribute VB_Name = "Z_Auxiliary"
Option Explicit
Option Private Module
Private currentSeriesIndex As Long
Public gScrollLockMode As Boolean
Private Const DATA_LABEL_STEP As Double = 1#
Private Const CHART_MOVE_STEP As Double = 50#

' Shared wrapper for temporarily disabling events / screen updates.

Private Function SuppressExcelUi(Optional ByVal hideStatusBar As Boolean = False) As ExcelUiGuard
    Dim guard As ExcelUiGuard
    ThisWorkbook.EnsureAppHook
    Set guard = New ExcelUiGuard
    guard.EnsureWindow
    If hideStatusBar Then guard.DisableStatusBar
    Set SuppressExcelUi = guard
End Function

Function ColorSelectionYellow(Optional ByVal g As String) As Boolean
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    On Error GoTo CleanFail

    If TypeName(Selection) = "Range" Then
        Selection.Interior.Color = vbYellow
    End If

CleanExit:
    ColorSelectionYellow = False
    Exit Function
CleanFail:
    Call ErrorHandler("ColorSelectionYellow")
    Resume CleanExit
End Function

Public Function SubstituteType(Optional ByVal g As String) As Boolean
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    On Error GoTo CleanFail
    ' TESTING FILE UPLOAD
    Call SubstituteFollowLangMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.SubstituteType g

CleanExit:
    SubstituteType = False
    Exit Function
CleanFail:
    Call ErrorHandler("SubstituteType")
    Resume CleanExit
End Function

Public Function SubstituteInsertEquals(Optional ByVal g As String) As Boolean
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    On Error GoTo CleanFail

    Call SubstituteFollowLangMode
    KeyStroke Equal_US_

CleanExit:
    SubstituteInsertEquals = False
    Exit Function
CleanFail:
    Call ErrorHandler("SubstituteInsertEquals")
    Resume CleanExit
End Function


Function TogglePlainKeyMappings(Optional ByVal g As String) As Boolean
    Dim uiGuard As ExcelUiGuard
    Dim statusMessage As String
    Set uiGuard = SuppressExcelUi(True)
    If gVim Is Nothing Then Exit Function

    If Not gVim.KeyMap.SuppressPlainKeys Then
        gVim.KeyMap.SuppressPlainKeys = True
        statusMessage = "Excel native shortcuts restored."
    Else
        gVim.KeyMap.SuppressPlainKeys = False
        statusMessage = "Vim shortcuts re-enabled."
    End If

    TogglePlainKeyMappings = False
    Call ClipboardRefresh
    Call SetStatusBarTemporarily(statusMessage, 2000)
End Function

Function ToggleMacroSafeMode(Optional ByVal g As String) As Boolean
    Dim statusMessage As String
    gMacroSafeMode = Not gMacroSafeMode

    If gMacroSafeMode Then
        statusMessage = "Fast mode ON: auto-color and outline highlight disabled."
    Else
        statusMessage = "Fast mode OFF: auto features restored."
    End If

    Call SetStatusBarTemporarily(statusMessage, 2500)
    ToggleMacroSafeMode = False
End Function

Function OverrideShortcuts(Optional ByVal g As String) As Boolean
    Dim uiGuard As ExcelUiGuard
    Dim statusMessage As String
    On Error GoTo CleanFail

    Set uiGuard = SuppressExcelUi(True)
    Call StartVim

    If Not gVim Is Nothing Then
        Call gVim.KeyMap.BindAll
        Call ClipboardRefresh
        statusMessage = "Vantage shortcuts now overriding other add-ins."
        Call SetStatusBarTemporarily(statusMessage, 2500)
    End If

CleanExit:
    OverrideShortcuts = False
    Exit Function

CleanFail:
    Call ErrorHandler("OverrideShortcuts")
    Resume CleanExit
End Function

Public Function LaunchResearchLink(Optional ByVal g As String) As Boolean
    Dim target As String
    Select Case LCase$(Trim$(g))
        Case "edgar": target = "https://www.sec.gov/search-filings"
        Case "alpha": target = "https://research.alpha-sense.com/gensearch"
        Case "fred": target = "https://fred.stlouisfed.org/"
        Case "bam": target = "https://www.bamsec.com/"
        Case "cap": target = "https://www.capitaliq.com/CIQDotNet/my/dashboard.aspx"
        Case Else: Exit Function
    End Select

    On Error Resume Next
    ThisWorkbook.FollowHyperlink Address:=target, NewWindow:=True
    LaunchResearchLink = False
End Function

Function ToggleScrollLockMode(Optional ByVal g As String) As Boolean
    Dim uiGuard As ExcelUiGuard
    Dim statusMessage As String
    Set uiGuard = SuppressExcelUi(True)

    gScrollLockMode = Not gScrollLockMode
    If gScrollLockMode Then
        ScrollLockCenterSelection
        statusMessage = "Scroll lock mode enabled. Use Ctrl+F11 to exit."
    Else
        statusMessage = "Scroll lock mode disabled."
    End If

    gVim.Count1 = 1
    ToggleScrollLockMode = False
    Call SetStatusBarTemporarily(statusMessage, 2500)
End Function

Private Sub ScrollLockCenterSelection()
    On Error GoTo CleanExit
    Dim win As Window
    Dim vis As Range
    Dim targetRow As Long
    Dim targetCol As Long

    Set win = ActiveWindow
    If win Is Nothing Then Exit Sub
    Set vis = win.VisibleRange
    If vis Is Nothing Then Exit Sub

    targetRow = vis.Row + vis.Rows.Count \ 2
    targetCol = vis.Column + vis.Columns.Count \ 2

    If targetRow < 1 Then targetRow = 1
    If targetCol < 1 Then targetCol = 1
    If targetRow > win.ActiveSheet.Rows.Count Then targetRow = win.ActiveSheet.Rows.Count
    If targetCol > win.ActiveSheet.Columns.Count Then targetCol = win.ActiveSheet.Columns.Count

    win.ActiveSheet.Cells(targetRow, targetCol).Select
CleanExit:
End Sub

Private Sub ScrollLockScroll(ByVal deltaRows As Long, ByVal deltaCols As Long)
    On Error GoTo CleanExit
    Dim win As Window
    Dim vis As Range
    Dim Sh As Worksheet
    Dim targetRow As Long
    Dim targetCol As Long

    Set win = ActiveWindow
    If win Is Nothing Then Exit Sub
    Set Sh = win.ActiveSheet
    If Sh Is Nothing Then Exit Sub

    If deltaRows > 0 Then win.SmallScroll Down:=deltaRows
    If deltaRows < 0 Then win.SmallScroll Up:=-deltaRows
    If deltaCols > 0 Then win.SmallScroll ToRight:=deltaCols
    If deltaCols < 0 Then win.SmallScroll ToLeft:=-deltaCols

    On Error Resume Next
    Set vis = win.VisibleRange
    On Error GoTo CleanExit
    If vis Is Nothing Then Exit Sub

    targetRow = vis.Row + vis.Rows.Count \ 2
    targetCol = vis.Column + vis.Columns.Count \ 2

    If targetRow < 1 Then targetRow = 1
    If targetCol < 1 Then targetCol = 1
    If targetRow > Sh.Rows.Count Then targetRow = Sh.Rows.Count
    If targetCol > Sh.Columns.Count Then targetCol = Sh.Columns.Count

    Sh.Cells(targetRow, targetCol).Select
CleanExit:
    Set vis = Nothing
End Sub

Function CycleFillColor(Optional ByVal g As String) As Boolean
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    Dim colors As Variant
    Static lastIndex As Long
    Static lastAddress As String
    Static lastSelectionStamp As Long
    Dim currentAddress As String
    colors = Array(RGB(0, 32, 96), RGB(226, 234, 250), _
                   RGB(240, 240, 240), RGB(255, 242, 204), xlNone)
    ' Check selection
    If TypeName(Selection) <> "Range" Then Exit Function
    currentAddress = Selection.Address
    ' Reset index if new selection or cursor moved
    If currentAddress <> lastAddress _
        Or gSelectionStamp <> lastSelectionStamp Then
        lastIndex = 0
    End If
    lastAddress = currentAddress
    lastSelectionStamp = gSelectionStamp
    ' Apply color
    If colors(lastIndex) = xlNone Then
        Selection.Interior.ColorIndex = xlNone
    Else
        Selection.Interior.Color = colors(lastIndex)
    End If
    ' Advance to next color
    lastIndex = lastIndex + 1
    If lastIndex > UBound(colors) Then lastIndex = 0
End Function

Sub CycleFontColor()
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    Static fontColorsArray As Variant
    Static fontCycleIndex As Long
    Static fontCycleLastStamp As Long
    ' Exact colors
    If IsEmpty(fontColorsArray) Then
        fontColorsArray = Array(RGB(0, 0, 0), RGB(255, 255, 255), RGB(0, 0, 255), RGB(0, 128, 0), RGB(153, 0, 0))
    End If
    ' Exit if selection is not a range
    If TypeName(Selection) <> "Range" Then Exit Sub
    ' Reset cycle if selection changes or cursor moved
    If gSelectionStamp <> fontCycleLastStamp Then
        fontCycleIndex = 0
    End If
    fontCycleLastStamp = gSelectionStamp
    ' Apply color
    Selection.Font.Color = fontColorsArray(fontCycleIndex)
    ' Advance index
    fontCycleIndex = fontCycleIndex + 1
    If fontCycleIndex > UBound(fontColorsArray) Then fontCycleIndex = 0
End Sub

Function CycleNumberFormat(Optional ByVal g As String) As Boolean
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.CycleNumberFormat
CleanExit:
    CycleNumberFormat = False
    Exit Function
CleanFail:
    Call ErrorHandler("CycleNumberFormat")
    Resume CleanExit
End Function

Function BinaryCycle(Optional ByVal g As String) As Boolean
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.BinaryCycle
CleanExit:
    BinaryCycle = False
    Exit Function
CleanFail:
    Call ErrorHandler("BinaryCycle")
    Resume CleanExit
End Function

Function YearDisplayCycle(Optional ByVal g As String) As Boolean
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.YearDisplayCycle
CleanExit:
    YearDisplayCycle = False
    Exit Function
CleanFail:
    Call ErrorHandler("YearDisplayCycle")
    Resume CleanExit
End Function

Public Sub ResizeSelectionToWidthInches(Optional ByVal targetInches As Double = 4.7, Optional ByVal requirePpt As Boolean = False)
    Dim engine As Object
    On Error GoTo CleanFail

    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.ResizeSelectionToWidthInches targetInches, requirePpt

CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("ResizeSelectionToWidthInches")
    Resume CleanExit
End Sub

Public Function ResizeSelectionToWidthPrompt(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    If TypeName(Selection) <> "Range" Then GoTo CleanExit

    Dim widthValue As Variant
    widthValue = Application.InputBox( _
        Prompt:="Target width in inches (as printed):", _
        Title:="Resize to printed width", _
        Default:=4.7, _
        Type:=1)
    If widthValue = False Then GoTo CleanExit

    Dim targetInches As Double
    targetInches = CDbl(widthValue)
    If targetInches <= 0 Then GoTo CleanExit

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    Call StopVisualMode
    engine.ResizeSelectionToWidthInches targetInches, True

CleanExit:
    ResizeSelectionToWidthPrompt = False
    Exit Function
CleanFail:
    Call ErrorHandler("ResizeSelectionToWidthPrompt")
    Resume CleanExit
End Function

Private Sub SnapWidthsToTarget(ByVal rng As Range, ByVal targetWidthPts As Double)
    Dim nonAdjustPts As Double
    Dim adjustablePts As Double
    Dim c As Range
    Dim desiredAdj As Double
    Dim scaleFactor As Double
    Dim targetColWidth As Double

    On Error GoTo CleanFail

    ' Calculate adjustable vs non-adjustable widths
    For Each c In rng.Columns
        If c.ColumnWidth > 0.5 Then
            adjustablePts = adjustablePts + c.Width
        Else
            nonAdjustPts = nonAdjustPts + c.Width
        End If
    Next c
    If adjustablePts <= 0 Then GoTo CleanExit

    desiredAdj = targetWidthPts - nonAdjustPts
    If desiredAdj <= 0 Then GoTo CleanExit

    scaleFactor = desiredAdj / adjustablePts
    If scaleFactor <= 0 Then GoTo CleanExit

    ' Apply uniform scaling to adjustable columns
    For Each c In rng.Columns
        If c.ColumnWidth > 0.5 Then
            targetColWidth = c.ColumnWidth * scaleFactor
            If targetColWidth < 0.1 Then targetColWidth = 0.1
            c.EntireColumn.ColumnWidth = targetColWidth
        End If
    Next c

CleanExit:
    Exit Sub
CleanFail:
    Resume CleanExit
End Sub

Private Function MeasureCopyPictureWidthPts(ByVal rng As Range, ByVal pptSlide As Object) As Double
    On Error GoTo CleanFail
    Dim pastedShp As Object
    Dim pasteObj As Object
    Dim attempt As Long
    Const PP_PASTE_ENHANCED_METAFILE As Long = 2

    If rng Is Nothing Then GoTo CleanFail
    If pptSlide Is Nothing Then GoTo CleanFail

    For attempt = 1 To 2
        If Not CopySelectionAsPicturePrintSafe(rng) Then GoTo NextAttempt

        On Error Resume Next
        Set pasteObj = pptSlide.Shapes.PasteSpecial(DataType:=PP_PASTE_ENHANCED_METAFILE)
        If pasteObj Is Nothing Then Set pasteObj = pptSlide.Shapes.Paste
        On Error GoTo CleanFail

        If Not pasteObj Is Nothing Then Exit For
NextAttempt:
        DoEvents
    Next attempt

    If pasteObj Is Nothing Then GoTo CleanFail

    If TypeName(pasteObj) = "ShapeRange" Then
        Set pastedShp = pasteObj(1)
    Else
        Set pastedShp = pasteObj
    End If
    If pastedShp Is Nothing Then GoTo CleanFail

    MeasureCopyPictureWidthPts = pastedShp.Width

CleanExit:
    On Error Resume Next
    If Not pasteObj Is Nothing Then pasteObj.Delete
    Exit Function
CleanFail:
    MeasureCopyPictureWidthPts = 0
    Resume CleanExit
End Function

Private Function TryGetPowerPointContext(ByRef pptApp As Object, ByRef pptWindow As Object, ByRef pptPres As Object, ByRef pptSlide As Object) As Boolean
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then Exit Function
    Set pptWindow = pptApp.ActiveWindow
    If pptWindow Is Nothing Then Exit Function
    Set pptPres = pptWindow.Presentation
    If pptPres Is Nothing Then Exit Function
    Set pptSlide = pptWindow.View.Slide
    If pptSlide Is Nothing Then Exit Function
    TryGetPowerPointContext = True
End Function

Private Function IsPowerPointReady() As Boolean
    Dim pptApp As Object
    Dim pptWindow As Object
    Dim pptPres As Object
    Dim pptSlide As Object
    IsPowerPointReady = TryGetPowerPointContext(pptApp, pptWindow, pptPres, pptSlide)
End Function

Private Sub ScaleSelectionColumns(ByVal rng As Range, ByVal scaleFactor As Double, Optional ByVal minAdjustWidth As Double = -1)
    On Error Resume Next
    If rng Is Nothing Then Exit Sub
    If scaleFactor <= 0 Then Exit Sub

    Dim c As Range
    Dim wPts As Double
    Dim wCol As Double
    Dim targetColWidth As Double

    For Each c In rng.Columns
        If c.EntireColumn.Hidden Then GoTo NextCol
        If minAdjustWidth >= 0 Then
            If c.ColumnWidth <= minAdjustWidth Then GoTo NextCol
        End If
        wPts = c.Width
        If wPts <= 0 Then GoTo NextCol
        wCol = c.ColumnWidth
        targetColWidth = wCol * scaleFactor
        If targetColWidth < 0.1 Then targetColWidth = 0.1
        If targetColWidth > 255# Then targetColWidth = 255#
        c.EntireColumn.ColumnWidth = targetColWidth
NextCol:
    Next c
End Sub

Private Sub GetColumnWidthBuckets(ByVal rng As Range, ByVal minAdjustWidth As Double, ByRef fixedPts As Double, ByRef adjustablePts As Double)
    fixedPts = 0
    adjustablePts = 0
    If rng Is Nothing Then Exit Sub

    Dim c As Range
    For Each c In rng.Columns
        If c.EntireColumn.Hidden Then GoTo NextCol
        If c.ColumnWidth <= minAdjustWidth Then
            fixedPts = fixedPts + c.Width
        Else
            adjustablePts = adjustablePts + c.Width
        End If
NextCol:
    Next c
End Sub

Private Function GetPrintScaleFactor(ByVal ws As Worksheet) As Double
    On Error Resume Next
    Dim zoomValue As Variant
    zoomValue = ws.PageSetup.Zoom
    On Error GoTo 0
    If IsNumeric(zoomValue) Then
        Dim z As Double
        z = CDbl(zoomValue)
        If z > 0 Then
            GetPrintScaleFactor = z / 100#
            Exit Function
        End If
    End If
    GetPrintScaleFactor = 1#
End Function

Public Sub FormatOverviewGraph()
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    On Error GoTo CleanFail

    Dim chtObj As ChartObject
    Dim ser As Series

    ' Detect chart selection
    If TypeName(Selection) = "ChartArea" Or TypeName(Selection) = "PlotArea" Then
        Set chtObj = ActiveChart.Parent
    ElseIf TypeName(Selection) = "ChartObject" Then
        Set chtObj = Selection
    Else
        MsgBox "Please select a chart before running this macro.", vbExclamation
        GoTo CleanExit
    End If

    With chtObj
        On Error Resume Next
        If .Chart.HasTitle Then .Chart.HasTitle = False
        .Chart.Axes(xlValue).HasMajorGridlines = False
        .Chart.Axes(xlValue).Delete
        .Chart.HasLegend = False

        .Chart.ChartType = xlColumnStacked100
        .Chart.PlotBy = IIf(.Chart.PlotBy = xlRows, xlColumns, xlRows)
        If .Chart.ChartGroups.Count > 0 Then .Chart.ChartGroups(1).GapWidth = 5

        For Each ser In .Chart.SeriesCollection
            With ser.Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(0, 0, 0)
                .Weight = 0.75
            End With
            ser.ApplyDataLabels ShowSeriesName:=True, ShowValue:=True
            With ser.DataLabels
                .Font.Name = "Garamond"
                .Font.Size = 11
                .Font.Color = RGB(0, 0, 0)
                .Font.Bold = True
                .Separator = ", "
            End With
        Next ser

        .Height = 149.76
        .Width = 357.84
        .IncrementLeft 9.75
        .IncrementTop -4.5

        With .Chart.PlotArea
            .Top = 15
            .Left = 9
            .Width = 335
            .Height = 130
        End With

        With .Chart.Axes(xlCategory)
            .MajorTickMark = xlOutside
            With .Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(0, 0, 0)
                .Weight = 1
            End With
        End With

        .Chart.ChartArea.Font.Name = "Garamond"
        .Chart.ChartArea.Font.Color = RGB(0, 0, 0)

        With .ShapeRange.Line
            .Visible = msoFalse
        End With

        With .Chart.Axes(xlCategory).TickLabels
            .Font.Name = "Garamond"
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = RGB(0, 0, 0)
        End With
        On Error GoTo 0
    End With

CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("FormatOverviewGraph")
    Resume CleanExit
End Sub

Sub CycleChartType()
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)

    Dim chartObj As Chart
    On Error Resume Next
    Set chartObj = ActiveChart
    If chartObj Is Nothing Then
        If TypeName(Selection) = "ChartObject" Then
            Set chartObj = Selection.Chart
        ElseIf TypeName(Selection) = "ChartArea" Or TypeName(Selection) = "PlotArea" Then
            Set chartObj = ActiveChart
        End If
    End If
    On Error GoTo 0
    If chartObj Is Nothing Then Exit Sub

    Static idx As Long
    Static lastStamp As Long
    If gSelectionStamp <> lastStamp Then idx = 0
    lastStamp = gSelectionStamp

    Dim types As Variant
    types = Array(xlColumnClustered, xlLine, xlColumnStacked100, xlXYScatter)

    On Error Resume Next
    chartObj.ChartType = types(idx)
    ' Force each series into the target type for reliable switching
    Dim i As Long
    For i = 1 To chartObj.FullSeriesCollection.Count
        chartObj.FullSeriesCollection(i).ChartType = types(idx)
    Next i
    ' Stacking fallback: ensure 100% stacked with proper overlap/plot orientation
    If types(idx) = xlColumnStacked100 Then
        chartObj.ChartType = xlColumnStacked100
        For i = 1 To chartObj.FullSeriesCollection.Count
            chartObj.FullSeriesCollection(i).ChartType = xlColumnStacked100
        Next i
        chartObj.PlotBy = IIf(chartObj.PlotBy = xlRows, xlColumns, xlRows)
        If chartObj.ChartGroups.Count > 0 Then
            chartObj.ChartGroups(1).GapWidth = 30
            chartObj.ChartGroups(1).Overlap = 100
        End If
        ' If still not stacked, try columns plot
        If chartObj.ChartType <> xlColumnStacked100 Then
            chartObj.PlotBy = xlColumns
            chartObj.ChartType = xlColumnStacked100
            For i = 1 To chartObj.FullSeriesCollection.Count
                chartObj.FullSeriesCollection(i).ChartType = xlColumnStacked100
            Next i
            If chartObj.ChartGroups.Count > 0 Then
                chartObj.ChartGroups(1).GapWidth = 30
                chartObj.ChartGroups(1).Overlap = 100
            End If
        End If
    Else
        ' Reset spacing/overlap that might linger from stacked columns for all other chart types
        On Error Resume Next
        chartObj.PlotBy = xlColumns
        If chartObj.ChartGroups.Count > 0 Then
            chartObj.ChartGroups(1).GapWidth = 150
            chartObj.ChartGroups(1).Overlap = 0
        End If
        On Error GoTo 0
    End If
    On Error GoTo 0

    idx = idx + 1
    If idx > UBound(types) Then idx = 0
End Sub

Function NumberNarrativeCycle(Optional ByVal g As String) As Boolean
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.NumberNarrativeCycle
CleanExit:
    NumberNarrativeCycle = False
    Exit Function
CleanFail:
    Call ErrorHandler("NumberNarrativeCycle")
    Resume CleanExit
End Function

Function PercentCycle(Optional ByVal g As String) As Boolean
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.PercentCycle
CleanExit:
    PercentCycle = False
    Exit Function
CleanFail:
    Call ErrorHandler("PercentCycle")
    Resume CleanExit
End Function

Function CurrencyCycle(Optional ByVal g As String) As Boolean
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.CurrencyCycle
CleanExit:
    CurrencyCycle = False
    Exit Function
CleanFail:
    Call ErrorHandler("CurrencyCycle")
    Resume CleanExit
End Function

Public Sub DeleteLikeExcel()
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    On Error GoTo CleanFail
    Dim sel As Object
    Set sel = Selection

    If sel Is Nothing Then GoTo CleanExit

    ' If a chart is active, delete the chart object (covers chart elements too).
    Dim activeCh As Object
    On Error Resume Next
    Set activeCh = ActiveChart
    On Error GoTo CleanFail
    If Not activeCh Is Nothing Then
        activeCh.Parent.Delete
        GoTo CleanExit
    End If

    If TypeName(sel) = "Range" Then
        sel.ClearContents
    ElseIf TypeName(sel) = "ChartObject" Then
        sel.Delete
    ElseIf TypeName(sel) = "Shape" Or TypeName(sel) = "ShapeRange" Then
        sel.Delete
    ElseIf VarType(sel) = vbObject Then
        On Error Resume Next
        sel.Delete
        If Err.Number <> 0 Then
            Err.Clear
            Dim parentObj As Object
            Set parentObj = Nothing
            Set parentObj = sel.Parent
            If Not parentObj Is Nothing Then
                parentObj.Delete
            End If
        End If
        On Error GoTo CleanFail
    End If
CleanExit:
    On Error GoTo 0
    Exit Sub
CleanFail:
    ' Swallow any delete errors to mirror native behaviour
    Resume CleanExit
End Sub

Sub ClearFormatting()
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.ClearFormatting
CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("ClearFormatting")
    Resume CleanExit
End Sub

Sub CycleRowHeight()
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    Dim sel As Range
    Dim currentHeight As Double
    Dim nextHeight As Double
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection
    currentHeight = sel.Rows(1).RowHeight
    ' Decide next height
    If Abs(currentHeight - 3) < 0.1 Then
        nextHeight = 15
    Else
        nextHeight = 3
    End If
    ' Apply to all selected rows
    sel.EntireRow.RowHeight = nextHeight
End Sub

Sub CycleColumnWidth()
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    Dim sel As Range
    Dim currentWidth As Double
    Dim nextWidth As Double
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection
    ' Use the first column in selection to determine current width
    currentWidth = sel.Columns(1).ColumnWidth
    ' Decide next width
    If Abs(currentWidth - 8.43) < 0.1 Then
        nextWidth = 0.5
    Else
        nextWidth = 8.43
    End If
    ' Apply width to all selected columns
    sel.EntireColumn.ColumnWidth = nextWidth
End Sub

Sub CycleBorder(side As String)
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    Dim borderStyles As Variant
    Static lastIndex As Long
    Static lastAddress As String
    Static lastSelectionStamp As Long
    Static lastSide As String
    Dim rng As Range
    Dim b As border
    Dim currentAddress As String
    Dim sideKey As String
    ' Only run on range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set rng = Selection
    currentAddress = rng.Address
    ' Reset cycle if selection changes or cursor moved
    sideKey = LCase$(side)
    If currentAddress <> lastAddress _
        Or gSelectionStamp <> lastSelectionStamp _
        Or sideKey <> lastSide Then
        lastIndex = 0
    End If
    lastAddress = currentAddress
    lastSelectionStamp = gSelectionStamp
    ' Border styles to cycle: dashed (first), double, none
    borderStyles = Array(xlDash, xlDouble, xlNone)
    ' Apply to the specified side
    Select Case sideKey
        Case "h": Set b = rng.Borders(xlEdgeLeft)
        Case "j": Set b = rng.Borders(xlEdgeBottom)
        Case "k": Set b = rng.Borders(xlEdgeTop)
        Case "l": Set b = rng.Borders(xlEdgeRight)
        Case Else: Exit Sub
    End Select
    lastSide = sideKey
    b.lineStyle = borderStyles(lastIndex)
    ' Advance cycle
    lastIndex = lastIndex + 1
    If lastIndex > UBound(borderStyles) Then lastIndex = 0
    
    ' === UI restore & final highlight ===
    Set uiGuard = Nothing
    On Error Resume Next
    SafeSelectRange rng
    On Error GoTo 0

End Sub

Sub CycleBorderLeft(): CycleBorder "h": End Sub

Sub CycleBorderBottom(): CycleBorder "j": End Sub

Sub CycleBorderTop(): CycleBorder "k": End Sub

Sub CycleBorderRight(): CycleBorder "l": End Sub

Sub InsertHyperlinkDialog()
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    On Error GoTo Catch

    Dim linkColor As Long
    linkColor = RGB(0, 102, 255)

    ' Align workbook hyperlink styles to the desired link color.
    On Error Resume Next
    ActiveWorkbook.Styles("Hyperlink").Font.Color = linkColor
    ActiveWorkbook.Styles("Followed Hyperlink").Font.Color = linkColor
    On Error GoTo Catch

    Application.Dialogs(xlDialogInsertHyperlink).Show

    ' Re-apply color to any hyperlinks that were inserted on the current range.
    If TypeName(Selection) = "Range" Then
        Dim hl As Hyperlink
        For Each hl In Selection.Hyperlinks
            hl.Range.Font.Color = linkColor
        Next hl
    End If
    Exit Sub

Catch:
    Call ErrorHandler("InsertHyperlinkDialog")
End Sub

Sub CopyPasteAsPictureToPPT()
    On Error GoTo CleanFail

    Dim pptApp As Object
    Dim pptWindow As Object
    Dim pptSlide As Object
    Dim pptShape As Object
    Dim pastedShp As Object
    Dim sel As Object
    Dim hadTarget As Boolean
    Dim targetLeft As Single, targetTop As Single, targetWidth As Single, targetHeight As Single
    Dim desiredZ As Long, i As Long

    If TypeName(Selection) = "Nothing" Then
        MsgBox "Please select a range, chart, or shape first.", vbExclamation
        Exit Sub
    End If
    Set sel = Selection

    If Not CopySelectionAsPicturePrintSafe(sel) Then
        MsgBox "Unable to copy selection as picture (print view). Try selecting a different range or chart.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    On Error GoTo CleanFail
    If pptApp Is Nothing Then
        MsgBox "PowerPoint is not running. Please open a presentation.", vbExclamation
        Exit Sub
    End If

    Set pptWindow = pptApp.ActiveWindow
    If pptWindow Is Nothing Then
        MsgBox "No active PowerPoint window detected. Please select a slide and try again.", vbExclamation
        Exit Sub
    End If

    Set pptSlide = pptWindow.View.Slide
    If pptSlide Is Nothing Then
        MsgBox "Please make sure a slide is selected in PowerPoint before running this command.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set pptShape = pptWindow.Selection.ShapeRange
    On Error GoTo 0
    If Not pptShape Is Nothing Then
        hadTarget = True
        targetLeft = pptShape.Left
        targetTop = pptShape.Top
        targetWidth = pptShape.Width
        targetHeight = pptShape.Height
        If TypeName(pptShape) = "ShapeRange" And pptShape.Count > 0 Then
            desiredZ = pptShape(1).ZOrderPosition
        Else
            desiredZ = pptShape.ZOrderPosition
        End If
    End If

    On Error Resume Next
    Set pastedShp = pptSlide.Shapes.PasteSpecial(DataType:=2)
    If pastedShp Is Nothing Then Set pastedShp = pptSlide.Shapes.Paste
    On Error GoTo CleanFail

    If TypeName(pastedShp) = "ShapeRange" Then
        Set pastedShp = pastedShp(1)
    End If

    If hadTarget Then
        Dim scaledWidth As Double
        Dim scaledHeight As Double
        Dim scaleFactor As Double
        Dim scaleY As Double
        Dim offsetLeft As Double
        Dim offsetTop As Double

        scaledWidth = pastedShp.Width
        scaledHeight = pastedShp.Height

        If pastedShp.Width > 0 And pastedShp.Height > 0 Then
            If targetWidth > 0 And targetHeight > 0 Then
                scaleFactor = targetWidth / pastedShp.Width
                scaleY = targetHeight / pastedShp.Height
                If scaleY < scaleFactor Then scaleFactor = scaleY
            ElseIf targetWidth > 0 Then
                scaleFactor = targetWidth / pastedShp.Width
            ElseIf targetHeight > 0 Then
                scaleFactor = targetHeight / pastedShp.Height
            Else
                scaleFactor = 1
            End If

            If scaleFactor > 0 Then
                scaledWidth = pastedShp.Width * scaleFactor
                scaledHeight = pastedShp.Height * scaleFactor
            End If
        End If

        pastedShp.Width = scaledWidth
        pastedShp.Height = scaledHeight

        If targetWidth > 0 Then
            offsetLeft = targetLeft + (targetWidth - scaledWidth) / 2
        Else
            offsetLeft = targetLeft
        End If

        If targetHeight > 0 Then
            offsetTop = targetTop + (targetHeight - scaledHeight) / 2
        Else
            offsetTop = targetTop
        End If

        pastedShp.Left = offsetLeft
        pastedShp.Top = offsetTop

        pptShape.Delete

        If desiredZ > 0 Then
            If desiredZ > pptSlide.Shapes.Count Then desiredZ = pptSlide.Shapes.Count
            pastedShp.ZOrder msoSendToBack
            For i = 1 To desiredZ - 1
                pastedShp.ZOrder msoBringForward
            Next i
        End If
    Else
        With pastedShp
            .Left = (pptSlide.Master.Width - .Width) / 2
            .Top = (pptSlide.Master.Height - .Height) / 2
        End With
    End If

    pastedShp.Select
    pptApp.Activate
    Exit Sub

CleanFail:
    MsgBox "Error: " & Err.Description, vbCritical, "CopyPasteAsPictureToPPT"
End Sub


'====================================================
' Helper: Copy selection "As shown when printed"
'====================================================
Private Sub CopyCellPresentation(ByVal srcCell As Range, ByVal destCell As Range)
    On Error Resume Next

    destCell.NumberFormat = srcCell.NumberFormat

    With destCell.Font
        .Name = srcCell.Font.Name
        .Size = srcCell.Font.Size
        .Bold = srcCell.Font.Bold
        .Italic = srcCell.Font.Italic
        .Underline = srcCell.Font.Underline
        .Color = srcCell.Font.Color
        .Strikethrough = srcCell.Font.Strikethrough
    End With

    With destCell.Interior
        If srcCell.Interior.Pattern = xlPatternNone _
            Or srcCell.Interior.ColorIndex = xlColorIndexNone Then
            .Pattern = xlPatternNone
            .ColorIndex = xlColorIndexNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        Else
            .Pattern = srcCell.Interior.Pattern
            .PatternColorIndex = srcCell.Interior.PatternColorIndex
            .Color = srcCell.Interior.Color
            .TintAndShade = srcCell.Interior.TintAndShade
            .PatternTintAndShade = srcCell.Interior.PatternTintAndShade
        End If
    End With

    destCell.HorizontalAlignment = srcCell.HorizontalAlignment
    destCell.VerticalAlignment = srcCell.VerticalAlignment
    destCell.WrapText = srcCell.WrapText
    destCell.Orientation = srcCell.Orientation
    destCell.AddIndent = srcCell.AddIndent
    destCell.IndentLevel = srcCell.IndentLevel
    destCell.ShrinkToFit = srcCell.ShrinkToFit
    destCell.ReadingOrder = srcCell.ReadingOrder

    Dim borderId As Variant
    For Each borderId In Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom)
        With destCell.Borders(borderId)
            .LineStyle = srcCell.Borders(borderId).LineStyle
            If .LineStyle <> xlNone Then
                .Weight = srcCell.Borders(borderId).Weight
                .Color = srcCell.Borders(borderId).Color
            End If
        End With
    Next borderId

    On Error GoTo 0
End Sub


Private Function ResolveChartSelection(ByVal sel As Object, ByRef chartObj As Object, ByRef chart As Object) As Boolean
    Dim cur As Object
    Dim depth As Long
    Dim t As String

    On Error Resume Next
    If sel Is Nothing Then Exit Function

    Select Case TypeName(sel)
        Case "ChartObject"
            Set chartObj = sel
            Set chart = sel.Chart
            ResolveChartSelection = Not chart Is Nothing
            Exit Function
        Case "Chart"
            Set chart = sel
            ResolveChartSelection = True
            Exit Function
        Case "Shape"
            If CBool(sel.HasChart) Then
                Set chart = sel.Chart
                ResolveChartSelection = Not chart Is Nothing
                Exit Function
            End If
    End Select

    Set cur = sel
    For depth = 1 To 6
        If cur Is Nothing Then Exit For
        t = TypeName(cur)
        If t = "ChartObject" Then
            Set chartObj = cur
            Set chart = cur.Chart
            ResolveChartSelection = Not chart Is Nothing
            Exit Function
        ElseIf t = "Chart" Then
            Set chart = cur
            ResolveChartSelection = True
            Exit Function
        End If
        Set cur = cur.Parent
    Next depth
End Function

Private Function CopySelectionAsPicturePrintSafe(sel As Object) As Boolean
    Dim t As String
    Dim resolvedChart As Object
    Dim resolvedChartObj As Object
    t = TypeName(sel)

    On Error Resume Next

    If ResolveChartSelection(sel, resolvedChartObj, resolvedChart) Then
        If Not resolvedChartObj Is Nothing Then
            resolvedChartObj.Parent.Activate
            resolvedChartObj.Activate
        ElseIf TypeName(resolvedChart.Parent) = "ChartObject" Then
            resolvedChart.Parent.Parent.Activate
            resolvedChart.Parent.Activate
        Else
            resolvedChart.Parent.Activate
        End If

        Err.Clear
        resolvedChart.CopyPicture format:=xlPicture, appearance:=xlScreen
        If Err.Number <> 0 Then
            Err.Clear: resolvedChart.CopyPicture format:=xlBitmap, appearance:=xlScreen
            If Err.Number <> 0 Then
                Err.Clear: resolvedChart.CopyPicture format:=xlPicture, appearance:=xlPrinter
                If Err.Number <> 0 Then
                    Err.Clear: resolvedChart.CopyPicture format:=xlBitmap, appearance:=xlPrinter
                    If Err.Number <> 0 Then
                        Err.Clear: resolvedChart.Copy
                    End If
                End If
            End If
        End If
        GoTo DoneCopy
    End If

    Select Case t
        Case "Range"
            sel.Parent.Activate
            sel.Worksheet.Activate
            sel.Select
            Dim r As Range
            Set r = sel
            If r.Areas.Count > 1 Then Set r = r.Areas(1)

            Err.Clear
            r.CopyPicture appearance:=xlPrinter, format:=xlPicture
            If Err.Number <> 0 Then
                Err.Clear: r.CopyPicture appearance:=xlPrinter, format:=xlBitmap
                If Err.Number <> 0 Then
                    Err.Clear: r.CopyPicture appearance:=xlScreen, format:=xlPicture
                    If Err.Number <> 0 Then
                        Err.Clear: r.Copy
                    End If
                End If
            End If

        Case "ChartObject"
            sel.Parent.Activate
            sel.Activate
            Err.Clear
            sel.Chart.CopyPicture format:=xlPicture, appearance:=xlScreen
            If Err.Number <> 0 Then
                Err.Clear: sel.Chart.CopyPicture format:=xlBitmap, appearance:=xlScreen
                If Err.Number <> 0 Then
                    Err.Clear: sel.Chart.CopyPicture format:=xlPicture, appearance:=xlPrinter
                    If Err.Number <> 0 Then
                        Err.Clear: sel.Chart.CopyPicture format:=xlBitmap, appearance:=xlPrinter
                        If Err.Number <> 0 Then
                            Err.Clear: sel.Chart.Copy
                        End If
                    End If
                End If
            End If

        Case "Chart"
            sel.Parent.Activate
            sel.Activate
            Err.Clear
            sel.CopyPicture format:=xlPicture, appearance:=xlScreen
            If Err.Number <> 0 Then
                Err.Clear: sel.CopyPicture format:=xlBitmap, appearance:=xlScreen
                If Err.Number <> 0 Then
                    Err.Clear: sel.CopyPicture format:=xlPicture, appearance:=xlPrinter
                    If Err.Number <> 0 Then
                        Err.Clear: sel.CopyPicture format:=xlBitmap, appearance:=xlPrinter
                        If Err.Number <> 0 Then
                            Err.Clear: sel.Copy
                        End If
                    End If
                End If
            End If

        Case "Shape"
            Dim hasChart As Boolean
            hasChart = False
            On Error Resume Next
            hasChart = CBool(sel.HasChart)
            On Error GoTo 0
            sel.Parent.Parent.Activate
            Err.Clear
            If hasChart Then
                sel.Chart.CopyPicture format:=xlPicture, appearance:=xlScreen
                If Err.Number <> 0 Then
                    Err.Clear: sel.Chart.CopyPicture format:=xlBitmap, appearance:=xlScreen
                    If Err.Number <> 0 Then
                        Err.Clear: sel.Chart.CopyPicture format:=xlPicture, appearance:=xlPrinter
                        If Err.Number <> 0 Then
                            Err.Clear: sel.Chart.CopyPicture format:=xlBitmap, appearance:=xlPrinter
                            If Err.Number <> 0 Then Err.Clear: sel.Chart.Copy
                        End If
                    End If
                End If
            Else
                sel.Select
                Err.Clear: sel.CopyPicture appearance:=xlPrinter, format:=xlPicture
                If Err.Number <> 0 Then
                    Err.Clear: sel.CopyPicture appearance:=xlPrinter, format:=xlBitmap
                    If Err.Number <> 0 Then Err.Clear: sel.Copy
                End If
            End If

        Case Else
            Err.Clear
            On Error Resume Next
            CallByName sel, "CopyPicture", VbMethod, xlPrinter, xlPicture
            If Err.Number <> 0 Then
                Err.Clear: CallByName sel, "CopyPicture", VbMethod, xlPrinter, xlBitmap
                If Err.Number <> 0 Then
                    Err.Clear: CallByName sel, "Copy", VbMethod
                End If
            End If
    End Select

DoneCopy:
    Call WaitForClipboardReady(600)
    CopySelectionAsPicturePrintSafe = HasClipboardContent()
    On Error GoTo 0
End Function

Public Sub UnlockWorkbookAndSheets()
    Dim targetWorkbook As Workbook
    Dim ws As Worksheet
    Dim unlockedSheets As Long
    Dim failedList As String
    Dim foundPassword As String
    Dim workbookUnlocked As Boolean

    Set targetWorkbook = ThisWorkbook

    If IsTargetProtected(targetWorkbook) Then
        workbookUnlocked = AttemptHashBypass(targetWorkbook, foundPassword)
        If Not workbookUnlocked Then
            failedList = "Workbook structure/windows" & vbNewLine
        End If
    End If

    For Each ws In targetWorkbook.Worksheets
        If IsTargetProtected(ws) Then
            If AttemptHashBypass(ws, foundPassword) Then
                unlockedSheets = unlockedSheets + 1
            Else
                failedList = failedList & ws.Name & vbNewLine
            End If
        End If
    Next ws

    If Len(failedList) = 0 Then
        MsgBox "Finished: unlocked " & unlockedSheets & " protected sheet(s)." & IIf(workbookUnlocked, _
            vbNewLine & "Workbook structure/windows also unlocked.", vbNullString), vbInformation
    Else
        MsgBox "Finished but the following items are still protected:" & vbNewLine & failedList, vbExclamation
    End If
End Sub

Private Function AttemptHashBypass(ByVal target As Object, ByRef recoveredPassword As String) As Boolean
    Dim alphabet(0 To 1) As Integer
    Dim positions(1 To 12) As Integer
    Dim candidate As String
    Dim idx As Integer
    Dim done As Boolean

    alphabet(0) = 65
    alphabet(1) = 66

    For idx = LBound(positions) To UBound(positions)
        positions(idx) = alphabet(0)
    Next idx

    AttemptHashBypass = TryUnprotectTarget(target, "", recoveredPassword)
    If AttemptHashBypass Then Exit Function

    Do While Not done
        candidate = BuildCandidate(positions)
        If TryUnprotectTarget(target, candidate, recoveredPassword) Then
            AttemptHashBypass = True
            Exit Function
        End If

        done = Not IncrementPositions(positions, alphabet)
    Loop
End Function

Private Function TryUnprotectTarget(ByVal target As Object, ByVal passwordAttempt As String, ByRef recoveredPassword As String) As Boolean
    Dim typeNameValue As String
    Dim stillProtected As Boolean

    typeNameValue = TypeName(target)

    On Error Resume Next
    target.Unprotect passwordAttempt

    Select Case typeNameValue
        Case "Worksheet"
            stillProtected = target.ProtectContents Or target.ProtectDrawingObjects Or target.ProtectScenarios
        Case "Workbook"
            stillProtected = target.ProtectStructure Or target.ProtectWindows
        Case Else
            stillProtected = target.ProtectContents
    End Select
    On Error GoTo 0

    If Not stillProtected Then
        recoveredPassword = passwordAttempt
        TryUnprotectTarget = True
    End If
End Function

Private Function BuildCandidate(ByRef positions() As Integer) As String
    Dim idx As Integer

    BuildCandidate = ""
    For idx = LBound(positions) To UBound(positions)
        BuildCandidate = BuildCandidate & Chr$(positions(idx))
    Next idx
End Function

Private Function IncrementPositions(ByRef positions() As Integer, ByRef alphabet() As Integer) As Boolean
    Dim idx As Integer

    For idx = UBound(positions) To LBound(positions) Step -1
        If positions(idx) = alphabet(UBound(alphabet)) Then
            positions(idx) = alphabet(LBound(alphabet))
        Else
            positions(idx) = alphabet(LBound(alphabet)) + 1
            IncrementPositions = True
            Exit Function
        End If
    Next idx
End Function

Private Function IsTargetProtected(ByVal target As Object) As Boolean
    Dim typeNameValue As String

    typeNameValue = TypeName(target)
    Select Case typeNameValue
        Case "Worksheet"
            IsTargetProtected = target.ProtectContents Or target.ProtectDrawingObjects Or target.ProtectScenarios
        Case "Workbook"
            IsTargetProtected = target.ProtectStructure Or target.ProtectWindows
        Case Else
            On Error Resume Next
            IsTargetProtected = target.ProtectContents
            On Error GoTo 0
    End Select
End Function

'===========================================
' Quick font helpers
'===========================================
Public Function SetFontGaramond(Optional ByVal g As String) As Boolean
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    On Error GoTo CleanFail

    If TypeName(Selection) <> "Range" Then GoTo CleanExit

    Selection.Font.Name = "Garamond"

CleanExit:
    SetFontGaramond = False
    Exit Function

CleanFail:
    Call ErrorHandler("SetFontGaramond")
    Resume CleanExit
End Function

Public Function CmdInsertNumbers(Optional ByVal g As String) As Boolean
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    On Error GoTo CleanFail

    Call StopVisualMode

    If TypeName(Selection) <> "Range" Then GoTo CleanExit

    Dim startCell As Range
    Set startCell = ActiveCell
    If startCell Is Nothing Then GoTo CleanExit

    Dim target As Range
    Set target = startCell.Resize(1, 11)

    Dim values(1 To 1, 1 To 11) As Variant
    Dim i As Long
    For i = 1 To 11
        values(1, i) = i - 1
    Next i

    target.Value2 = values
    With target
        .Font.Name = "Garamond"
        .Font.Italic = True
        .Font.Color = vbBlack
        .NumberFormat = "#,##0_);(#,##0);--_)"
    End With

CleanExit:
    CmdInsertNumbers = False
    Exit Function

CleanFail:
    Call ErrorHandler("CmdInsertNumbers")
    Resume CleanExit
End Function


Function PasteExact(Optional ByVal g As String) As Boolean
    Dim engine As Object
    On Error GoTo CleanFail
    Call RepeatRegister("PasteExact")
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.PasteExact
CleanExit:
    PasteExact = False
    Exit Function
CleanFail:
    Call ErrorHandler("PasteExact")
    Resume CleanExit
End Function

Function PasteCondensed(Optional ByVal g As String) As Boolean
    Dim engine As Object
    On Error GoTo CleanFail
    Call RepeatRegister("PasteCondensed")
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.PasteCondensed
CleanExit:
    PasteCondensed = False
    Exit Function
CleanFail:
    Call ErrorHandler("PasteCondensed")
    Resume CleanExit
End Function

Function HighlightYellow(Optional ByVal g As String) As Boolean
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.HighlightSelectionYellow
CleanExit:
    HighlightYellow = False
    Exit Function
CleanFail:
    Call ErrorHandler("HighlightYellow")
    Resume CleanExit
End Function

Private Function CopySelectionAsPicture(ByVal sel As Variant) As Boolean
    Dim target As Object
    Set target = ResolveCopyTarget(sel)
    If target Is Nothing Then Exit Function

    ' Dedicated, reliable path for ranges (mirrors Alt, H, C, P, P)
    If TypeName(target) = "Range" Then
        CopySelectionAsPicture = CopyRangePictureReliable(target)
        Exit Function
    End If

    ' Non-range (charts/shapes): try direct CopyPicture up the parent chain
    CopySelectionAsPicture = TryCopyAsPicture(target, xlPrinter, xlPicture)
    If CopySelectionAsPicture Then Exit Function
    CopySelectionAsPicture = TryCopyAsPicture(target, xlScreen, xlPicture)
    If CopySelectionAsPicture Then Exit Function
    If Not target Is Nothing Then
        Dim parentObj As Object
        On Error Resume Next
        Set parentObj = target
        Do While Not parentObj Is Nothing
            Set parentObj = parentObj.Parent
            If parentObj Is Nothing Then Exit Do
            CopySelectionAsPicture = TryCopyAsPicture(parentObj, xlPrinter, xlPicture)
            If CopySelectionAsPicture Then Exit Do
            CopySelectionAsPicture = TryCopyAsPicture(parentObj, xlScreen, xlPicture)
            If CopySelectionAsPicture Then Exit Do
        Loop
        On Error GoTo 0
    End If
End Function

Private Function CopyRangePictureReliable(ByVal rng As Range) As Boolean
    Dim ws As Worksheet
    Dim area As Range
    Dim attempt As Long
    Dim tmp As Object
    Dim shp As Object
    Dim w As Double, h As Double
    Dim okSize As Boolean
    Dim originalWindow As Window
    Dim originalSheet As Worksheet

    On Error GoTo CleanFail

    If rng Is Nothing Then GoTo CleanExit
    Set ws = rng.Worksheet
    Set area = rng
    If area.Areas.Count > 1 Then Set area = area.Areas(1)

    On Error Resume Next
    Set originalWindow = Application.ActiveWindow
    If Not originalWindow Is Nothing Then
        Set originalSheet = originalWindow.ActiveSheet
    End If
    On Error GoTo CleanFail

    For attempt = 1 To 4
        Application.CutCopyMode = False
        ' Ensure the sheet is active to allow Paste and shape operations
        On Error Resume Next
        If Not ws Is ActiveSheet Then ws.Activate
        On Error GoTo CleanFail
        ' Bring focus near the range
        SafeSelectRange area.Cells(1, 1)
        Err.Clear
        Select Case attempt
            Case 1: area.CopyPicture appearance:=xlScreen, format:=xlPicture
            Case 2: area.CopyPicture appearance:=xlScreen, format:=xlBitmap
            Case 3: area.CopyPicture appearance:=xlPrinter, format:=xlPicture
            Case Else: area.CopyPicture appearance:=xlPrinter, format:=xlBitmap
        End Select
        Call WaitForClipboardReady(400)

        On Error Resume Next
        Set tmp = ws.Pictures.Paste
        If tmp Is Nothing Then Set tmp = ws.Shapes.Paste
        On Error GoTo CleanFail

        If Not tmp Is Nothing Then
            If TypeName(tmp) = "ShapeRange" Then
                Set shp = tmp(1)
            Else
                Set shp = tmp
            End If

            w = area.Width: h = area.Height
            okSize = (Abs(shp.Width - w) <= (w * 0.1 + 2)) And (Abs(shp.Height - h) <= (h * 0.1 + 2))

            If okSize Then
                On Error Resume Next
                shp.Copy
                On Error GoTo 0
                Call WaitForClipboardReady(500)
                On Error Resume Next
                shp.Delete
                On Error GoTo 0
                CopyRangePictureReliable = True
                GoTo CleanExit
            Else
                On Error Resume Next
                shp.Delete
                On Error GoTo 0
            End If
        End If
        DoEvents
    Next attempt

CleanExit:
    On Error Resume Next
    If Not originalWindow Is Nothing And originalWindow.Visible Then originalWindow.Activate
    If Not originalSheet Is Nothing Then originalSheet.Activate
    Set originalSheet = Nothing
    Set originalWindow = Nothing
    On Error GoTo 0
    Exit Function

CleanFail:
    On Error Resume Next
    If Not shp Is Nothing Then shp.Delete
    On Error GoTo 0
    GoTo CleanExit
End Function

Private Function TryCopyAsPicture(ByVal target As Variant, _
                                  ByVal appearance As XlPictureAppearance, _
                                  ByVal format As XlCopyPictureFormat) As Boolean
    On Error Resume Next
    target.CopyPicture appearance:=appearance, format:=format
    TryCopyAsPicture = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Sub SmartFillRight()
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.SmartFillRight
CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("SmartFillRight")
    Resume CleanExit
End Sub

' ===== Apply finance formatting across contiguous right cells =====
Sub SmartFormatRight()
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.SmartFormatRight
CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("SmartFormatRight")
    Resume CleanExit
End Sub

' ===== Outline selection with navy box and corner markers =====
Sub OutlineSelectionHighlight()
    If gMacroSafeMode Then
        Call SetStatusBarTemporarily("Fast mode on: outline highlight disabled.", 2000)
        Exit Sub
    End If

    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.OutlineSelectionHighlight
CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("OutlineSelectionHighlight")
    Resume CleanExit
End Sub
' ===== Helper: nearest row scan ÃƒÆ’Ã†â�™ÃƒÂ¢Ã¢â�šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ�šÃ‚Â±maxOffset =====
Function ClearUnnecessaryFormatting(Optional ByVal g As String) As Boolean
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.ClearUnnecessaryFormatting
CleanExit:
    ClearUnnecessaryFormatting = False
    Exit Function
CleanFail:
    Call ErrorHandler("ClearUnnecessaryFormatting")
    Resume CleanExit
End Function

Public Sub TrimConditionalFormatting()
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.TrimConditionalFormatting
CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("TrimConditionalFormatting")
    Resume CleanExit
End Sub

Private Function ShouldRestartToFlushCaches(ByVal Wb As Workbook) As Boolean
    On Error GoTo Fail
    If Wb Is Nothing Then Exit Function

    Dim wbPath As String
    wbPath = Wb.FullName
    If Len(wbPath) = 0 Then Exit Function

    Dim response As VbMsgBoxResult
    response = MsgBox("Optimization complete." & vbCrLf & _
                      "Reopen Excel with this workbook to flush caches?" & vbCrLf & vbCrLf & _
                      "This will save the workbook and close the current Excel session.", _
                      vbQuestion + vbYesNo, "Reopen workbook")
    If response <> vbYes Then Exit Function

    ShouldRestartToFlushCaches = RestartWorkbookFresh(Wb)
    Exit Function

Fail:
    ShouldRestartToFlushCaches = False
End Function

Private Function RestartWorkbookFresh(ByVal Wb As Workbook) As Boolean
    On Error GoTo Fail
    If Wb Is Nothing Then Exit Function

    Dim wbPath As String
    wbPath = Wb.FullName
    If Len(wbPath) = 0 Then
        MsgBox "Please save the workbook before running the refresh.", vbExclamation
        Exit Function
    End If

    Dim appPath As String
    appPath = Application.Path & "\EXCEL.EXE"

    Application.StatusBar = "Reopening workbook..."
    Wb.Save

    Shell """" & appPath & """ """ & wbPath & """", vbNormalFocus
    RestartWorkbookFresh = True
    Application.Quit
    Exit Function

Fail:
    Application.StatusBar = False
    MsgBox "Could not restart Excel automatically. Please reopen the workbook manually.", vbExclamation
    RestartWorkbookFresh = False
End Function

Private Sub CleanupBrokenNames(ByVal Wb As Workbook)
    On Error GoTo CleanExit
    Dim nm As Name
    Dim ws As Worksheet
    Dim refersTo As String

    For Each nm In Wb.names
        refersTo = ""
        On Error Resume Next
        refersTo = nm.refersTo
        If Err.Number <> 0 Then
            Err.Clear
            nm.Delete
        ElseIf InStr(1, refersTo, "#REF!", vbTextCompare) > 0 Then
            nm.Delete
        End If
        On Error GoTo 0
    Next nm

    For Each ws In Wb.Worksheets
        For Each nm In ws.names
            refersTo = ""
            On Error Resume Next
            refersTo = nm.refersTo
            If Err.Number <> 0 Then
                Err.Clear
                nm.Delete
            ElseIf InStr(1, refersTo, "#REF!", vbTextCompare) > 0 Then
                nm.Delete
            End If
            On Error GoTo 0
        Next nm
    Next ws

CleanExit:
    On Error GoTo 0
End Sub


Private Function CellHasVisualMarker(ByVal cell As Range) As Boolean
    Const XL_COLOR_NONE As Long = -4142
    Dim hasVal As Boolean
    Dim hasFill As Boolean
    Dim hasBorder As Boolean
    Dim displayColorIndex As Variant

    On Error Resume Next
    If Not IsError(cell.value) Then
        hasVal = (Len(Trim$(CStr(cell.value))) > 0)
    End If

    displayColorIndex = cell.DisplayFormat.Interior.ColorIndex
    If Not IsError(displayColorIndex) Then
        hasFill = (displayColorIndex <> XL_COLOR_NONE)
    ElseIf Not IsError(cell.Interior.ColorIndex) Then
        hasFill = (cell.Interior.ColorIndex <> XL_COLOR_NONE)
    End If

    hasBorder = False
    If cell.DisplayFormat.Borders(xlEdgeLeft).lineStyle <> xlLineStyleNone _
       Or cell.DisplayFormat.Borders(xlEdgeRight).lineStyle <> xlLineStyleNone _
       Or cell.DisplayFormat.Borders(xlEdgeTop).lineStyle <> xlLineStyleNone _
       Or cell.DisplayFormat.Borders(xlEdgeBottom).lineStyle <> xlLineStyleNone _
       Or cell.DisplayFormat.Borders(xlInsideHorizontal).lineStyle <> xlLineStyleNone _
       Or cell.DisplayFormat.Borders(xlInsideVertical).lineStyle <> xlLineStyleNone Then
        hasBorder = True
    ElseIf cell.Borders(xlEdgeLeft).lineStyle <> xlNone _
        Or cell.Borders(xlEdgeRight).lineStyle <> xlNone _
        Or cell.Borders(xlEdgeTop).lineStyle <> xlNone _
        Or cell.Borders(xlEdgeBottom).lineStyle <> xlNone _
        Or cell.Borders(xlInsideHorizontal).lineStyle <> xlNone _
        Or cell.Borders(xlInsideVertical).lineStyle <> xlNone Then
        hasBorder = True
    End If
    On Error GoTo 0

    CellHasVisualMarker = hasVal Or hasFill Or hasBorder
End Function

Private Function GetRowLastMarker(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal startCol As Long, ByVal colRight As Long, ByVal cache As Object) As Long
    Dim key As String
    key = CStr(rowIndex) & "|" & CStr(startCol) & "|" & CStr(colRight)
    If Not cache Is Nothing Then
        If cache.Exists(key) Then
            GetRowLastMarker = cache(key)
            Exit Function
        End If
    End If

    Dim col As Long
    Dim lastMarker As Long

    lastMarker = startCol

    If colRight > ws.Columns.Count Then colRight = ws.Columns.Count
    For col = startCol + 1 To colRight
        If CellHasVisualMarker(ws.Cells(rowIndex, col)) Then
            lastMarker = col
        Else
            Exit For
        End If
    Next col

    If Not cache Is Nothing Then
        cache(key) = lastMarker
    End If

    GetRowLastMarker = lastMarker
End Function

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    On Error Resume Next
    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If Not lastCell Is Nothing Then GetLastUsedRow = lastCell.Row
End Function

Private Function GetLastUsedColumn(ByVal ws As Worksheet) As Long
    On Error Resume Next
    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If Not lastCell Is Nothing Then GetLastUsedColumn = lastCell.Column
End Function

Sub RefreshExcelCaches(Optional ByVal Wb As Workbook = Nothing)
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    On Error GoTo CleanExit

    If Wb Is Nothing Then
        Set Wb = ActiveWorkbook
        If Wb Is Nothing Then Exit Sub
    End If

    Application.StatusBar = "Refreshing workbook caches..."
    Application.CutCopyMode = False
    ClipboardSetCopyRange Nothing

    Dim ws As Worksheet
    For Each ws In Wb.Worksheets
        ws.DisplayPageBreaks = False
    Next ws

    Dim pc As PivotCache
    For Each pc In Wb.PivotCaches
        On Error Resume Next
        pc.MissingItemsLimit = xlMissingItemsNone
        pc.Refresh
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo CleanExit
    Next pc

    Wb.Application.CalculateFullRebuild

CleanExit:
    Application.StatusBar = False
End Sub


Sub SmartFillDown()
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.SmartFillDown
CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("SmartFillDown")
    Resume CleanExit
End Sub


Sub CenterAcrossSelection()
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    Dim sel As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection
    ' Apply "Center Across Selection" alignment
    With sel
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    End With
End Sub

Sub WrapInIFERROR()
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.CountLarge = 0 Then Exit Sub
    Dim sel As Range
    Set sel = Selection
    Dim formulas As Variant
    formulas = sel.formula
    Dim currFormula As String
    ' Single-cell selections return a scalar, not a 2-D array.
    If Not IsArray(formulas) Then
        If Not IsError(formulas) Then
            currFormula = CStr(formulas)
            If Len(currFormula) > 0 And Left$(currFormula, 1) = "=" Then
                sel.formula = "=IFERROR(" & Mid$(currFormula, 2) & ",0)"
            End If
        End If
        Exit Sub
    End If
    Dim r As Long, c As Long
    For r = LBound(formulas, 1) To UBound(formulas, 1)
        For c = LBound(formulas, 2) To UBound(formulas, 2)
            If Not IsError(formulas(r, c)) Then
                currFormula = CStr(formulas(r, c))
                If Len(currFormula) > 0 And Left$(currFormula, 1) = "=" Then
                    formulas(r, c) = "=IFERROR(" & Mid$(currFormula, 2) & ",0)"
                End If
            End If
        Next c
    Next r
    sel.formula = formulas
End Sub

Sub LockCellReference()
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.LockCellReference
CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("LockCellReference")
    Resume CleanExit
End Sub

Public Sub CycleFormatting()
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.CycleFormatting
CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("CycleFormatting")
    Resume CleanExit
End Sub

Private Sub WaitForClipboardReady(ByVal maxMillis As Long)
    Dim start As Double
    Dim elapsed As Double
    start = Timer
    Do
        If HasClipboardContent() Then Exit Sub
        DoEvents
        elapsed = (Timer - start)
        If elapsed < 0 Then elapsed = elapsed + 86400# ' handle midnight wrap
        If elapsed * 1000# >= maxMillis Then Exit Do
    Loop
End Sub

Private Function HasClipboardContent() As Boolean
    On Error Resume Next
    Dim fmts As Variant
    fmts = Application.ClipboardFormats
    On Error GoTo 0

    If IsEmpty(fmts) Then
        HasClipboardContent = False
    ElseIf IsArray(fmts) Then
        HasClipboardContent = (UBound(fmts) >= LBound(fmts))
    Else
        HasClipboardContent = (VarType(fmts) <> vbEmpty)
    End If
End Function

Private Function InsertRangePictureViaFile(ByVal rng As Range, ByVal pptSlide As Object) As Object
    On Error GoTo Fail
    If rng Is Nothing Then Exit Function

    Dim ws As Worksheet
    Dim ar As Range
    Dim co As ChartObject
    Dim tmp As String
    Dim shp As Object

    Set ws = rng.Worksheet
    Set ar = rng
    If ar.Areas.Count > 1 Then Set ar = ar.Areas(1)

    ' Create a temporary chart sized to the range and paste the range picture
    Set co = ws.ChartObjects.Add(Left:=ar.Left, Top:=ar.Top, Width:=ar.Width, Height:=ar.Height)
    On Error Resume Next
    co.ShapeRange.Line.Visible = msoFalse
    On Error GoTo Fail

    ar.CopyPicture appearance:=xlScreen, format:=xlPicture
    Call WaitForClipboardReady(300)
    co.Activate
    co.Chart.Paste

    ' Export chart as EMF and insert into PowerPoint
    tmp = Environ$("TEMP")
    If Len(tmp) = 0 Then tmp = ThisWorkbook.Path
    If Right$(tmp, 1) <> "\\" Then tmp = tmp & "\\"
    Randomize
    tmp = tmp & "xp_rng_" & format(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Rnd * 1000000)) & ".png"

    On Error Resume Next
    co.Chart.Export tmp, "PNG"
    On Error GoTo Fail

    co.Delete
    If Len(Dir$(tmp)) = 0 Then GoTo Fail

    Set shp = pptSlide.Shapes.AddPicture(FileName:=tmp, _
                                         LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                                         Left:=0, Top:=0, Width:=-1, Height:=-1)
    On Error Resume Next
    Kill tmp
    On Error GoTo 0
    Set InsertRangePictureViaFile = shp
    Exit Function

Fail:
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    Set InsertRangePictureViaFile = Nothing
End Function

Sub FlipSign()
    On Error GoTo CleanFail
    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub
    engine.FlipSign
CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("FlipSign")
    Resume CleanExit
End Sub

Sub ReverseSelectionOrder()
    On Error GoTo CleanFail
    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub
    engine.ReverseSelectionOrder
CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("ReverseSelectionOrder")
    Resume CleanExit
End Sub

Sub FormatChart_FG()
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.FormatChartFg
CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("FormatChart_FG")
    Resume CleanExit
End Sub

' ==================================================
'  LIGHTNING-FAST DATA LABEL MOVE HANDLER
' ==================================================



Private Function ResolveSelectedChartContainer(ByVal seed As Object) As Object
    Const MAX_PARENT_HOPS As Long = 20
    Dim current As Object
    Dim nextParent As Object
    Dim hopCount As Long
    Set current = seed

    On Error Resume Next
    Do While Not current Is Nothing And hopCount < MAX_PARENT_HOPS
        hopCount = hopCount + 1
        Set nextParent = Nothing

        Select Case TypeName(current)
            Case "ChartObject"
                Set ResolveSelectedChartContainer = current
                Exit Function
            Case "Shape"
                If current.HasChart Then
                    Set ResolveSelectedChartContainer = current
                    Exit Function
                End If
                Set nextParent = CallByName(current, "Parent", VbGet)
            Case "ShapeRange"
                If current.Count = 1 Then
                    Set current = current.Item(1)
                    hopCount = hopCount - 1
                    GoTo ContinueLoop
                Else
                    Exit Do
                End If
            Case "Chart"
                Set nextParent = current.Parent
                If TypeName(nextParent) = "ChartObject" Or TypeName(nextParent) = "Shape" Then
                    Set ResolveSelectedChartContainer = nextParent
                    Exit Function
                End If
            Case Else
                Set nextParent = CallByName(current, "Parent", VbGet)
        End Select

        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If

        If nextParent Is Nothing Then Exit Do
        If nextParent Is current Then Exit Do
        Set current = nextParent
ContinueLoop:
    Loop
    On Error GoTo 0
End Function

Private Function IsChartMoveSelection(ByVal target As Object) As Boolean
    If target Is Nothing Then Exit Function
    Dim t As String
    t = TypeName(target)
    Select Case t
        Case "ChartObject", "ChartArea", "PlotArea", _
             "Chart", "Series", "DataPoint", "Legend", _
             "LegendEntry", "LegendKey", "ChartTitle", _
             "Shape", "ShapeRange"
            IsChartMoveSelection = True
        Case Else
            IsChartMoveSelection = False
    End Select
End Function

' ==================================================
'  SMART MOVE SHORTCUTS
' ==================================================

Sub MoveLeftSmart()
    Dim steps As Long: steps = gVim.Count1: If steps < 1 Then steps = 1

    If gScrollLockMode Then
        ScrollLockScroll 0, -steps
        gVim.Count1 = 1
        Exit Sub
    End If


    Dim selType As String
    On Error Resume Next
    selType = TypeName(Selection)
    On Error GoTo 0
    If selType = "Range" Then
        Call MoveLeft
        gVim.Count1 = 1
        Exit Sub
    End If

    If TryMoveSelectedLabel(-DATA_LABEL_STEP * steps, 0#) Then
        gVim.Count1 = 1
        Exit Sub
    End If

    If TryMoveSelectedChart(-CHART_MOVE_STEP * steps, 0#) Then
        gVim.Count1 = 1
        Exit Sub
    End If

    Call MoveLeft
    gVim.Count1 = 1
End Sub

Sub MoveRightSmart()
    Dim steps As Long: steps = gVim.Count1: If steps < 1 Then steps = 1

    If gScrollLockMode Then
        ScrollLockScroll 0, steps
        gVim.Count1 = 1
        Exit Sub
    End If


    Dim selType As String
    On Error Resume Next
    selType = TypeName(Selection)
    On Error GoTo 0
    If selType = "Range" Then
        Call MoveRight
        gVim.Count1 = 1
        Exit Sub
    End If

    If TryMoveSelectedLabel(DATA_LABEL_STEP * steps, 0#) Then
        gVim.Count1 = 1
        Exit Sub
    End If

    If TryMoveSelectedChart(CHART_MOVE_STEP * steps, 0#) Then
        gVim.Count1 = 1
        Exit Sub
    End If

    Call MoveRight
    gVim.Count1 = 1
End Sub

Sub MoveUpSmart()
    Dim steps As Long: steps = gVim.Count1: If steps < 1 Then steps = 1

    If gScrollLockMode Then
        ScrollLockScroll -steps, 0
        gVim.Count1 = 1
        Exit Sub
    End If


    Dim selType As String
    On Error Resume Next
    selType = TypeName(Selection)
    On Error GoTo 0
    If selType = "Range" Then
        Call MoveUp
        gVim.Count1 = 1
        Exit Sub
    End If

    If TryMoveSelectedLabel(0#, -DATA_LABEL_STEP * steps) Then
        gVim.Count1 = 1
        Exit Sub
    End If

    If TryMoveSelectedChart(0#, -CHART_MOVE_STEP * steps) Then
        gVim.Count1 = 1
        Exit Sub
    End If

    Call MoveUp
    gVim.Count1 = 1
End Sub

Sub MoveDownSmart()
    Dim steps As Long: steps = gVim.Count1: If steps < 1 Then steps = 1

    If gScrollLockMode Then
        ScrollLockScroll steps, 0
        gVim.Count1 = 1
        Exit Sub
    End If


    Dim selType As String
    On Error Resume Next
    selType = TypeName(Selection)
    On Error GoTo 0
    If selType = "Range" Then
        Call MoveDown
        gVim.Count1 = 1
        Exit Sub
    End If

    If TryMoveSelectedLabel(0#, DATA_LABEL_STEP * steps) Then
        gVim.Count1 = 1
        Exit Sub
    End If

    If TryMoveSelectedChart(0#, CHART_MOVE_STEP * steps) Then
        gVim.Count1 = 1
        Exit Sub
    End If

    Call MoveDown
    gVim.Count1 = 1
End Sub

Private Function TryMoveSelectedLabel(ByVal dx As Double, ByVal dy As Double) As Boolean
    Dim lbl As Excel.DataLabel
    Dim pt As Excel.Point
    Dim seriesObj As Excel.Series
    Dim t As String

    If gScrollLockMode Then Exit Function

    On Error Resume Next
    t = TypeName(Selection)
    On Error GoTo 0

    Select Case t
        Case "DataLabel"
            On Error Resume Next
            Set lbl = Selection
            lbl.Position = xlLabelPositionCustom
            lbl.Left = lbl.Left + dx
            lbl.Top = lbl.Top + dy
            On Error GoTo 0
            TryMoveSelectedLabel = True
        Case "DataLabels"
            On Error Resume Next
            For Each lbl In Selection
                lbl.Position = xlLabelPositionCustom
                lbl.Left = lbl.Left + dx
                lbl.Top = lbl.Top + dy
            Next lbl
            On Error GoTo 0
            TryMoveSelectedLabel = True
        Case "Point", "DataPoint"
            On Error Resume Next
            Set pt = Selection
            If pt.HasDataLabel Then
                Set lbl = pt.DataLabel
                lbl.Position = xlLabelPositionCustom
                lbl.Left = lbl.Left + dx
                lbl.Top = lbl.Top + dy
                TryMoveSelectedLabel = True
            End If
            On Error GoTo 0
        Case "Series"
            On Error Resume Next
            Set seriesObj = Selection
            If Not seriesObj Is Nothing Then
                For Each pt In seriesObj.Points
                    If pt.HasDataLabel Then
                        Set lbl = pt.DataLabel
                        lbl.Position = xlLabelPositionCustom
                        lbl.Left = lbl.Left + dx
                        lbl.Top = lbl.Top + dy
                    End If
                Next pt
                TryMoveSelectedLabel = True
            End If
            On Error GoTo 0
        Case Else
            TryMoveSelectedLabel = False
    End Select
End Function

Private Function TryMoveSelectedChart(ByVal dx As Double, ByVal dy As Double) As Boolean
    If gScrollLockMode Then Exit Function

    Dim target As Object
    Set target = ResolveChartContainer(Selection)

    If target Is Nothing Then
        Dim parentObj As Object
        On Error Resume Next
        Set parentObj = Selection.Parent
        On Error GoTo 0
        Set target = ResolveChartContainer(parentObj)
    End If

    If target Is Nothing Then Exit Function

    On Error Resume Next
    target.Left = target.Left + dx
    target.Top = target.Top + dy
    If Err.Number = 0 Then
        TryMoveSelectedChart = True
    End If
    On Error GoTo 0
End Function

Private Function ResolveChartContainer(ByVal candidate As Object) As Object
    Dim current As Object
    Dim depth As Long

    On Error Resume Next
    Set current = candidate
    On Error GoTo 0

    For depth = 1 To 8
        If current Is Nothing Then Exit For

        Dim typeNameStr As String
        On Error Resume Next
        typeNameStr = TypeName(current)
        On Error GoTo 0

        Select Case typeNameStr
            Case "ChartObject", "ChartArea", "PlotArea", "Legend", "ChartTitle"
                Set ResolveChartContainer = current
                Exit Function
            Case "Chart"
                On Error Resume Next
                Set ResolveChartContainer = current.Parent
                On Error GoTo 0
                Exit Function
            Case "Shape"
                On Error Resume Next
                If current.HasChart Then
                    Set ResolveChartContainer = current
                    Exit Function
                End If
                On Error GoTo 0
            Case "ShapeRange"
                On Error Resume Next
                If current.Count = 1 Then
                    Set current = current.Item(1)
                    GoTo ContinueLoop
                End If
                On Error GoTo 0
        End Select

        On Error Resume Next
        Set current = current.Parent
        On Error GoTo 0
ContinueLoop:
    Next depth
End Function

Function SelectNearestChart(Optional ByVal g As String) As Boolean
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.SelectNearestChart
CleanExit:
    SelectNearestChart = False
    Exit Function
CleanFail:
    Call ErrorHandler("SelectNearestChart")
    Resume CleanExit
End Function

Private Sub EnsureChartElementSelection(ByVal chartContainer As Object)
    Dim ch As Chart
    Set ch = ResolveChartFromContainer(chartContainer)
    If ch Is Nothing Then Exit Sub

    On Error Resume Next
    ch.ChartArea.Select
    If Err.Number <> 0 Then
        Err.Clear
        ch.Parent.Select
    End If
    On Error GoTo 0
End Sub

Private Function ResolveChartFromContainer(ByVal container As Object) As Chart
    If container Is Nothing Then Exit Function
    On Error Resume Next
    Select Case TypeName(container)
        Case "ChartObject"
            Set ResolveChartFromContainer = container.Chart
        Case "Shape"
            If container.HasChart Then Set ResolveChartFromContainer = container.Chart
        Case "Chart"
            Set ResolveChartFromContainer = container
        Case "ChartArea"
            Set ResolveChartFromContainer = container.Parent
    End Select
    On Error GoTo 0
End Function

Private Sub RunDecimalCommand(ByVal controlId As Long)
    Dim uiGuard As ExcelUiGuard
    Dim ctrl As CommandBarControl
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set uiGuard = SuppressExcelUi(True)
    ' Find the ribbon command button: Type = msoControlButton
    Set ctrl = Application.CommandBars.FindControl(Type:=msoControlButton, ID:=controlId)
    If Not ctrl Is Nothing Then ctrl.Execute
CleanExit:
    Set ctrl = Nothing
    Set uiGuard = Nothing
End Sub

Public Sub IncreaseDecimalPlaces()
    RunDecimalCommand 398   ' built-in Increase Decimal button
End Sub

Public Sub DecreaseDecimalPlaces()
    RunDecimalCommand 399   ' built-in Decrease Decimal button
End Sub

Sub WrapFormulaWithCircCheck()
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    engine.WrapFormulaWithCircCheck
CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("WrapFormulaWithCircCheck")
    Resume CleanExit
End Sub









Private Function ResolveFallbackCopyTarget(ByVal sel As Variant) As Object
    Dim probe As Object

    On Error Resume Next
    If VarType(sel) = vbObject Then
        Set probe = sel
    End If
    On Error GoTo 0

    Do While Not probe Is Nothing
        Select Case TypeName(probe)
            Case "Range", "ChartObject", "Chart", "Picture", "Shape"
                Set ResolveFallbackCopyTarget = probe
                Exit Function
        End Select
        On Error Resume Next
        Set probe = probe.Parent
        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If
        On Error GoTo 0
    Loop

    On Error Resume Next
    If Not ActiveChart Is Nothing Then
        Set ResolveFallbackCopyTarget = ActiveChart
        Exit Function
    End If
    If TypeName(Selection) = "Range" Then
        Set ResolveFallbackCopyTarget = Selection
    End If
    On Error GoTo 0
End Function

Private Function TryGetSelectionCenter(ByRef centerX As Double, ByRef centerY As Double) As Boolean
    Dim rng As Range
    On Error Resume Next
    If TypeName(Selection) = "Range" Then
        Set rng = Selection
    End If
    On Error GoTo 0

    If rng Is Nothing Then
        On Error Resume Next
        Set rng = ActiveCell
        On Error GoTo 0
    End If

    If Not rng Is Nothing Then
        centerX = rng.Left + rng.Width / 2
        centerY = rng.Top + rng.Height / 2
        TryGetSelectionCenter = True
        Exit Function
    End If

    Dim win As Window
    On Error Resume Next
    Set win = ActiveWindow
    If Not win Is Nothing Then
        Dim vis As Range
        Set vis = win.VisibleRange
        If Not vis Is Nothing Then
            centerX = vis.Left + vis.Width / 2
            centerY = vis.Top + vis.Height / 2
            TryGetSelectionCenter = True
        End If
    End If
    On Error GoTo 0
End Function

Private Function DistanceSquared(ByVal x1 As Double, ByVal y1 As Double, _
                                 ByVal x2 As Double, ByVal y2 As Double) As Double
    DistanceSquared = (x1 - x2) ^ 2 + (y1 - y2) ^ 2
End Function

Private Function ChartObjectIsVisible(ByVal cbo As ChartObject) As Boolean
    On Error Resume Next
    If cbo Is Nothing Then Exit Function
    ChartObjectIsVisible = (CBool(cbo.Visible) And cbo.Width > 0 And cbo.Height > 0)
    On Error GoTo 0
End Function

Private Function ShapeHasVisibleChart(ByVal shp As shape) As Boolean
    On Error Resume Next
    If shp Is Nothing Then Exit Function
    If shp.Visible = msoFalse Then Exit Function
    ShapeHasVisibleChart = shp.hasChart
    On Error GoTo 0
End Function

Private Function ResolveOwningWorkbook(ByVal target As Object) As Workbook
    Dim current As Object
    Set current = target
    On Error Resume Next
    Do While Not current Is Nothing
        If TypeName(current) = "Workbook" Then
            Set ResolveOwningWorkbook = current
            Exit Function
        End If
        Set current = CallByName(current, "Parent", VbGet)
        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If
    Loop
    On Error GoTo 0
End Function

Private Sub ActivateChartContainer(ByVal target As Object)
    On Error Resume Next
    Dim owningWorkbook As Workbook
    Set owningWorkbook = ResolveOwningWorkbook(target)
    If owningWorkbook Is Nothing Then Exit Sub
    If Not IsActiveWorkbookRef(owningWorkbook) Then Exit Sub

    Select Case TypeName(target)
        Case "ChartObject"
            SafeActivateWorksheet target.Parent
            target.Activate
        Case "Shape"
            SafeActivateWorksheet target.Parent
            If target.hasChart Then
                target.Chart.Parent.Activate
            Else
                target.Select
            End If
        Case "Chart"
            Dim parentObj As Object
            Set parentObj = target.Parent
            If TypeName(parentObj) = "ChartObject" Then
                SafeActivateWorksheet parentObj.Parent
                parentObj.Activate
            Else
                target.Parent.Activate
            End If
        Case Else
            CallByName target, "Activate", VbMethod
    End Select
    On Error GoTo 0
End Sub

Private Function IsAltKeyDown() As Boolean
    ' Poll the Windows keyboard state so we can mirror Alt+Arrow behaviour even inside dialogs.
    IsAltKeyDown = ((GetAsyncKeyState(AltLeft_) And &H8000) <> 0) _
                   Or ((GetAsyncKeyState(AltRight_) And &H8000) <> 0)
End Function

'===========================================
' Workbook utilities
'===========================================
Sub Econs_Output_PPT_V2()
    Const layoutName As String = "content no text"
    Const targetSheetName As String = "Inputs"

    Dim originalCalc As XlCalculation
    Dim originalEvents As Boolean
    Dim originalScreen As Boolean

    originalCalc = Application.Calculation
    originalEvents = Application.EnableEvents
    originalScreen = Application.ScreenUpdating

    On Error GoTo CleanFail
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim pptApp As Object, pptPres As Object
    Dim slide As Object, customLayout As Object
    Dim wb As Workbook, wsInputs As Worksheet, pic As Object
    Dim slideWidth As Single
    Dim cases As Variant, c As Variant
    Dim userInput As String
    Dim i As Long, key As Variant
    Dim rng As Range
    Dim ws As Worksheet
    Dim restartInput As Boolean
    Dim caseCell As Range
    Static wbOutputMap As Object
    Static wbOutputOwner As String

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No active workbook found. Please open your Excel file and try again.", vbExclamation
        GoTo CleanExit
    End If

    Set wsInputs = Nothing
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, targetSheetName, vbTextCompare) = 0 Then
            Set wsInputs = ws
            Exit For
        End If
    Next ws
    If wsInputs Is Nothing Then
        MsgBox "Worksheet '" & targetSheetName & "' was not found in " & wb.Name & ".", vbExclamation
        GoTo CleanExit
    End If

    Set caseCell = Nothing
    On Error Resume Next
    Set caseCell = wsInputs.Range("case1")
    If caseCell Is Nothing Then
        Set caseCell = wb.Names("case1").RefersToRange
    End If
    On Error GoTo 0
    If caseCell Is Nothing Then
        MsgBox "Named cell or range 'case1' was not found on the Inputs sheet.", vbExclamation
        GoTo CleanExit
    End If

    If wbOutputMap Is Nothing Then Set wbOutputMap = CreateObject("Scripting.Dictionary")
    If wbOutputOwner <> wb.FullName Then
        wbOutputMap.RemoveAll
        wbOutputOwner = wb.FullName
    End If

    On Error Resume Next
    Set pptApp = GetObject(Class:="PowerPoint.Application")
    If pptApp Is Nothing Then Set pptApp = CreateObject(Class:="PowerPoint.Application")
    On Error GoTo CleanFail
    If pptApp Is Nothing Then
        MsgBox "Unable to start PowerPoint.", vbExclamation
        GoTo CleanExit
    End If
    pptApp.Visible = True

    If pptApp.Presentations.Count = 0 Then
        MsgBox "No PowerPoint presentations are open. Please open one and try again.", vbExclamation
        GoTo CleanExit
    End If
    Set pptPres = pptApp.ActivePresentation

    Set customLayout = Nothing
    Dim d As Object, cl As Object
    For Each d In pptPres.Designs
        For Each cl In d.SlideMaster.CustomLayouts
            If LCase$(cl.Name) = layoutName Then
                Set customLayout = cl
                Exit For
            End If
        Next cl
        If Not customLayout Is Nothing Then Exit For
    Next d
    If customLayout Is Nothing Then
        MsgBox "Custom layout '" & layoutName & "' not found in the active presentation.", vbExclamation
        GoTo CleanExit
    End If
    slideWidth = pptPres.PageSetup.SlideWidth

SelectOutputs:
    If wbOutputMap.Count = 0 Or restartInput Then
        wbOutputMap.RemoveAll
        Dim numOutputs As Long
        Dim outName As String, rngName As String
        Dim defaultsNames As Variant
        Dim defaultsRanges As Variant
        Dim idx As Long

        defaultsNames = Array("Cash Flows", "Valuations & Returns", "Operating Build", "Output 4", "Output 5")
        defaultsRanges = Array("CF", "RET", "OP", "OUT4", "OUT5")

        numOutputs = Application.InputBox("How many outputs to create? (1-5)", "Number of Outputs", 2, , , , , 1)
        If numOutputs < 1 Or numOutputs > 5 Then GoTo CleanExit

        For idx = 1 To numOutputs
            Dim namePrompt As String
            Dim rangePrompt As String
            Dim defaultName As String
            Dim defaultRange As String

            If idx <= UBound(defaultsNames) + 1 Then
                defaultName = defaultsNames(idx - 1)
                defaultRange = defaultsRanges(idx - 1)
            Else
                defaultName = "Output " & idx
                defaultRange = "OUT" & idx
            End If

            namePrompt = "Enter display name for output #" & idx & ":"
            rangePrompt = "Enter named range for '" & defaultName & "':"

            outName = InputBox(namePrompt, "Output Name", defaultName)
            If Trim$(outName) = "" Then GoTo CleanExit
            rngName = InputBox(rangePrompt, "Named Range", defaultRange)
            If Trim$(rngName) = "" Then GoTo CleanExit
            wbOutputMap(outName) = rngName
        Next idx

        restartInput = False
    End If

CasesInput:
    userInput = InputBox("Enter the cases to print, separated by commas:" & vbCrLf & _
                         "Type 'restart' to change output selection.", _
                         "Case Selection", "Base, Upside, Downside")
    If Trim$(userInput) = "" Then GoTo CleanExit
    If LCase$(Trim$(userInput)) = "restart" Then
        restartInput = True
        GoTo SelectOutputs
    End If

    cases = Split(userInput, ",")

    Dim caseName As String
    For i = LBound(cases) To UBound(cases)
        caseName = Trim$(cases(i))
        If caseName <> "" Then
            caseCell.Value = caseName
            Application.CalculateFull
            DoEvents

            For Each key In wbOutputMap.Keys
                Set rng = Nothing
                On Error Resume Next
                Set rng = wb.Names(wbOutputMap(key)).RefersToRange
                On Error GoTo 0

                If rng Is Nothing Then
                    MsgBox "Named range '" & wbOutputMap(key) & "' not found. Skipping output '" & key & "'.", vbExclamation
                Else
                    Set slide = pptPres.Slides.AddSlide(pptPres.Slides.Count + 1, customLayout)
                    rng.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
                    slide.Shapes.Paste
                    Set pic = slide.Shapes(slide.Shapes.Count)
                    With pic
                        .LockAspectRatio = msoTrue
                        If .Width > 0 Then .ScaleWidth (9.5 * 72) / .Width, msoFalse, msoScaleFromTopLeft
                        .Left = (slideWidth - .Width) / 2
                        .Top = 0.74 * 72
                    End With

                    If Not slide.Shapes.Title Is Nothing Then
                        slide.Shapes.Title.TextFrame.TextRange.Text = key & " | " & caseName
                    End If
                End If
            Next key
        End If
    Next i

CleanExit:
    On Error Resume Next
    Application.Calculation = originalCalc
    Application.ScreenUpdating = originalScreen
    Application.EnableEvents = originalEvents
    On Error GoTo 0
    Exit Sub

CleanFail:
    MsgBox "Error running econs export: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
