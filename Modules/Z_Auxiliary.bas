Attribute VB_Name = "Z_Auxiliary"
Option Explicit
Option Private Module
Private currentSeriesIndex As Long
Public gScrollLockMode As Boolean
Private Const DATA_LABEL_STEP As Double = 2#
Private Const CHART_MOVE_STEP As Double = DATA_LABEL_STEP * 10#
' Shared wrapper for temporarily disabling events / screen updates.

Private Function SuppressExcelUi(Optional ByVal hideStatusBar As Boolean = False) As ExcelUiGuard
    Dim guard As ExcelUiGuard
    ThisWorkbook.EnsureAppHook
    Set guard = New ExcelUiGuard
    guard.EnsureWindow
    If hideStatusBar Then guard.DisableStatusBar
    Set SuppressExcelUi = guard
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
    Dim fontColorsArray As Variant
    Static fontCycleIndex As Long
    Static fontCycleLastAddress As String
    Static fontCycleLastStamp As Long
    Dim currentAddress As String
    ' Exact colors
    fontColorsArray = Array(RGB(0, 0, 0), RGB(255, 255, 255), RGB(0, 0, 255), RGB(0, 128, 0), RGB(153, 0, 0))
    ' Exit if selection is not a range
    If TypeName(Selection) <> "Range" Then Exit Sub
    ' Reset cycle if selection changes or cursor moved
    currentAddress = Selection.Address
    If currentAddress <> fontCycleLastAddress _
        Or gSelectionStamp <> fontCycleLastStamp Then
        fontCycleIndex = 0
    End If
    fontCycleLastAddress = currentAddress
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
    If TypeName(Selection) = "Range" Then
        Selection.ClearContents
    ElseIf VarType(Selection) = vbObject Then
        Selection.Delete
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
    ' Border styles to cycle: none, dotted, double
    borderStyles = Array(xlDot, xlDouble, xlNone)
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

    Application.Dialogs(xlDialogInsertHyperlink).Show
    Exit Sub

Catch:
    Call ErrorHandler("InsertHyperlinkDialog")
End Sub

Sub CopyPasteAsPictureToPPT()
    Dim engine As Object
    On Error GoTo CleanFail
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub
    engine.CopySelectionToPowerPoint
    Exit Sub
CleanFail:
    Call ErrorHandler("CopyPasteAsPictureToPPT")
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

Private Function CopySelectionAsPicturePrintSafe(sel As Object) As Boolean
    Dim engine As Object
    On Error GoTo CleanExit
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit
    CopySelectionAsPicturePrintSafe = engine.CopySelectionAsPicturePrintSafe
CleanExit:
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
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    Dim sel As Range
    Dim cell As Range
    Dim formulaText As String
    Dim arrayWarning As Boolean
    Dim innerFormula As String
    ' Ensure a valid range is selected
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.CountLarge = 0 Then Exit Sub
    Set sel = Selection
    ' Loop through all cells in the selection
    For Each cell In sel.Cells
        If cell.HasArray Then
            arrayWarning = True
        ElseIf cell.hasFormula Then
            If Not IsError(cell.value) And IsNumeric(cell.value) And Not IsEmpty(cell.value) Then
                formulaText = cell.formula
                If Len(formulaText) > 3 _
                    And Left$(formulaText, 3) = "=-(" _
                    And Right$(formulaText, 1) = ")" Then
                    innerFormula = Mid$(formulaText, 4, Len(formulaText) - 4)
                    cell.formula = "=" & innerFormula
                ElseIf Len(formulaText) > 1 And Left$(formulaText, 1) = "=" Then
                    cell.formula = "=-(" & Mid$(formulaText, 2) & ")"
                End If
            End If
        ElseIf Not IsError(cell.value) Then
            If IsNumeric(cell.value) And Not IsEmpty(cell.value) Then
                cell.value = -cell.value
            End If
        End If
    Next cell
    If arrayWarning Then
        Call SetStatusBarTemporarily("Skipped array formulas when flipping signs.", 2000)
    End If
End Sub

Sub ReverseSelectionOrder()
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    On Error GoTo Catch

    Dim sel As Range
    Dim cell As Range
    Dim total As Long
    Dim values() As Variant
    Dim formulas() As String
    Dim hasFormula() As Boolean
    Dim i As Long
    Dim srcIndex As Long

    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Areas.Count > 1 Then
        Call SetStatusBarTemporarily("Reverse order requires a contiguous range.", 2000)
        Exit Sub
    End If

    Set sel = Selection
    total = sel.Cells.CountLarge
    If total < 2 Then Exit Sub

    ReDim values(1 To total)
    ReDim formulas(1 To total)
    ReDim hasFormula(1 To total)

    i = 1
    For Each cell In sel.Cells
        If cell.MergeCells Then
            Call SetStatusBarTemporarily("Reverse order skips merged cells.", 2000)
            Exit Sub
        End If
        If cell.HasArray Then
            Call SetStatusBarTemporarily("Reverse order does not support array formulas.", 2000)
            Exit Sub
        End If
        hasFormula(i) = cell.hasFormula
        If hasFormula(i) Then
            formulas(i) = cell.formula
        Else
            values(i) = cell.value
        End If
        i = i + 1
    Next cell

    For i = 1 To total
        srcIndex = total - i + 1
        With sel.Cells(i)
            If hasFormula(srcIndex) Then
                .formula = formulas(srcIndex)
            Else
                .value = values(srcIndex)
            End If
        End With
    Next i

    Call SetStatusBarTemporarily("Selection order reversed.", 2000)
    Exit Sub

Catch:
    Call ErrorHandler("ReverseSelectionOrder")
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
    Dim engine As Object
    If gScrollLockMode Then ScrollLockScroll 0, -steps: gVim.Count1 = 1: Exit Sub
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo FallBackLeft
    If Not engine.MoveSelectedLabels(-DATA_LABEL_STEP, 0#) Then
        If Not engine.MoveSelectedChart(-CHART_MOVE_STEP, 0#) Then GoTo FallBackLeft
    End If
    gVim.Count1 = 1
    Exit Sub
FallBackLeft:
    Call MoveLeft
    gVim.Count1 = 1
End Sub

Sub MoveRightSmart()
    Dim steps As Long: steps = gVim.Count1: If steps < 1 Then steps = 1
    Dim engine As Object
    If gScrollLockMode Then ScrollLockScroll 0, steps: gVim.Count1 = 1: Exit Sub
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo FallBackRight
    If Not engine.MoveSelectedLabels(DATA_LABEL_STEP, 0#) Then
        If Not engine.MoveSelectedChart(CHART_MOVE_STEP, 0#) Then GoTo FallBackRight
    End If
    gVim.Count1 = 1
    Exit Sub
FallBackRight:
    Call MoveRight
    gVim.Count1 = 1
End Sub

Sub MoveUpSmart()
    Dim steps As Long: steps = gVim.Count1: If steps < 1 Then steps = 1
    Dim engine As Object
    If gScrollLockMode Then ScrollLockScroll -steps, 0: gVim.Count1 = 1: Exit Sub
    Set engine = NetAddin()
    If Not engine Is Nothing Then
        If engine.MoveSelectedLabels(0#, -DATA_LABEL_STEP) Then gVim.Count1 = 1: Exit Sub
        If engine.MoveSelectedChart(0#, -CHART_MOVE_STEP) Then gVim.Count1 = 1: Exit Sub
    End If
    Dim i As Long: For i = 1 To steps: KeyStroke Up_: Next i
    gVim.Count1 = 1
End Sub

Sub MoveDownSmart()
    Dim steps As Long: steps = gVim.Count1: If steps < 1 Then steps = 1
    Dim engine As Object
    If gScrollLockMode Then ScrollLockScroll steps, 0: gVim.Count1 = 1: Exit Sub
    Set engine = NetAddin()
    If Not engine Is Nothing Then
        If engine.MoveSelectedLabels(0#, DATA_LABEL_STEP) Then gVim.Count1 = 1: Exit Sub
        If engine.MoveSelectedChart(0#, CHART_MOVE_STEP) Then gVim.Count1 = 1: Exit Sub
    End If
    Dim i As Long: For i = 1 To steps: KeyStroke Down_: Next i
    gVim.Count1 = 1
End Sub



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









Private Function ResolveCopyTarget(ByVal sel As Variant) As Object
    Dim obj As Object
    On Error Resume Next
    If VarType(sel) = vbObject Then
        Set obj = sel
    End If
    On Error GoTo 0

    If obj Is Nothing Then Exit Function

    Dim attempt As Object
    Set attempt = obj
    Do While Not attempt Is Nothing
        Select Case TypeName(attempt)
            Case "Range", "ChartObject", "Chart", "Picture", "Shape"
                If TypeName(attempt) = "Shape" Then
                    If attempt.hasChart Then
                        Set ResolveCopyTarget = attempt.Chart
                        Exit Function
                    End If
                End If
                Set ResolveCopyTarget = attempt
                Exit Function
            Case Else
                On Error Resume Next
                Set attempt = attempt.Parent
                On Error GoTo 0
        End Select
    Loop

    Set ResolveCopyTarget = obj
End Function

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

