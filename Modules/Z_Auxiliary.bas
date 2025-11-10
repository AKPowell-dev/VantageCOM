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

    Select Case g
        Case "0": KeyStroke k0_
        Case "1": KeyStroke k1_
        Case "2": KeyStroke k2_
        Case "3": KeyStroke k3_
        Case "4": KeyStroke k4_
        Case "5": KeyStroke k5_
        Case "6": KeyStroke k6_
        Case "7": KeyStroke k7_
        Case "8": KeyStroke k8_
        Case "9": KeyStroke k9_
        Case Else
            If Len(g) = 1 Then KeyStroke AscW(g)
    End Select

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

Private Function ApplyNumberFormatCycle(ByRef formats As Variant, _
    ByRef lastIndex As Long, _
    ByRef lastAddress As String, _
    ByRef lastActiveCellAddress As String, _
    ByRef lastSelectionStamp As Long) As Boolean

    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)

    If TypeName(Selection) <> "Range" Then
        ApplyNumberFormatCycle = False
        Exit Function
    End If

    If Selection.Address <> lastAddress _
        Or ActiveCell.Address <> lastActiveCellAddress _
        Or gSelectionStamp <> lastSelectionStamp Then
        lastIndex = 0
    End If

    lastAddress = Selection.Address
    lastActiveCellAddress = ActiveCell.Address
    lastSelectionStamp = gSelectionStamp

    Selection.NumberFormat = formats(lastIndex)

    lastIndex = lastIndex + 1
    If lastIndex > UBound(formats) Then lastIndex = 0

    ApplyNumberFormatCycle = False
End Function

Function CycleNumberFormat(Optional ByVal g As String) As Boolean
    Static lastIndex As Long
    Static lastAddress As String
    Static lastActiveCellAddress As String
    Static lastSelectionStamp As Long
    Dim formats As Variant
    formats = Array( _
        "#,##0_);(#,##0);--", _
        "#,##0_);(#,##0);--", _
        "#,##0.0%_);(#,##0.0%);--\%_)", _
        "#,##0.0x_);(#,##0.0x);--x_)", _
        "#,##0""bps""_);(#,##0""bps"");""--bps """, _
        """On"";"""";""Off""", _
        "[>=1]""Yes"";""No"";""No""", _
        "[=1]0"" Year"";0"" Years""", _
        """Year ""0; ""Year ""-0; ""Year 0""; """"" _
    )
    CycleNumberFormat = ApplyNumberFormatCycle(formats, lastIndex, lastAddress, lastActiveCellAddress, lastSelectionStamp)
End Function

Function BinaryCycle(Optional ByVal g As String) As Boolean
    Static lastIndex As Long
    Static lastAddress As String
    Static lastActiveCellAddress As String
    Static lastSelectionStamp As Long
    Dim formats As Variant
    formats = Array( _
        "[>=1]""Yes"";""No"";""No""", _
        """On"";"""";""Off""" _
    )
    BinaryCycle = ApplyNumberFormatCycle(formats, lastIndex, lastAddress, lastActiveCellAddress, lastSelectionStamp)
End Function

Function YearDisplayCycle(Optional ByVal g As String) As Boolean
    Static lastIndex As Long
    Static lastAddress As String
    Static lastActiveCellAddress As String
    Static lastSelectionStamp As Long
    Dim formats As Variant
    formats = Array( _
        "yyyy", _
        "mmm-yyyy" _
    )
    YearDisplayCycle = ApplyNumberFormatCycle(formats, lastIndex, lastAddress, lastActiveCellAddress, lastSelectionStamp)
End Function

Function NumberNarrativeCycle(Optional ByVal g As String) As Boolean
    Static lastIndex As Long
    Static lastAddress As String
    Static lastActiveCellAddress As String
    Static lastSelectionStamp As Long
    Dim formats As Variant
    formats = Array( _
        "#,##0_);(#,##0);--_", _
        "#,##0.0x_);(#,##0.0x);--x_", _
        "[=1]0"" Year"";0"" Years""", _
        """Year ""0; ""Year ""-0; ""Year 0""; """"" _
    )
    NumberNarrativeCycle = ApplyNumberFormatCycle(formats, lastIndex, lastAddress, lastActiveCellAddress, lastSelectionStamp)
End Function

Function PercentCycle(Optional ByVal g As String) As Boolean
    Static lastIndex As Long
    Static lastAddress As String
    Static lastActiveCellAddress As String
    Static lastSelectionStamp As Long
    Dim formats As Variant
    formats = Array( _
        "#,##0.0%_);(#,##0.0%);--\%_)", _
        "#,##0""bps""_);(#,##0""bps"");""--bps """ _
    )
    PercentCycle = ApplyNumberFormatCycle(formats, lastIndex, lastAddress, lastActiveCellAddress, lastSelectionStamp)
End Function

Function CurrencyCycle(Optional ByVal g As String) As Boolean
    Static lastIndex As Long
    Static lastAddress As String
    Static lastActiveCellAddress As String
    Static lastSelectionStamp As Long
    Dim formats As Variant
    Dim poundSymbol As String
    Dim euroSymbol As String

    poundSymbol = ChrW$(&HA3)
    euroSymbol = ChrW$(&H20AC)

    formats = Array( _
        "$#,##0_);($#,##0);$--_)", _
        poundSymbol & "#,##0_);(" & poundSymbol & "#,##0);" & poundSymbol & "--_)", _
        euroSymbol & "#,##0_);(" & euroSymbol & "#,##0);" & euroSymbol & "--_)" _
    )

    CurrencyCycle = ApplyNumberFormatCycle(formats, lastIndex, lastAddress, lastActiveCellAddress, lastSelectionStamp)
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
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    Dim rng As Range
    Set rng = Selection
    If rng Is Nothing Then Exit Sub
    On Error Resume Next
    With rng
        .Borders.lineStyle = xlNone
        .Interior.ColorIndex = xlNone
        .NumberFormat = "#,##0_);(#,##0);--_)"
        ' Remove bold and italic for the whole selection in one pass
        .Font.Bold = False
        .Font.Italic = False
        .Font.Underline = xlUnderlineStyleNone
        .Font.Color = vbBlack
    End With
    On Error GoTo 0
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
    On Error GoTo CleanFail

    Dim pptApp As Object
    Dim pptWindow As Object
    Dim pptSlide As Object
    Dim pptShape As Object
    Dim pastedShp As Object
    Dim sel As Object
    Dim hadTarget As Boolean
    Dim targetLeft As Single, targetTop As Single, targetWidth As Single, targetHeight As Single
    Dim targetRight As Single
    Dim desiredZ As Long, i As Long

    '---- Ensure valid Excel selection ----
    If TypeName(Selection) = "Nothing" Then
        MsgBox "Please select a range, chart, or shape first.", vbExclamation
        Exit Sub
    End If
    Set sel = Selection

    '---- Copy as picture (print appearance) ----
    Dim copied As Boolean
    copied = CopySelectionAsPicturePrintSafe(sel)
    If Not copied Then
        MsgBox "Unable to copy selection as picture (print view). Try selecting a different range or chart.", vbExclamation
        Exit Sub
    End If

    '---- Connect to running PowerPoint ----
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

    '---- Detect selected shape (optional replace target) ----
    On Error Resume Next
    Set pptShape = pptWindow.Selection.ShapeRange
    On Error GoTo 0
    If Not pptShape Is Nothing Then
        hadTarget = True
        targetLeft = pptShape.Left
        targetTop = pptShape.Top
        targetWidth = pptShape.Width
        targetHeight = pptShape.Height
        targetRight = targetLeft + targetWidth
        If TypeName(pptShape) = "ShapeRange" And pptShape.Count > 0 Then
            desiredZ = pptShape(1).ZOrderPosition
        Else
            desiredZ = pptShape.ZOrderPosition
        End If
    End If

    '---- Paste from clipboard ----
    On Error Resume Next
    Set pastedShp = pptSlide.Shapes.PasteSpecial(DataType:=2) ' 2 = ppPasteEnhancedMetafile
    If pastedShp Is Nothing Then Set pastedShp = pptSlide.Shapes.Paste
    On Error GoTo CleanFail

    If TypeName(pastedShp) = "ShapeRange" Then
        Set pastedShp = pastedShp(1)
    End If

    '---- Replace behavior: delete old shape, keep new size, align top-right ----
    If hadTarget Then
        Dim scaledWidth As Double
        Dim scaledHeight As Double
        Dim scaleFactor As Double
        Dim scaleY As Double

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

        With pastedShp
            .Width = scaledWidth
            .Height = scaledHeight

            Dim offsetLeft As Double
            Dim offsetTop As Double

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

            .Left = offsetLeft
            .Top = offsetTop
        End With

        pptShape.Delete

        ' Maintain same z-order level
        If desiredZ > 0 Then
            If desiredZ > pptSlide.Shapes.Count Then desiredZ = pptSlide.Shapes.Count
            pastedShp.ZOrder msoSendToBack
            For i = 1 To desiredZ - 1
                pastedShp.ZOrder msoBringForward
            Next i
        End If
    Else
        '---- Center if no target ----
        With pastedShp
            .Left = (pptSlide.Master.Width - .Width) / 2
            .Top = (pptSlide.Master.Height - .Height) / 2
        End With
    End If

    '---- Select pasted shape ----
    pastedShp.Select
    pptApp.Activate
    Exit Sub

CleanFail:
    MsgBox "Error: " & Err.description, vbCritical, "CopyPasteAsPictureToPPT"
End Sub


'====================================================
' Helper: Copy selection "As shown when printed"
'====================================================
Private Function CopySelectionAsPicturePrintSafe(sel As Object) As Boolean
    Dim t As String
    t = TypeName(sel)

    Dim originalWindow As Window
    Dim originalSheet As Worksheet
    On Error Resume Next
    Set originalWindow = Application.ActiveWindow
    If Not originalWindow Is Nothing Then
        Set originalSheet = originalWindow.ActiveSheet
    End If
    On Error GoTo 0

    On Error Resume Next

    ' Activate the source to ensure CopyPicture works
    Select Case t
        Case "Range"
            sel.Parent.Activate
            sel.Worksheet.Activate
            sel.Select
            Dim r As Range
            Set r = sel
            If r.Areas.Count > 1 Then Set r = r.Areas(1)

            Err.Clear
            ' Primary: print appearance
            r.CopyPicture appearance:=xlPrinter, format:=xlPicture
            If Err.Number <> 0 Then
                Err.Clear: r.CopyPicture appearance:=xlPrinter, format:=xlBitmap
                If Err.Number <> 0 Then
                    ' Fallbacks
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
            sel.Chart.CopyPicture format:=xlPicture, appearance:=xlPrinter
            If Err.Number <> 0 Then
                Err.Clear: sel.Chart.CopyPicture format:=xlBitmap, appearance:=xlPrinter
                If Err.Number <> 0 Then
                    Err.Clear: sel.Chart.Copy
                End If
            End If

        Case "Chart"
            sel.Parent.Activate
            sel.Activate
            Err.Clear
            sel.CopyPicture format:=xlPicture, appearance:=xlPrinter
            If Err.Number <> 0 Then
                Err.Clear: sel.CopyPicture format:=xlBitmap, appearance:=xlPrinter
                If Err.Number <> 0 Then
                    Err.Clear: sel.Copy
                End If
            End If

        Case "Shape"
            Dim hasChart As Boolean
            hasChart = False
            On Error Resume Next
            hasChart = CBool(sel.hasChart)
            On Error GoTo 0
            sel.Parent.Parent.Activate ' worksheet
            Err.Clear
            If hasChart Then
                sel.Chart.CopyPicture format:=xlPicture, appearance:=xlPrinter
                If Err.Number <> 0 Then
                    Err.Clear: sel.Chart.CopyPicture format:=xlBitmap, appearance:=xlPrinter
                    If Err.Number <> 0 Then Err.Clear: sel.Chart.Copy
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
            ' Unknown: try CopyPicture via CallByName, then fallback to Copy
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

    ' Wait for clipboard payload
    Call WaitForClipboardReady(600)
    CopySelectionAsPicturePrintSafe = HasClipboardContent()

    On Error Resume Next
    If Not originalWindow Is Nothing And originalWindow.Visible Then originalWindow.Activate
    If Not originalSheet Is Nothing Then originalSheet.Activate
    Set originalWindow = Nothing
    Set originalSheet = Nothing
    On Error GoTo 0
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
    Dim uiGuard As ExcelUiGuard
    Dim src As Range
    Dim dest As Range
    Dim rowsCount As Long
    Dim colsCount As Long
    Dim wasCopyMode As Boolean

    On Error GoTo CleanFail

    Set uiGuard = SuppressExcelUi(True)

    Call RepeatRegister("PasteExact")

    Set src = ClipboardGetCopyRange()
    If src Is Nothing Then
        Application.CommandBars.ExecuteMso "Paste"
        Exit Function
    End If

    If TypeName(Selection) <> "Range" Then Exit Function
    Set dest = Selection

    rowsCount = src.Rows.Count
    colsCount = src.Columns.Count

    If dest.Cells.Count = 1 Then
        Set dest = dest.Resize(rowsCount, colsCount)
    ElseIf dest.Rows.Count <> rowsCount Or dest.Columns.Count <> colsCount Then
        Set dest = dest.Resize(rowsCount, colsCount)
    End If

    dest.formula = src.formula
    dest.NumberFormat = src.NumberFormat

    src.Copy
    dest.PasteSpecial xlPasteFormats

    ClipboardSetCopyRange src
    wasCopyMode = (Application.CutCopyMode = xlCopy)
    If wasCopyMode Then src.Copy
    ClipboardRefresh

CleanExit:
    Exit Function

CleanFail:
    Call ErrorHandler("PasteExact")
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
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    Dim sel As Range, ws As Worksheet
    Dim r As Long, startRow As Long, startCol As Long, lastCol As Long
    Dim headerRows As Long, c As Long
    Dim sourceCell As Range
    Dim sourceHasFill As Boolean, sourceFillColor As Long
    Dim overallLastCol As Long
    Dim windowTop As Long, windowBottom As Long
    Dim sourceTopStyle As Long, sourceTopWeight As Long, sourceTopColor As Long
    Dim sourceBottomStyle As Long, sourceBottomWeight As Long, sourceBottomColor As Long
    Dim selectionUnion As Range

    headerRows = 20
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection
    Set ws = sel.Worksheet
    startRow = sel.Row
    startCol = sel.Column
    overallLastCol = startCol

    For r = 1 To sel.Rows.Count
        Set sourceCell = sel.Cells(r, 1)

        ' ---- SAFETY: guard against invalid ColorIndex ----
        If Not IsError(sourceCell.Interior.ColorIndex) Then
            sourceHasFill = (sourceCell.Interior.ColorIndex <> xlNone)
            sourceFillColor = sourceCell.Interior.Color
        Else
            sourceHasFill = False
        End If
        ' --------------------------------------------------

        ' Capture source TOP/BOTTOM borders (we will apply these to the pasted range)
        With sourceCell.Borders(xlEdgeTop)
            sourceTopStyle = .lineStyle
            sourceTopWeight = .Weight
            sourceTopColor = .Color
        End With
        With sourceCell.Borders(xlEdgeBottom)
            sourceBottomStyle = .lineStyle
            sourceBottomWeight = .Weight
            sourceBottomColor = .Color
        End With

        ' Determine last column to fill for this row using NEAREST ROW within ÃƒÆ’Ã†â�™ÃƒÂ¢Ã¢â�šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ�šÃ‚Â±50
        lastCol = ComputeLastColFromNearestRow(ws, startRow + r - 1, startCol + 1, 50)

        If lastCol > startCol Then
            Dim rangeFull As Range, targetRange As Range, cell As Range

            ' FULL span for this row (INCLUDES source col)
            Set rangeFull = ws.Range(ws.Cells(startRow + r - 1, startCol), ws.Cells(startRow + r - 1, lastCol))
            ' The "pasted" area to the right of the source
            Set targetRange = ws.Range(ws.Cells(startRow + r - 1, startCol + 1), ws.Cells(startRow + r - 1, lastCol))

            ' ===== SNAPSHOT LEFT/RIGHT BORDERS =====
            Dim n As Long, j As Long
            Dim LStyle() As Long, LWeight() As Long, LColor() As Long
            Dim RStyle() As Long, RWeight() As Long, RColor() As Long

            n = rangeFull.Cells.Count
            ReDim LStyle(1 To n)
            ReDim LWeight(1 To n)
            ReDim LColor(1 To n)
            ReDim RStyle(1 To n)
            ReDim RWeight(1 To n)
            ReDim RColor(1 To n)

            j = 1
            For Each cell In rangeFull.Cells
                On Error Resume Next
                With cell.Borders(xlEdgeLeft)
                    LStyle(j) = .lineStyle
                    LWeight(j) = .Weight
                    LColor(j) = .Color
                End With
                With cell.Borders(xlEdgeRight)
                    RStyle(j) = .lineStyle
                    RWeight(j) = .Weight
                    RColor(j) = .Color
                End With
                On Error GoTo 0
                j = j + 1
            Next cell
            ' ===== END SNAPSHOT =====

            ' FillRight
            rangeFull.FillRight

            ' Apply formatting from the source to the targets (no L/R border impact)
            With targetRange
                .Font.Name = sourceCell.Font.Name
                .Font.Size = sourceCell.Font.Size
                .Font.Bold = sourceCell.Font.Bold
                .Font.Italic = sourceCell.Font.Italic
                .NumberFormat = sourceCell.NumberFormat
                If sourceHasFill Then
                    .Interior.Color = sourceFillColor
                Else
                    .Interior.Pattern = xlNone
                End If
            End With

            ' Apply ONLY TOP/BOTTOM from source
            If Not targetRange Is Nothing Then
                For Each cell In targetRange.Cells
                    On Error Resume Next
                    If sourceTopStyle <> xlNone Then
                        With cell.Borders(xlEdgeTop)
                            .lineStyle = sourceTopStyle
                            .Weight = sourceTopWeight
                            .Color = sourceTopColor
                        End With
                    Else
                        cell.Borders(xlEdgeTop).lineStyle = xlNone
                    End If
                    If sourceBottomStyle <> xlNone Then
                        With cell.Borders(xlEdgeBottom)
                            .lineStyle = sourceBottomStyle
                            .Weight = sourceBottomWeight
                            .Color = sourceBottomColor
                        End With
                    Else
                        cell.Borders(xlEdgeBottom).lineStyle = xlNone
                    End If
                    On Error GoTo 0
                Next cell
            End If

            ' Restore LEFT/RIGHT borders exactly as before
            j = 1
            For Each cell In rangeFull.Cells
                On Error Resume Next
                With cell.Borders(xlEdgeLeft)
                    .lineStyle = LStyle(j)
                    If LStyle(j) <> xlNone Then .Weight = LWeight(j): .Color = LColor(j)
                End With
                With cell.Borders(xlEdgeRight)
                    .lineStyle = RStyle(j)
                    If RStyle(j) <> xlNone Then .Weight = RWeight(j): .Color = RColor(j)
                End With
                On Error GoTo 0
                j = j + 1
            Next cell

            If lastCol > overallLastCol Then overallLastCol = lastCol
        End If
    Next r

    ' --- Release the guard BEFORE setting the final selection ---
    Set uiGuard = Nothing

    ' ===== FINAL SELECTION: rectangular "full operated block" selection =====
    Dim resultRange As Range
    If overallLastCol > startCol Then
        Set resultRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + sel.Rows.Count - 1, overallLastCol))
        SafeSelectRange resultRange
    Else
        SafeSelectRange sel
    End If
End Sub

' ===== Apply finance formatting across contiguous right cells =====
Sub SmartFormatRight()
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    On Error GoTo CleanFail

    Dim sel As Range
    Dim ws As Worksheet
    Dim startRow As Long, startCol As Long
    Dim lastCol As Long
    Dim r As Long
    Dim formattedUnion As Range

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection
    Set ws = sel.Worksheet
    startRow = sel.Row
    startCol = sel.Column

    For r = 1 To sel.Rows.Count
        lastCol = ComputeLastColFromNearestRow(ws, startRow + r - 1, startCol + 1, 50)
        If lastCol >= startCol Then
            Dim rowSpan As Range
            Dim targetRange As Range
            Set rowSpan = ws.Range(ws.Cells(startRow + r - 1, startCol), ws.Cells(startRow + r - 1, lastCol))
            If lastCol > startCol Then
                Set targetRange = ws.Range(ws.Cells(startRow + r - 1, startCol + 1), ws.Cells(startRow + r - 1, lastCol))
            Else
                Set targetRange = Nothing
            End If

            Dim fmt As String
            fmt = "$#,##0_);($#,##0);$--_)"

            On Error Resume Next
            If Not targetRange Is Nothing Then
                targetRange.NumberFormat = fmt
                If Err.Number <> 0 Then
                    Err.Clear
                    targetRange.NumberFormatLocal = fmt
                End If
                On Error GoTo CleanFail

                With targetRange.Font
                    .Name = "Garamond"
                    .Bold = True
                End With
            Else
                Err.Clear
            End If
            On Error GoTo CleanFail

            Dim sourceCell As Range
            Set sourceCell = sel.Cells(r, 1)
            If Not targetRange Is Nothing Then
                If Not IsError(sourceCell.Interior.Pattern) _
                    And sourceCell.Interior.Pattern <> xlNone Then
                    targetRange.Interior.Color = sourceCell.Interior.Color
                Else
                    targetRange.Interior.Pattern = xlNone
                End If
            End If

            With rowSpan.Font
                .Name = "Garamond"
                .Bold = True
            End With
            With rowSpan.Borders(xlEdgeTop)
                .lineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbBlack
            End With

            If formattedUnion Is Nothing Then
                Set formattedUnion = rowSpan
            Else
                Set formattedUnion = Union(formattedUnion, rowSpan)
            End If
        End If
    Next r

    If Not formattedUnion Is Nothing Then
        SafeSelectRange formattedUnion
    End If

CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("SmartFormatRight")
    Resume CleanExit
End Sub

' ===== Outline selection with navy box and corner markers =====
Sub OutlineSelectionHighlight()
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    On Error GoTo CleanFail

    If TypeName(Selection) <> "Range" Then Exit Sub

    Dim sel As Range
    Dim ws As Worksheet
    Set sel = Selection
    Set ws = sel.Worksheet

    Dim firstRow As Long, lastRow As Long
    Dim firstCol As Long, lastCol As Long
    firstRow = sel.Row
    firstCol = sel.Column
    lastRow = firstRow + sel.Rows.Count - 1
    lastCol = firstCol + sel.Columns.Count - 1

    Dim skipTop As Boolean, skipLeft As Boolean
    skipTop = (firstRow = 1)
    skipLeft = (firstCol = 1)

    Dim highlightColor As Long
    highlightColor = RGB(0, 32, 96)

    Dim processed As Object
    Set processed = CreateObject("Scripting.Dictionary")

    Dim r As Long, c As Long

    If Not skipTop Then
        For c = firstCol To lastCol
            HighlightOutlineCell ws, firstRow, c, highlightColor, firstRow, lastRow, firstCol, lastCol, skipTop, skipLeft, processed
        Next c
    End If

    For c = firstCol To lastCol
        HighlightOutlineCell ws, lastRow, c, highlightColor, firstRow, lastRow, firstCol, lastCol, skipTop, skipLeft, processed
    Next c

    If Not skipLeft Then
        For r = firstRow To lastRow
            HighlightOutlineCell ws, r, firstCol, highlightColor, firstRow, lastRow, firstCol, lastCol, skipTop, skipLeft, processed
        Next r
    End If

    For r = firstRow To lastRow
        HighlightOutlineCell ws, r, lastCol, highlightColor, firstRow, lastRow, firstCol, lastCol, skipTop, skipLeft, processed
    Next r

    On Error Resume Next
    ws.Columns(firstCol).ColumnWidth = 2
    ws.Columns(lastCol).ColumnWidth = 2
    On Error GoTo CleanFail

CleanExit:
    Exit Sub
CleanFail:
    Call ErrorHandler("OutlineSelectionHighlight")
    Resume CleanExit
End Sub

Private Sub HighlightOutlineCell(ws As Worksheet, ByVal rowIdx As Long, ByVal colIdx As Long, _
                                 ByVal highlightColor As Long, _
                                 ByVal firstRow As Long, ByVal lastRow As Long, _
                                 ByVal firstCol As Long, ByVal lastCol As Long, _
                                 ByVal skipTop As Boolean, ByVal skipLeft As Boolean, _
                                 ByVal processed As Object)
    If rowIdx < 1 Or colIdx < 1 Then Exit Sub
    If rowIdx > ws.Rows.Count Or colIdx > ws.Columns.Count Then Exit Sub

    Dim key As String
    key = CStr(rowIdx) & "|" & CStr(colIdx)
    If processed.Exists(key) Then Exit Sub
    processed.Add key, True

    Dim cell As Range
    Set cell = ws.Cells(rowIdx, colIdx)

    If Len(cell.Value2 & vbNullString) <> 0 Then Exit Sub

    cell.Interior.Pattern = xlSolid
    cell.Interior.Color = highlightColor

    Dim isCorner As Boolean
    isCorner = False
    If rowIdx = firstRow And colIdx = firstCol Then
        If Not skipTop And Not skipLeft Then isCorner = True
    ElseIf rowIdx = firstRow And colIdx = lastCol Then
        isCorner = True
    ElseIf rowIdx = lastRow And colIdx = firstCol Then
        isCorner = True
    ElseIf rowIdx = lastRow And colIdx = lastCol Then
        isCorner = True
    End If

    If isCorner Then
        cell.value = "x"
        cell.HorizontalAlignment = xlCenter
        cell.VerticalAlignment = xlCenter
        cell.Font.Name = "Garamond"
        cell.Font.Size = 11
        cell.Font.Color = vbWhite
        cell.Font.Bold = True
    End If
End Sub
' ===== Helper: nearest row scan ÃƒÆ’Ã†â�™ÃƒÂ¢Ã¢â�šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ�šÃ‚Â±maxOffset =====
Private Function ComputeLastColFromNearestRow(ws As Worksheet, baseRow As Long, startCol As Long, maxOffset As Long) As Long
    Dim offset As Long, upRow As Long, downRow As Long, lastC As Long
    lastC = ContiguousSpanLastCol(ws, baseRow, startCol)
    If lastC > startCol Then
        ComputeLastColFromNearestRow = lastC
        Exit Function
    End If
    For offset = 1 To maxOffset
        upRow = baseRow - offset
        downRow = baseRow + offset
        If upRow >= 1 Then
            lastC = ContiguousSpanLastCol(ws, upRow, startCol)
            If lastC > startCol Then ComputeLastColFromNearestRow = lastC: Exit Function
        End If
        If downRow <= ws.Rows.Count Then
            lastC = ContiguousSpanLastCol(ws, downRow, startCol)
            If lastC > startCol Then ComputeLastColFromNearestRow = lastC: Exit Function
        End If
    Next offset
    ComputeLastColFromNearestRow = startCol
End Function

' ===== Helper: walk right until a cell has NO value, NO fill, and NO borders =====
Private Function ContiguousSpanLastCol(ws As Worksheet, rowIndex As Long, startCol As Long) As Long
    Dim c As Long, lastC As Long
    lastC = startCol
    For c = startCol + 1 To ws.Columns.Count
        If HasAnyVisual(ws.Cells(rowIndex, c)) Then
            lastC = c
        Else
            Exit For
        End If
    Next c
    ContiguousSpanLastCol = lastC
End Function

' ===== Helper: cell has value/fill/border =====
Private Function HasAnyVisual(cell As Range) As Boolean
    On Error Resume Next
    Dim hasVal As Boolean, hasFill As Boolean, hasBorder As Boolean
    hasVal = (Len(cell.Value2) > 0)
    If Not IsError(cell.Interior.Pattern) Then
        hasFill = (cell.Interior.Pattern <> xlNone)
    Else
        hasFill = False
    End If
    hasBorder = False
    If cell.Borders(xlEdgeLeft).lineStyle <> xlNone Then hasBorder = True
    If cell.Borders(xlEdgeRight).lineStyle <> xlNone Then hasBorder = True
    If cell.Borders(xlEdgeTop).lineStyle <> xlNone Then hasBorder = True
    If cell.Borders(xlEdgeBottom).lineStyle <> xlNone Then hasBorder = True
    On Error GoTo 0
    HasAnyVisual = (hasVal Or hasFill Or hasBorder)
End Function

Function ClearUnnecessaryFormatting(Optional ByVal g As String) As Boolean
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    On Error GoTo CleanExit

    Dim Wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range
    Dim originalCalc As XlCalculation
    Dim originalSheet As Worksheet
    Dim originalAddress As String

    Set Wb = ActiveWorkbook
    If Wb Is Nothing Then Exit Function

    If Not ActiveSheet Is Nothing Then Set originalSheet = ActiveSheet
    If TypeName(Selection) = "Range" Then
        originalAddress = Selection.Cells(1, 1).Address(False, False)
    Else
        originalAddress = ""
    End If

    originalCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.CutCopyMode = False
    Application.StatusBar = "Optimizing workbook..."

    CleanupBrokenNames Wb

    For Each ws In Wb.Worksheets
        If ws.ProtectContents Then GoTo NextSheet

        lastRow = GetLastUsedRow(ws)
        lastCol = GetLastUsedColumn(ws)

        If lastRow = 0 Or lastCol = 0 Then
            ws.Cells.ClearFormats
            ws.Cells.Interior.ColorIndex = xlColorIndexNone
            ws.DisplayPageBreaks = False
            GoTo NextSheet
        End If

        ws.DisplayPageBreaks = False

        If lastRow < ws.Rows.Count Then
            On Error Resume Next
            ws.Rows(CStr(lastRow + 1) & ":" & ws.Rows.Count).ClearFormats
            ws.Rows(CStr(lastRow + 1) & ":" & ws.Rows.Count).Interior.ColorIndex = xlColorIndexNone
            ws.Rows(CStr(lastRow + 1) & ":" & ws.Rows.Count).FormatConditions.Delete
            On Error GoTo CleanExit
        End If

        If lastCol < ws.Columns.Count Then
            On Error Resume Next
            ws.Range(ws.Columns(lastCol + 1), ws.Columns(ws.Columns.Count)).ClearFormats
            ws.Range(ws.Columns(lastCol + 1), ws.Columns(ws.Columns.Count)).Interior.ColorIndex = xlColorIndexNone
            ws.Range(ws.Columns(lastCol + 1), ws.Columns(ws.Columns.Count)).FormatConditions.Delete
            On Error GoTo CleanExit
        End If

        TrimConditionalFormatting ws, lastRow, lastCol

        On Error Resume Next
        Set rng = ws.UsedRange
        Set rng = Nothing
        On Error GoTo CleanExit

NextSheet:
    Next ws

    If Wb.Connections.Count > 0 Then
        On Error Resume Next
        Wb.RefreshAll
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo CleanExit
    End If

    RefreshExcelCaches Wb

    If Not originalSheet Is Nothing Then
        If IsActiveWorkbookSheet(originalSheet) Then
            SafeActivateWorksheet originalSheet
            If Len(originalAddress) > 0 Then
                On Error Resume Next
                SafeSelectRange originalSheet.Range(originalAddress)
                On Error GoTo CleanExit
            End If
        End If
    End If

CleanExit:
    Application.Calculation = originalCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    If ShouldRestartToFlushCaches(Wb) Then
        ClearUnnecessaryFormatting = False
        Exit Function
    End If

    ClearUnnecessaryFormatting = False
    Call SetStatusBarTemporarily("Workbook formatting cache cleared.", 2500)
End Function

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

Private Sub TrimConditionalFormatting(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    On Error Resume Next
    If ws.Cells.FormatConditions.Count > 0 Then
        ws.Cells.FormatConditions.Delete
    End If
    On Error GoTo 0
End Sub

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
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)

    Dim sel As Range, ws As Worksheet
    Dim c As Long, startRow As Long, startCol As Long, lastRow As Long
    Dim sourceCell As Range, originalCell As Range
    Dim sourceHasFill As Boolean, sourceFillColor As Long
    Dim filledRange As Range
    Dim nearestCol As Long, colOffset As Long
    Dim leftDist As Long, rightDist As Long
    Dim leftCol As Long, rightCol As Long
    Dim leftLastRow As Long, rightLastRow As Long
    Dim neighborCol As Long, neighborLastRow As Long
    Dim scanStart As Long

    If TypeName(Selection) <> "Range" Then Exit Sub

    Set sel = Selection
    Set ws = sel.Worksheet
    Set originalCell = ActiveCell

    startRow = sel.Row
    startCol = sel.Column
    scanStart = startRow + 1

    For c = 1 To sel.Columns.Count
        Set sourceCell = sel.Cells(1, c)

        sourceHasFill = (sourceCell.Interior.ColorIndex <> xlNone)
        sourceFillColor = sourceCell.Interior.Color

        ' --- Find the nearest non-empty column within 5 left/right ---
        leftDist = 9999
        rightDist = 9999
        leftCol = 0
        rightCol = 0
        nearestCol = 0

        ' look left up to 5
        For colOffset = 1 To 5
            If (startCol + c - 1 - colOffset) >= 1 Then
                If scanStart <= ws.Rows.Count Then
                    If Application.CountA(ws.Range(ws.Cells(scanStart, startCol + c - 1 - colOffset), ws.Cells(ws.Rows.Count, startCol + c - 1 - colOffset))) > 0 Then
                        leftDist = colOffset
                        leftCol = startCol + c - 1 - colOffset
                        Exit For
                    End If
                End If
            End If
        Next colOffset

        ' look right up to 5
        For colOffset = 1 To 5
            If (startCol + c - 1 + colOffset) <= ws.Columns.Count Then
                If scanStart <= ws.Rows.Count Then
                    If Application.CountA(ws.Range(ws.Cells(scanStart, startCol + c - 1 + colOffset), ws.Cells(ws.Rows.Count, startCol + c - 1 + colOffset))) > 0 Then
                        rightDist = colOffset
                        rightCol = startCol + c - 1 + colOffset
                        Exit For
                    End If
                End If
            End If
        Next colOffset

        ' --- Pick whichever side is closer ---
        If leftDist < rightDist Then
            nearestCol = leftCol
        ElseIf rightDist < leftDist Then
            nearestCol = rightCol
        ElseIf leftDist = rightDist And leftCol > 0 And rightCol > 0 Then
            ' === NEW: both sides equally close ? pick the shorter one ===
            leftLastRow = startRow
            rightLastRow = startRow
            If scanStart <= ws.Rows.Count Then
                If Application.CountA(ws.Range(ws.Cells(scanStart, leftCol), ws.Cells(ws.Rows.Count, leftCol))) > 0 Then
                    leftLastRow = ws.Cells(ws.Rows.Count, leftCol).End(xlUp).Row
                End If
                If Application.CountA(ws.Range(ws.Cells(scanStart, rightCol), ws.Cells(ws.Rows.Count, rightCol))) > 0 Then
                    rightLastRow = ws.Cells(ws.Rows.Count, rightCol).End(xlUp).Row
                End If
            End If
            If leftLastRow < rightLastRow Then
                nearestCol = leftCol
            Else
                nearestCol = rightCol
            End If
        End If
        ' ===============================================================

        ' --- Determine how far down to fill ---
        If nearestCol > 0 Then
            If scanStart <= ws.Rows.Count Then
                If Application.CountA(ws.Range(ws.Cells(scanStart, nearestCol), ws.Cells(ws.Rows.Count, nearestCol))) > 0 Then
                    neighborLastRow = ws.Cells(ws.Rows.Count, nearestCol).End(xlUp).Row
                Else
                    neighborLastRow = startRow
                End If
            Else
                neighborLastRow = startRow
            End If
            lastRow = neighborLastRow
        Else
            lastRow = startRow
        End If

        ' stop if there?s data in the path
        Dim r As Long
        For r = startRow + 1 To lastRow
            If Not IsEmpty(ws.Cells(r, startCol + c - 1).value) Then
                lastRow = r - 1
                Exit For
            End If
        Next r

        ' --- Fill down if range is valid ---
        If lastRow > startRow Then
            ws.Range(ws.Cells(startRow, startCol + c - 1), ws.Cells(lastRow, startCol + c - 1)).FillDown

            ' Copy source formatting
            With ws.Range(ws.Cells(startRow + 1, startCol + c - 1), ws.Cells(lastRow, startCol + c - 1))
                .Font.Name = sourceCell.Font.Name
                .Font.Size = sourceCell.Font.Size
                .Font.Bold = sourceCell.Font.Bold
                .Font.Italic = sourceCell.Font.Italic
                .NumberFormat = sourceCell.NumberFormat
                If sourceHasFill Then
                    .Interior.Color = sourceFillColor
                Else
                    .Interior.Pattern = xlNone
                End If
            End With

            ' track filled range
            If filledRange Is Nothing Then
                Set filledRange = ws.Range(ws.Cells(startRow, startCol + c - 1), ws.Cells(lastRow, startCol + c - 1))
            Else
                Set filledRange = Union(filledRange, ws.Range(ws.Cells(startRow, startCol + c - 1), ws.Cells(lastRow, startCol + c - 1)))
            End If
        End If
    Next c

    If Not filledRange Is Nothing Then
        filledRange.Select
    Else
        originalCell.Select
    End If
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
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    Dim sel As Range
    Dim formulas As Variant
    Dim i As Long, j As Long
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection
    ' Read all formulas into an array
    formulas = sel.formula
    ' Loop through the array and make references absolute
    For i = 1 To UBound(formulas, 1)
        For j = 1 To UBound(formulas, 2)
            If Left(formulas(i, j), 1) = "=" Then
                ' Convert to absolute reference
                formulas(i, j) = Application.ConvertFormula(formulas(i, j), xlA1, xlA1, xlAbsolute)
            End If
        Next j
    Next i
    ' Write back the array all at once
    sel.formula = formulas
End Sub

Public Sub CycleFormatting()
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    Dim sel As Range, firstCell As Range
    Dim nextStyle As Integer
    Dim BLUE_COLOR As Long, RED_COLOR As Long, LIGHTBLUE_COLOR As Long
    Static lastAddress As String
    Static lastSelectionStamp As Long
    ' define colors (RGB can't be used in Const)
    BLUE_COLOR = RGB(0, 32, 96)
    RED_COLOR = RGB(153, 0, 0)
    LIGHTBLUE_COLOR = RGB(226, 234, 250)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection
    Set firstCell = sel.Cells(1, 1)   ' use first cell to determine next style
    ' Determine next style safely
    If firstCell.Font.Color = RED_COLOR Then
        nextStyle = 2   ' move to light-blue fill
    ElseIf firstCell.Interior.Pattern <> xlNone And firstCell.Interior.Color = BLUE_COLOR Then
        nextStyle = 3   ' move to red font
    ElseIf firstCell.Interior.Pattern <> xlNone And firstCell.Interior.Color = LIGHTBLUE_COLOR Then
        nextStyle = 0   ' reset (clear)
    Else
        nextStyle = 1   ' first style (dark blue fill)
    End If
    ' Reset cycle after any cursor movement
    If Selection.Address <> lastAddress Or gSelectionStamp <> lastSelectionStamp Then
        nextStyle = 1
    End If
    lastAddress = Selection.Address
    lastSelectionStamp = gSelectionStamp
    ' Apply chosen style to full selection
    With sel
        .Font.Name = "Garamond"
        Select Case nextStyle
            Case 1  ' Red font, no fill, bold, underlined
                .Interior.Pattern = xlNone
                .Font.Color = RED_COLOR
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleSingle
                .Borders(xlEdgeTop).lineStyle = xlNone
                .Borders(xlEdgeBottom).lineStyle = xlNone
            Case 2  ' Dark blue fill, white font, bold
                .Interior.Pattern = xlSolid
                .Interior.Color = BLUE_COLOR
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleNone
                .Borders(xlEdgeTop).lineStyle = xlNone
                .Borders(xlEdgeBottom).lineStyle = xlNone

            Case 3  ' Light blue fill, bold, top & bottom borders
                .Interior.Pattern = xlSolid
                .Interior.Color = LIGHTBLUE_COLOR
                .Font.Color = RGB(0, 0, 0)
                .Font.Bold = True
                .Font.Underline = xlUnderlineStyleNone
                With .Borders(xlEdgeTop)
                    .lineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With .Borders(xlEdgeBottom)
                    .lineStyle = xlContinuous
                    .Weight = xlThin
                End With
            Case 0  ' Reset / clear
                .Interior.Pattern = xlNone
                .Font.Color = RGB(0, 0, 0)
                .Font.Bold = False
                .Font.Underline = xlUnderlineStyleNone
                .Borders(xlEdgeTop).lineStyle = xlNone
                .Borders(xlEdgeBottom).lineStyle = xlNone
        End Select
    End With
End Sub
' Go to the first cell referenced in the current cell's formula (cross-sheet)

Public Sub GoToPreviousReference()
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.Count = 0 Then Exit Sub

    Static baseKey As String
    Static baseStamp As Long
    Static linkIndex As Long

    Dim baseCell As Range
    Dim currentKey As String
    Dim triedReset As Boolean

    Set baseCell = Selection.Cells(1, 1)
    currentKey = baseCell.Parent.Parent.Name & "|" & baseCell.Parent.Name & "|" & baseCell.Address(False, False)

    If currentKey <> baseKey Or baseStamp <> gSelectionStamp Then
        baseKey = currentKey
        baseStamp = gSelectionStamp
        linkIndex = 1
    End If

    ' Prefer Excel's native trace precedence navigation (handles cross-sheet/workbook)
    On Error Resume Next
    baseCell.ShowPrecedents
    If Err.Number = 0 Then
        Do
            Err.Clear
            baseCell.NavigateArrow True, 1, linkIndex
            If Err.Number = 0 Then
                baseCell.ShowPrecedents Remove:=True
                linkIndex = linkIndex + 1
                On Error GoTo 0
                Exit Sub
            End If
            If Not triedReset Then
                triedReset = True
                linkIndex = 1
                Err.Clear
                baseCell.ShowPrecedents Remove:=True
                baseCell.ShowPrecedents
            Else
                Exit Do
            End If
        Loop
    End If
    On Error GoTo 0

    ' Fallback: parse first reference textually (A1 style within same workbook)
    Dim target As Range
    Dim ws As Worksheet
    Dim win As Window
    Dim visibleRows As Long, visibleCols As Long
    Dim centerRow As Long, centerCol As Long

    Set target = ResolveFirstReference(baseCell)
    If target Is Nothing Then
        Call SetStatusBarTemporarily("No precedents found.", 2000)
        Exit Sub
    End If

    Set ws = target.Worksheet
    If Not IsActiveWorkbookSheet(ws) Then
        Call SetStatusBarTemporarily("Reference located in another workbook; navigation skipped.", 2000)
        Exit Sub
    End If
    SafeSelectRange target.Cells(1, 1)

    Set win = ActiveWindow
    If win Is Nothing Then Exit Sub
    On Error Resume Next
    visibleRows = win.VisibleRange.Rows.Count
    visibleCols = win.VisibleRange.Columns.Count
    centerRow = Application.Max(target.Row - visibleRows \ 2, 1)
    centerCol = Application.Max(target.Column - visibleCols \ 2, 1)
    win.ScrollRow = centerRow
    win.ScrollColumn = centerCol
    On Error GoTo 0
End Sub


Private Function ResolveFirstReference(ByVal sourceCell As Range) As Range
    If sourceCell Is Nothing Then Exit Function

    Dim formulaBody As String
    formulaBody = sourceCell.formula
    If Len(formulaBody) = 0 Then Exit Function
    If Left$(formulaBody, 1) = "=" Then formulaBody = Mid$(formulaBody, 2)

    Dim sanitized As String
    sanitized = RemoveQuotedLiterals(formulaBody)

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = False
    regex.Global = True
    regex.Pattern = "((?:'[^']+'|[A-Za-z0-9_]+)!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?"

    Dim matches As Object
    Set matches = regex.Execute(sanitized)
    If matches Is Nothing Then Exit Function
    If matches.Count = 0 Then Exit Function

    Dim matchText As String
    matchText = matches(0).value

    Dim sheetPart As String
    Dim addrPart As String
    Dim exclPos As Long
    exclPos = InStr(matchText, "!")

    If exclPos > 0 Then
        sheetPart = Left$(matchText, exclPos - 1)
        addrPart = Mid$(matchText, exclPos + 1)
    Else
        sheetPart = ""
        addrPart = matchText
    End If

    If InStr(addrPart, ":") > 0 Then
        addrPart = Split(addrPart, ":")(0)
    End If

    addrPart = Trim$(addrPart)
    If Len(addrPart) = 0 Then Exit Function

    On Error GoTo Fail
    If Len(sheetPart) > 0 Then
        sheetPart = Replace(sheetPart, "'", "")
        Set ResolveFirstReference = sourceCell.Parent.Parent.Worksheets(sheetPart).Range(addrPart)
    Else
        Set ResolveFirstReference = sourceCell.Parent.Range(addrPart)
    End If
    Exit Function

Fail:
    Set ResolveFirstReference = Nothing
End Function

Private Function RemoveQuotedLiterals(ByVal text As String) As String
    Dim inQuote As Boolean
    Dim i As Long
    Dim ch As String
    Dim buffer As String

    buffer = text
    For i = 1 To Len(buffer)
        ch = Mid$(buffer, i, 1)
        If ch = """" Then
            inQuote = Not inQuote
            Mid$(buffer, i, 1) = " "
        ElseIf inQuote Then
            Mid$(buffer, i, 1) = " "
        End If
    Next i

    RemoveQuotedLiterals = buffer
End Function


Private Function CollectDependents(ByVal source As Range) As Collection
    Dim results As New Collection
    Dim depRange As Range
    Dim area As Range
    Dim cell As Range
    Dim seen As Object
    Dim key As String
    Set seen = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set depRange = source.DirectDependents
    If depRange Is Nothing Then
        Err.Clear
        Set depRange = source.Dependents
    End If
    On Error GoTo 0
    If Not depRange Is Nothing Then
        For Each area In depRange.Areas
            For Each cell In area.Cells
                key = cell.Parent.Parent.Name & "|" & cell.Parent.Name & "|" & cell.Address(False, False)
                If Not seen.Exists(key) Then
                    seen.Add key, True
                    results.Add cell
                End If
            Next cell
        Next area
    End If
    Set CollectDependents = results
End Function

Public Sub GoToNextDependent()
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    If TypeName(Selection) <> "Range" Then Exit Sub

    Static baseKey As String
    Static baseStamp As Long
    Static nextLink As Long

    Dim baseCell As Range
    Dim currentKey As String
    Dim triedReset As Boolean
    Dim ok As Boolean

    Set baseCell = Selection.Cells(1, 1)
    currentKey = baseCell.Parent.Parent.Name & "|" & baseCell.Parent.Name & "|" & baseCell.Address(False, False)

    If currentKey <> baseKey Or baseStamp <> gSelectionStamp Then
        baseKey = currentKey
        baseStamp = gSelectionStamp
        nextLink = 1
    End If

    On Error Resume Next
    baseCell.ShowDependents
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ' Fallback to old method if ShowDependents fails
        Dim deps As Collection
        Dim idx As Long
        Set deps = CollectDependents(baseCell)
        If deps Is Nothing Or deps.Count = 0 Then
            Call SetStatusBarTemporarily("No dependents found.", 2000)
            Exit Sub
        End If
        idx = ((nextLink - 1) Mod deps.Count) + 1
        If IsRangeValid(deps(idx)) Then
            SafeSelectRange deps(idx)
            nextLink = idx + 1
            Exit Sub
        Else
            Call SetStatusBarTemporarily("No dependents found.", 2000)
            Exit Sub
        End If
    End If

    ' Try navigating to the next dependent via arrows (cross-sheet capable)
    Do
        Err.Clear
        ok = False
        baseCell.NavigateArrow False, 1, nextLink
        If Err.Number = 0 Then
            ok = True
        End If
        If ok Then
            baseCell.ShowDependents Remove:=True
            nextLink = nextLink + 1
            On Error GoTo 0
            Exit Sub
        End If

        If Not triedReset Then
            triedReset = True
            nextLink = 1
            ' Re-show to ensure arrows exist
            Err.Clear
            baseCell.ShowDependents Remove:=True
            baseCell.ShowDependents
        Else
            Exit Do
        End If
    Loop

    baseCell.ShowDependents Remove:=True
    On Error GoTo 0
    Call SetStatusBarTemporarily("No dependents found.", 2000)
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
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    On Error GoTo CleanExit
    Dim ch As Chart
    Dim srs As Series
    Dim ax As Axis
    Dim shp As shape
    ' --- Get active or selected chart safely ---
    If Not ActiveChart Is Nothing Then
        Set ch = ActiveChart
    ElseIf TypeName(Selection) = "ChartObject" Then
        Set ch = Selection.Chart
    ElseIf TypeName(Selection) = "Shape" Then
        If Selection.hasChart Then Set ch = Selection.Chart
    End If
    If ch Is Nothing Then Exit Sub
    ' =================================================
    ' 1. Global font formatting (safe)
    ' =================================================
    On Error Resume Next
    With ch.ChartArea
        .Font.Name = "Garamond"
        .Font.Size = 11
        .Font.Color = RGB(0, 0, 0)
    End With
    On Error GoTo 0
    ' =================================================
    ' 2. Format X-axis (outline + ticks)
    ' =================================================
    On Error Resume Next
    Set ax = ch.Axes(xlCategory)
    On Error GoTo 0
    If Not ax Is Nothing Then
        With ax
            On Error Resume Next
            .TickLabelFont.Name = "Garamond"
            .TickLabelFont.Size = 11
            .TickLabelFont.Color = RGB(0, 0, 0)
            .MajorTickMark = xlOutside
            .format.Line.Visible = msoTrue
            .format.Line.ForeColor.RGB = RGB(0, 0, 0)
            On Error GoTo 0
        End With
    End If
    ' =================================================
    ' 3. Remove chart title (if present)
    ' =================================================
    On Error Resume Next
    If ch.HasTitle Then ch.HasTitle = False
    On Error GoTo 0
    ' =================================================
    ' 4. Remove outer chart border (ChartArea outline)
    ' =================================================
    On Error Resume Next
    With ch.ChartArea.format.Line
        .Visible = msoFalse
    End With
    On Error GoTo 0
    ' =================================================
    ' 5. Remove gray background gridlines
    ' =================================================
    On Error Resume Next
    ch.Axes(xlCategory).HasMajorGridlines = False
    ch.Axes(xlValue).HasMajorGridlines = False
    On Error GoTo 0
    ' =================================================
    ' 6. Format bar/column outlines and gap width
    ' =================================================
    Dim chartType As XlChartType
    chartType = ch.chartType
    Select Case chartType
        Case xlBarClustered, xlBarStacked, xlBarStacked100, _
             xlColumnClustered, xlColumnStacked, xlColumnStacked100
            For Each srs In ch.SeriesCollection
                On Error Resume Next
                srs.format.Line.Visible = msoTrue
                srs.format.Line.ForeColor.RGB = RGB(0, 0, 0)
                srs.format.Line.Weight = 0.75
                On Error GoTo 0
            Next srs
            ' --- NEW: Adjust gap width for bar/column charts ---
            On Error Resume Next
            ch.ChartGroups(1).GapWidth = 15
            On Error GoTo 0
    End Select
    ' =================================================
    ' 7. Axis titles + legend (safe)
    ' =================================================
    On Error Resume Next
    If ch.Axes(xlCategory).HasTitle Then
        With ch.Axes(xlCategory).AxisTitle
            .Font.Name = "Garamond"
            .Font.Size = 11
            .Font.Color = RGB(0, 0, 0)
        End With
    End If
    If ch.Axes(xlValue).HasTitle Then
        With ch.Axes(xlValue).AxisTitle
            .Font.Name = "Garamond"
            .Font.Size = 11
            .Font.Color = RGB(0, 0, 0)
        End With
    End If
    If ch.HasLegend Then
        With ch.Legend
            .Font.Name = "Garamond"
            .Font.Size = 11
            .Font.Color = RGB(0, 0, 0)
        End With
    End If
    On Error GoTo 0
CleanExit:
    Set ch = Nothing
    Set ax = Nothing
    Set srs = Nothing
End Sub

' ==================================================
'  LIGHTNING-FAST DATA LABEL MOVE HANDLER
' ==================================================

Private Function TryMoveSelectedLabel(ByVal dx As Double, ByVal dy As Double) As Boolean
    ' Extremely fast ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â�šÂ¬Ã‚Â touches only the current label(s)
    Dim t As String
    Dim pt As Excel.point
    Dim lbl As Excel.DataLabel
    Dim s As Excel.Series

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

        Case "DataPoint"
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

        Case "DataLabels"
            ' Whole label collection selected ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â�šÂ¬Ã‚Â move each very quickly
            On Error Resume Next
            For Each lbl In Selection
                lbl.Position = xlLabelPositionCustom
                lbl.Left = lbl.Left + dx
                lbl.Top = lbl.Top + dy
            Next lbl
            On Error GoTo 0
            TryMoveSelectedLabel = True

        Case "Series"
            ' Move labels for all points in series (rare case)
            On Error Resume Next
            Set s = Selection
            If Not s Is Nothing Then
                For Each pt In s.Points
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
    Dim chartContainer As Object
    Dim sel As Object
    If gScrollLockMode Then Exit Function

    On Error Resume Next
    Set sel = Selection
    On Error GoTo 0
    If sel Is Nothing Then Exit Function
    If Not IsChartMoveSelection(sel) Then Exit Function

    Set chartContainer = ResolveSelectedChartContainer(sel)
    If chartContainer Is Nothing Then Exit Function

    On Error Resume Next
    chartContainer.Left = chartContainer.Left + dx
    chartContainer.Top = chartContainer.Top + dy
    If Err.Number = 0 Then
        TryMoveSelectedChart = True
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Function

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
    If gScrollLockMode Then ScrollLockScroll 0, -steps: gVim.Count1 = 1: Exit Sub
    If Not TryMoveSelectedLabel(-DATA_LABEL_STEP, 0) Then
        If Not TryMoveSelectedChart(-CHART_MOVE_STEP, 0) Then Call MoveLeft
    End If
    gVim.Count1 = 1
End Sub

Sub MoveRightSmart()
    Dim steps As Long: steps = gVim.Count1: If steps < 1 Then steps = 1
    If gScrollLockMode Then ScrollLockScroll 0, steps: gVim.Count1 = 1: Exit Sub
    If Not TryMoveSelectedLabel(DATA_LABEL_STEP, 0) Then
        If Not TryMoveSelectedChart(CHART_MOVE_STEP, 0) Then Call MoveRight
    End If
    gVim.Count1 = 1
End Sub

Sub MoveUpSmart()
    Dim steps As Long: steps = gVim.Count1: If steps < 1 Then steps = 1
    If gScrollLockMode Then ScrollLockScroll -steps, 0: gVim.Count1 = 1: Exit Sub
    If TryMoveSelectedLabel(0, -DATA_LABEL_STEP) Then gVim.Count1 = 1: Exit Sub
    If TryMoveSelectedChart(0, -CHART_MOVE_STEP) Then gVim.Count1 = 1: Exit Sub
    Dim i As Long: For i = 1 To steps: KeyStroke Up_: Next i
    gVim.Count1 = 1
End Sub

Sub MoveDownSmart()
    Dim steps As Long: steps = gVim.Count1: If steps < 1 Then steps = 1
    If gScrollLockMode Then ScrollLockScroll steps, 0: gVim.Count1 = 1: Exit Sub
    If TryMoveSelectedLabel(0, DATA_LABEL_STEP) Then gVim.Count1 = 1: Exit Sub
    If TryMoveSelectedChart(0, CHART_MOVE_STEP) Then gVim.Count1 = 1: Exit Sub
    Dim i As Long: For i = 1 To steps: KeyStroke Down_: Next i
    gVim.Count1 = 1
End Sub



Function SelectNearestChart(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Function

    Dim refX As Double, refY As Double
    If Not TryGetSelectionCenter(refX, refY) Then
        Call SetStatusBarTemporarily("Unable to determine reference position.", 2000)
        Exit Function
    End If

    Dim bestTarget As Object
    Dim bestDist As Double
    bestDist = -1#

    Dim cbo As ChartObject
    For Each cbo In ws.ChartObjects
        If ChartObjectIsVisible(cbo) Then
            Dim dist As Double
            dist = DistanceSquared(refX, refY, cbo.Left + cbo.Width / 2, cbo.Top + cbo.Height / 2)
            If bestDist < 0 Or dist < bestDist Then
                Set bestTarget = cbo
                bestDist = dist
            End If
        End If
    Next cbo

    Dim shp As shape
    For Each shp In ws.Shapes
        If ShapeHasVisibleChart(shp) Then
            dist = DistanceSquared(refX, refY, shp.Left + shp.Width / 2, shp.Top + shp.Height / 2)
            If bestDist < 0 Or dist < bestDist Then
                Set bestTarget = shp
                bestDist = dist
            End If
        End If
    Next shp

    If bestTarget Is Nothing Then
        Call SetStatusBarTemporarily("No charts found on this sheet.", 2000)
    Else
        ActivateChartContainer bestTarget
        EnsureChartElementSelection bestTarget
    End If

    SelectNearestChart = False
    Exit Function

CleanFail:
    Call ErrorHandler("SelectNearestChart")
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
    Dim uiGuard As ExcelUiGuard
    Set uiGuard = SuppressExcelUi(True)
    On Error Resume Next
    Dim sel As Range, cell As Range
    Dim origFormula As String
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection
    For Each cell In sel.Cells
        If cell.hasFormula Then
            origFormula = cell.formula
            ' Safely handle missing or malformed formulas
            If Left$(origFormula, 1) = "=" Then
                origFormula = Mid$(origFormula, 2) ' remove "="
                cell.formula = "=IF(circ=1,0," & origFormula & ")"
            End If
        End If
    Next cell
    On Error GoTo 0
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
