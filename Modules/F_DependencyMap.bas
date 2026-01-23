Attribute VB_Name = "F_DependencyMap"
Option Explicit

Private Type FormulaNavRef
    Token As String
    StartIndex As Long
    Length As Long
End Type

Private mNavActive As Boolean
Private mNavFormula As String
Private mNavRefs() As FormulaNavRef
Private mNavRefCount As Long
Private mNavIndex As Long
Private mNavStartAddress As String
Private mNavStartSheetName As String
Private mNavStartWorkbookName As String
Private mNavLastSelectionStamp As Long
Private mFormulaNavigatorForm As UF_FormulaNavigator
Private mCtrlBracketPassthrough As Boolean
Private Const NAV_STATUS_MAX As Long = 220
Private Const NAV_DEBUG As Boolean = True

Sub DrawDependencyMap()
    Dim engine As Object
    On Error GoTo CleanFail

    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub
    engine.DrawDependencyMap
    Exit Sub

CleanFail:
    Call ErrorHandler("DrawDependencyMap")
End Sub

Public Function TraceIn(Optional ByVal g As String) As Boolean
    Dim engine As Object
    On Error GoTo CleanFail

    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    engine.TracePrecedentsDialog
    Exit Function

CleanFail:
    Call ErrorHandler("TraceIn")
End Function

Public Function TraceOut(Optional ByVal g As String) As Boolean
    Dim engine As Object
    On Error GoTo CleanFail

    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    engine.TraceDependentsDialog
    Exit Function

CleanFail:
    Call ErrorHandler("TraceOut")
End Function

Public Function FormulaNavigatorNext(Optional ByVal g As String) As Boolean
    Dim activeCell As Range
    Dim formulaText As String
    Dim nextIndex As Long
    Dim selectionType As String
    Dim continueNav As Boolean

    On Error GoTo CleanFail

    Call NavTrace("Enter FormulaNavigatorNext")
    On Error Resume Next
    selectionType = TypeName(Selection)
    On Error GoTo CleanFail
    If selectionType <> "" Then
        Call NavTrace("Selection type: " & selectionType)
    End If

    continueNav = mNavActive And (mNavRefCount > 0) And (gSelectionStamp = mNavLastSelectionStamp)
    If continueNav Then
        Call NavTrace("Continuing existing cycle")
    Else
        If Not TryGetActiveCell(activeCell) Then
            Call NavTrace("Active cell unresolved")
            Call SetStatusBarTemporarily("No active cell (selection: " & selectionType & ")", 2000)
            Exit Function
        End If
        Call NavTrace("Active cell: " & activeCell.Address(External:=True))

        If Not activeCell.HasFormula Then
            Call FormulaNavigatorReset
            If IsEmpty(activeCell.Value2) Then
                Call NavTrace("Active cell empty (no formula)")
                Call SetStatusBarTemporarily("Active cell is empty.", 2000)
            Else
                Call NavTrace("Active cell hard-coded value")
                Call SetStatusBarTemporarily("Active cell has a hard-coded value.", 2000)
            End If
            Exit Function
        End If

        formulaText = CStr(activeCell.Formula2)
        If Len(formulaText) = 0 Then
            Call FormulaNavigatorReset
            Call NavTrace("Active cell formula is empty string")
            Call SetStatusBarTemporarily("Active cell has no formula.", 2000)
            Exit Function
        End If

        If (Not mNavActive) _
            Or (mNavStartAddress <> activeCell.Address(External:=True)) _
            Or (mNavFormula <> formulaText) Then
            Call NavTrace("Init required (new cell or formula)")
            If Not InitFormulaNavigator(activeCell, formulaText) Then Exit Function
        End If
    End If

    nextIndex = mNavIndex + 1
    If nextIndex > mNavRefCount Then
        Call NavTrace("Return to start")
        Call GoToFormulaNavigatorStart
        mNavLastSelectionStamp = gSelectionStamp
        mNavIndex = 0
        Call UpdateFormulaNavigatorUI(0)
        Exit Function
    End If

    mNavIndex = nextIndex
    Call NavTrace("Selecting ref " & CStr(mNavIndex) & "/" & CStr(mNavRefCount))
    If Not SelectFormulaReference(mNavRefs(mNavIndex).Token) Then
        Call NavTrace("SelectFormulaReference failed")
        Call SetStatusBarTemporarily("Reference not found: " & mNavRefs(mNavIndex).Token, 2500)
        Call UpdateFormulaNavigatorUI(mNavIndex)
        Exit Function
    End If

    Call CenterActiveCellInWindow
    mNavLastSelectionStamp = gSelectionStamp
    Call UpdateFormulaNavigatorUI(mNavIndex)
    Exit Function

CleanFail:
    Call ErrorHandler("FormulaNavigatorNext")
End Function

Public Function FormulaNavigatorCancel(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    If Not mNavActive Then Exit Function

    Call FormulaNavigatorReset
    Exit Function

CleanFail:
    Call ErrorHandler("FormulaNavigatorCancel")
End Function

Public Function FormulaNavigatorHotkeyFallback(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    If gVim Is Nothing Or Not gVim.Enabled Then
        Call PassThroughCtrlBracketKey
        Exit Function
    End If

    FormulaNavigatorHotkeyFallback = FormulaNavigatorNext
    Exit Function

CleanFail:
    Call ErrorHandler("FormulaNavigatorHotkeyFallback")
End Function

Private Function InitFormulaNavigator(ByVal cell As Range, ByVal formulaText As String) As Boolean
    mNavActive = False
    mNavFormula = formulaText
    mNavStartAddress = cell.Address(External:=True)
    mNavStartSheetName = cell.Parent.Name
    mNavStartWorkbookName = cell.Parent.Parent.Name
    mNavIndex = 0
    mNavRefCount = 0
    mNavLastSelectionStamp = gSelectionStamp
    Erase mNavRefs

    mNavRefCount = CollectFormulaReferences(formulaText, mNavRefs)
    If mNavRefCount <= 0 Then
        Call NavTrace("No references parsed")
        Call SetStatusBarTemporarily("No references found in formula.", 2000)
        Exit Function
    End If

    Call NavTrace("Parsed refs: " & CStr(mNavRefCount))
    mNavActive = True
    InitFormulaNavigator = True
End Function

Private Sub FormulaNavigatorReset()
    mNavActive = False
    mNavFormula = ""
    mNavIndex = 0
    mNavRefCount = 0
    mNavStartAddress = ""
    mNavStartSheetName = ""
    mNavStartWorkbookName = ""
    mNavLastSelectionStamp = 0
    Erase mNavRefs
    Call HideFormulaNavigatorUI
End Sub

Private Sub UpdateFormulaNavigatorUI(ByVal refIndex As Long)
    Dim highlightStart As Long
    Dim highlightLen As Long
    Dim token As String
    Dim navForm As UF_FormulaNavigator

    If refIndex >= 1 And refIndex <= mNavRefCount Then
        highlightStart = mNavRefs(refIndex).StartIndex
        highlightLen = mNavRefs(refIndex).Length
        token = mNavRefs(refIndex).Token
    Else
        highlightStart = 0
        highlightLen = 0
        token = ""
    End If

    If mNavRefCount > 1 Then
        Set navForm = GetFormulaNavigatorForm()
        If Not navForm Is Nothing Then
            On Error GoTo RetryForm
            navForm.Launch mNavFormula, highlightStart, highlightLen, refIndex, mNavRefCount, token, mNavStartAddress
            GoTo ContinueStatus
        End If
    Else
        Call HideFormulaNavigatorUI
    End If
    GoTo ContinueStatus

RetryForm:
    Err.Clear
    Set mFormulaNavigatorForm = Nothing
    Set navForm = GetFormulaNavigatorForm()
    If Not navForm Is Nothing Then
        On Error Resume Next
        navForm.Launch mNavFormula, highlightStart, highlightLen, refIndex, mNavRefCount, token, mNavStartAddress
        On Error GoTo 0
    End If

ContinueStatus:
    Dim statusText As String
    statusText = BuildFormulaStatusText(refIndex, highlightStart, highlightLen)
    If Len(statusText) > 0 Then
        Call SetStatusBarTemporarily(statusText, 3000, True)
    End If
End Sub

Private Sub HideFormulaNavigatorUI()
    On Error Resume Next
    If Not mFormulaNavigatorForm Is Nothing Then
        mFormulaNavigatorForm.HideNavigator
    End If
    On Error GoTo 0
End Sub

Private Function GetFormulaNavigatorForm() As UF_FormulaNavigator
    Dim needsNew As Boolean
    On Error Resume Next
    If mFormulaNavigatorForm Is Nothing Then
        needsNew = True
    Else
        Dim tmpVisible As Boolean
        tmpVisible = mFormulaNavigatorForm.Visible
        If Err.Number <> 0 Then
            Err.Clear
            needsNew = True
        End If
    End If
    On Error GoTo CleanFail

    If needsNew Then
        Set mFormulaNavigatorForm = New UF_FormulaNavigator
    End If

    Set GetFormulaNavigatorForm = mFormulaNavigatorForm
    Exit Function

CleanFail:
    Set mFormulaNavigatorForm = Nothing
End Function

Private Sub GoToFormulaNavigatorStart()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim target As Range
    Dim wbName As String
    Dim wsName As String
    Dim addr As String

    On Error Resume Next
    If Len(mNavStartAddress) > 0 Then
        Application.Goto Reference:=mNavStartAddress, Scroll:=False
    End If
    If Err.Number <> 0 Then
        Err.Clear
        If TryParseQualifiedReference(mNavStartAddress, wbName, wsName, addr) Then
            If Len(wbName) > 0 Then
                On Error Resume Next
                Set wb = Workbooks(wbName)
                On Error GoTo 0
            End If
            If wb Is Nothing Then
                If Len(mNavStartWorkbookName) > 0 Then
                    On Error Resume Next
                    Set wb = Workbooks(mNavStartWorkbookName)
                    On Error GoTo 0
                End If
            End If
            If wb Is Nothing Then Set wb = ActiveWorkbook

            If Not wb Is Nothing Then
                On Error Resume Next
                Set ws = wb.Worksheets(wsName)
                On Error GoTo 0
            End If

            If Not ws Is Nothing Then
                Set target = ws.Range(addr)
                If Not target Is Nothing Then
                    Call SafeActivateWorkbook(wb)
                    Call SafeActivateWorksheet(ws)
                    Call SafeSelectRange(target)
                End If
            End If
        End If
    End If
    On Error GoTo 0
    Call CenterActiveCellInWindow
End Sub

Private Sub CenterActiveCellInWindow()
    On Error Resume Next
    Dim win As Window
    Dim visibleRows As Long
    Dim visibleCols As Long
    Dim targetRow As Long
    Dim targetCol As Long

    Set win = Application.ActiveWindow
    If win Is Nothing Then Exit Sub

    visibleRows = win.VisibleRange.Rows.Count
    visibleCols = win.VisibleRange.Columns.Count

    targetRow = ActiveCell.Row - visibleRows \ 2
    targetCol = ActiveCell.Column - visibleCols \ 2

    If targetRow < 1 Then targetRow = 1
    If targetCol < 1 Then targetCol = 1

    win.ScrollRow = targetRow
    win.ScrollColumn = targetCol
End Sub

Private Function TryGetActiveCell(ByRef cell As Range) As Boolean
    On Error Resume Next
    Set cell = ActiveCell
    On Error GoTo 0
    If Not cell Is Nothing Then
        Call NavTrace("ActiveCell resolved via ActiveCell")
        TryGetActiveCell = True
        Exit Function
    End If

    Dim sel As Object
    On Error Resume Next
    Set sel = Selection
    On Error GoTo 0
    If TypeName(sel) = "Range" Then
        On Error Resume Next
        Set cell = sel.Cells(1, 1)
        On Error GoTo 0
        If Not cell Is Nothing Then
            Call NavTrace("ActiveCell resolved via Selection")
            TryGetActiveCell = True
        End If
        Exit Function
    End If

    Dim win As Window
    On Error Resume Next
    Set win = Application.ActiveWindow
    If Not win Is Nothing Then
        Set cell = win.RangeSelection.Cells(1, 1)
    End If
    On Error GoTo 0
    If Not cell Is Nothing Then
        Call NavTrace("ActiveCell resolved via RangeSelection")
        TryGetActiveCell = True
    End If
End Function

Private Function IsCtrlKeyDown() As Boolean
    On Error Resume Next
    IsCtrlKeyDown = ((GetKeyState(CtrlLeft_) And &H8000) <> 0) _
        Or ((GetKeyState(CtrlRight_) And &H8000) <> 0)
End Function

Private Function IsCtrlBracketKeyPressed() As Boolean
    On Error Resume Next
    If (GetKeyState(OpeningSquareBracket_) And &H8000) <> 0 Then
        IsCtrlBracketKeyPressed = True
        Exit Function
    End If

    ' Ctrl+[ can arrive as Esc; fall back to recent bracket press while Ctrl is down.
    If IsCtrlKeyDown() Then
        IsCtrlBracketKeyPressed = ((GetAsyncKeyState(OpeningSquareBracket_) And &H1) <> 0)
    End If
End Function

Private Function BuildFormulaStatusText(ByVal refIndex As Long, ByVal highlightStart As Long, ByVal highlightLen As Long) As String
    If Len(mNavFormula) = 0 Then Exit Function

    Dim displayText As String
    displayText = mNavFormula

    Dim startPos As Long
    startPos = highlightStart + 1
    If highlightLen > 0 And startPos > 0 And startPos <= Len(displayText) Then
        Dim beforeText As String
        Dim highlightText As String
        Dim afterText As String
        beforeText = Left$(displayText, startPos - 1)
        highlightText = Mid$(displayText, startPos, highlightLen)
        afterText = Mid$(displayText, startPos + highlightLen)
        displayText = beforeText & "<<" & highlightText & ">>" & afterText
    End If

    Dim prefix As String
    If refIndex > 0 Then
        prefix = "Ref " & CStr(refIndex) & "/" & CStr(mNavRefCount) & ": "
    ElseIf mNavRefCount > 0 Then
        prefix = "Base: "
    End If

    BuildFormulaStatusText = prefix & displayText
    If Len(BuildFormulaStatusText) > NAV_STATUS_MAX Then
        BuildFormulaStatusText = Left$(BuildFormulaStatusText, NAV_STATUS_MAX - 3) & "..."
    End If
End Function

Private Sub NavTrace(ByVal msg As String)
    If Not NAV_DEBUG Then Exit Sub
    Debug.Print "[" & Now & "] [NAV] " & msg
    Call SetStatusBarTemporarily("[NAV] " & msg, 1200, True)
End Sub

Private Sub PassThroughCtrlBracketKey()
    If mCtrlBracketPassthrough Then Exit Sub
    mCtrlBracketPassthrough = True
    On Error Resume Next
    Application.OnKey "^{[}"
    KeyStroke Ctrl_ + OpeningSquareBracket_
    Application.OnKey "^{[}", "FormulaNavigatorHotkeyFallback"
    mCtrlBracketPassthrough = False
End Sub

Private Function CollectFormulaReferences(ByVal formulaText As String, ByRef refs() As FormulaNavRef) As Long
    Dim reg As Object
    Dim matches As Object
    Dim matchItem As Object
    Dim count As Long

    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = BuildFormulaReferencePattern()
    reg.Global = True
    reg.IgnoreCase = True

    Set matches = reg.Execute(formulaText)
    If matches.Count = 0 Then Exit Function

    ReDim refs(1 To matches.Count)

    For Each matchItem In matches
        If IsIndexInStringLiteral(formulaText, matchItem.FirstIndex) Then GoTo ContinueLoop
        If IsFunctionLikeToken(formulaText, matchItem) Then GoTo ContinueLoop
        If IsBareNameToken(matchItem.value) Then
            If IsReservedNameToken(matchItem.value) Then GoTo ContinueLoop
            If Not IsDefinedNameToken(matchItem.value) Then GoTo ContinueLoop
        End If

        count = count + 1
        refs(count).Token = matchItem.value
        refs(count).StartIndex = matchItem.FirstIndex
        refs(count).Length = matchItem.Length

ContinueLoop:
    Next matchItem

    If count = 0 Then Exit Function
    If count < UBound(refs) Then ReDim Preserve refs(1 To count)
    CollectFormulaReferences = count
End Function

Private Function BuildFormulaReferencePattern() As String
    Dim prefix As String
    Dim cell As String
    Dim cellRange As String
    Dim colRange As String
    Dim rowRange As String
    Dim structRef As String
    Dim nameToken As String

    prefix = "(?:'[^']+'|\[[^\]]+\][^!]+|[A-Za-z0-9_.]+)!"
    cell = "\$?[A-Za-z]{1,3}\$?\d{1,7}"
    cellRange = cell & "(?::" & cell & ")?"
    colRange = "\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"
    rowRange = "\$?\d{1,7}:\$?\d{1,7}"
    structRef = "[A-Za-z_][A-Za-z0-9_]*\[[^\]]+\]"
    nameToken = "[A-Za-z_][A-Za-z0-9_.]*"

    BuildFormulaReferencePattern = "(?:" & prefix & ")?(?:" & cellRange & "|" & colRange & "|" & rowRange & "|" & nameToken & "|" & structRef & ")"
End Function

Private Function IsFunctionLikeToken(ByVal formulaText As String, ByVal matchItem As Object) As Boolean
    Dim token As String
    Dim idx As Long
    Dim ch As String

    token = matchItem.value
    If InStr(token, "!") > 0 Then Exit Function
    If InStr(token, ":") > 0 Then Exit Function
    If InStr(token, "[") > 0 Then Exit Function

    idx = matchItem.FirstIndex + matchItem.Length + 1
    Do While idx <= Len(formulaText)
        ch = Mid$(formulaText, idx, 1)
        If ch = " " Or ch = vbTab Then
            idx = idx + 1
        Else
            If ch = "(" Then IsFunctionLikeToken = True
            Exit Function
        End If
    Loop
End Function

Private Function IsBareNameToken(ByVal token As String) As Boolean
    If InStr(token, "!") > 0 Then Exit Function
    If InStr(token, ":") > 0 Then Exit Function
    If InStr(token, "[") > 0 Then Exit Function
    If IsPlainAddressToken(token) Then Exit Function
    IsBareNameToken = True
End Function

Private Function IsReservedNameToken(ByVal token As String) As Boolean
    Dim upperToken As String
    upperToken = UCase$(token)
    If upperToken = "TRUE" Or upperToken = "FALSE" Then
        IsReservedNameToken = True
    End If
End Function

Private Function IsDefinedNameToken(ByVal token As String) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nm As Name

    If Len(mNavStartWorkbookName) > 0 Then
        On Error Resume Next
        Set wb = Workbooks(mNavStartWorkbookName)
        On Error GoTo 0
    End If

    If wb Is Nothing Then
        Set wb = ActiveWorkbook
    End If

    On Error Resume Next
    If Not wb Is Nothing Then
        Set nm = wb.Names.Item(token)
        If Err.Number = 0 And Not nm Is Nothing Then
            IsDefinedNameToken = True
            Exit Function
        End If
        Err.Clear
    End If
    On Error GoTo 0

    On Error Resume Next
    If Not wb Is Nothing Then
        If Len(mNavStartSheetName) > 0 Then
            Set ws = wb.Worksheets(mNavStartSheetName)
        Else
            Set ws = ActiveSheet
        End If
        If Not ws Is Nothing Then
            Set nm = ws.Names.Item(token)
            If Err.Number = 0 And Not nm Is Nothing Then
                IsDefinedNameToken = True
                Exit Function
            End If
        End If
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function IsIndexInStringLiteral(ByVal text As String, ByVal zeroBasedIndex As Long) As Boolean
    Dim i As Long
    Dim inString As Boolean
    Dim target As Long

    If zeroBasedIndex < 0 Then Exit Function

    target = zeroBasedIndex + 1
    i = 1
    Do While i <= Len(text) And i <= target
        If Mid$(text, i, 1) = """" Then
            If inString Then
                If i < Len(text) And Mid$(text, i + 1, 1) = """" Then
                    i = i + 1
                Else
                    inString = False
                End If
            Else
                inString = True
            End If
        End If
        If i >= target Then Exit Do
        i = i + 1
    Loop

    IsIndexInStringLiteral = inString
End Function

Private Function SelectFormulaReference(ByVal token As String) As Boolean
    Dim qualified As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim target As Range
    Dim wbName As String
    Dim wsName As String
    Dim addr As String
    On Error GoTo CleanFail

    If TryParseQualifiedReference(token, wbName, wsName, addr) Then
        Call NavTrace("Parsed token: [" & wbName & "]" & wsName & "!" & addr)
        If Len(wbName) > 0 Then
            On Error Resume Next
            Set wb = Workbooks(wbName)
            On Error GoTo CleanFail
        End If
        If wb Is Nothing Then
            If Len(mNavStartWorkbookName) > 0 Then
                On Error Resume Next
                Set wb = Workbooks(mNavStartWorkbookName)
                On Error GoTo CleanFail
            End If
        End If
        If wb Is Nothing Then Set wb = ActiveWorkbook

        If Not wb Is Nothing Then
            On Error Resume Next
            Set ws = wb.Worksheets(wsName)
            On Error GoTo CleanFail
        End If

        If Not ws Is Nothing Then
            Set target = ws.Range(addr)
            If Not target Is Nothing Then
                Call NavTrace("Select range " & ws.Name & "!" & addr)
                Call SafeActivateWorkbook(wb)
                Call SafeActivateWorksheet(ws)
                Call SafeSelectRange(target)
                SelectFormulaReference = True
                Exit Function
            End If
        End If
    End If

    If IsBareNameToken(token) Then
        If TryResolveNameRange(token, target) Then
            Call NavTrace("Select named range " & token)
            Call SafeActivateWorkbook(target.Parent.Parent)
            Call SafeActivateWorksheet(target.Parent)
            Call SafeSelectRange(target)
            SelectFormulaReference = True
            Exit Function
        End If
    End If

    If IsPlainAddressToken(token) Then
        On Error Resume Next
        Set wb = Workbooks(mNavStartWorkbookName)
        On Error GoTo CleanFail
        If wb Is Nothing Then
            Set wb = ActiveWorkbook
        End If

        If Not wb Is Nothing Then
            On Error Resume Next
            Set ws = wb.Worksheets(mNavStartSheetName)
            On Error GoTo CleanFail
            If Not ws Is Nothing Then
                Set target = ws.Range(token)
                If Not target Is Nothing Then
                    Call NavTrace("Select range " & ws.Name & "!" & token)
                    Call SafeActivateWorkbook(wb)
                    Call SafeActivateWorksheet(ws)
                    Call SafeSelectRange(target)
                    SelectFormulaReference = True
                    Exit Function
                End If
            End If
        End If
    End If

    qualified = BuildQualifiedReference(token)
    Call NavTrace("Goto reference: " & qualified)
    Application.Goto Reference:=qualified, Scroll:=False
    Call NavTrace("Goto success")
    SelectFormulaReference = True
    Exit Function

CleanFail:
    Call NavTrace("Goto failed for: " & token)
    Err.Clear
End Function

Private Function TryResolveNameRange(ByVal token As String, ByRef target As Range) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nm As Name

    If Len(mNavStartWorkbookName) > 0 Then
        On Error Resume Next
        Set wb = Workbooks(mNavStartWorkbookName)
        On Error GoTo 0
    End If

    If wb Is Nothing Then
        Set wb = ActiveWorkbook
    End If

    On Error Resume Next
    If Not wb Is Nothing Then
        Set nm = wb.Names.Item(token)
        If Err.Number = 0 And Not nm Is Nothing Then
            Set target = nm.RefersToRange
            If Not target Is Nothing Then
                TryResolveNameRange = True
                Exit Function
            End If
        End If
        Err.Clear
    End If
    On Error GoTo 0

    On Error Resume Next
    If Not wb Is Nothing Then
        If Len(mNavStartSheetName) > 0 Then
            Set ws = wb.Worksheets(mNavStartSheetName)
        Else
            Set ws = ActiveSheet
        End If
        If Not ws Is Nothing Then
            Set nm = ws.Names.Item(token)
            If Err.Number = 0 And Not nm Is Nothing Then
                Set target = nm.RefersToRange
                If Not target Is Nothing Then
                    TryResolveNameRange = True
                    Exit Function
                End If
            End If
        End If
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function TryParseQualifiedReference(ByVal token As String, _
                                            ByRef workbookName As String, _
                                            ByRef sheetName As String, _
                                            ByRef address As String) As Boolean
    Dim bangPos As Long
    Dim leftPart As String
    Dim cleaned As String
    Dim wbStart As Long
    Dim wbEnd As Long

    bangPos = InStrRev(token, "!")
    If bangPos = 0 Then Exit Function

    leftPart = Left$(token, bangPos - 1)
    address = Mid$(token, bangPos + 1)
    If Len(address) = 0 Then Exit Function

    cleaned = leftPart
    If Len(cleaned) >= 2 And Left$(cleaned, 1) = "'" And Right$(cleaned, 1) = "'" Then
        cleaned = Mid$(cleaned, 2, Len(cleaned) - 2)
        cleaned = Replace(cleaned, "''", "'")
    End If

    wbStart = InStr(cleaned, "[")
    wbEnd = InStr(cleaned, "]")
    If wbStart > 0 And wbEnd > wbStart Then
        workbookName = Mid$(cleaned, wbStart + 1, wbEnd - wbStart - 1)
        sheetName = Mid$(cleaned, wbEnd + 1)
    Else
        sheetName = cleaned
    End If

    If Len(sheetName) = 0 Then Exit Function
    TryParseQualifiedReference = True
End Function

Private Function BuildQualifiedReference(ByVal token As String) As String
    If InStr(token, "!") > 0 Then
        BuildQualifiedReference = token
        Exit Function
    End If

    If Not IsPlainAddressToken(token) Then
        BuildQualifiedReference = token
        Exit Function
    End If

    If Len(mNavStartSheetName) = 0 Then
        BuildQualifiedReference = token
        Exit Function
    End If

    BuildQualifiedReference = QuoteSheetReference(mNavStartWorkbookName, mNavStartSheetName) & "!" & token
End Function

Private Function IsPlainAddressToken(ByVal token As String) As Boolean
    Static reg As Object
    If reg Is Nothing Then
        Set reg = CreateObject("VBScript.RegExp")
        reg.Pattern = "^\$?[A-Za-z]{1,3}\$?\d{1,7}(:\$?[A-Za-z]{1,3}\$?\d{1,7})?$|^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$|^\$?\d{1,7}:\$?\d{1,7}$"
        reg.IgnoreCase = True
    End If

    IsPlainAddressToken = reg.Test(token)
End Function

Private Function QuoteSheetReference(ByVal workbookName As String, ByVal sheetName As String) As String
    Dim fullName As String
    fullName = sheetName

    If Len(workbookName) > 0 Then
        fullName = "[" & workbookName & "]" & fullName
    End If

    If InStr(fullName, " ") > 0 Or InStr(fullName, "'") > 0 Then
        fullName = Replace(fullName, "'", "''")
        fullName = "'" & fullName & "'"
    End If

    QuoteSheetReference = fullName
End Function
