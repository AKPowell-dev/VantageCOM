Attribute VB_Name = "F_Moving"
Option Explicit
Option Private Module

Private gFastRepeatActive As Boolean
Private gFastRepeatKey As Long

Function MoveUp()
    Dim r As Long
    Dim prevSuppress As Boolean
    prevSuppress = gSuppressSelectionEvents
    gSuppressSelectionEvents = True
    On Error GoTo CleanExit

    If gVim.Count1 = 1 Then
        keybd_event vbKeyUp, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyUp, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        r = ActiveCell.Row - gVim.Count1
        If r < 1 Then
            r = 1
        End If
        ActiveSheet.Cells(r, ActiveCell.Column).Select
    End If

    gSelectionStamp = gSelectionStamp + 1

CleanExit:
    gSuppressSelectionEvents = prevSuppress
    On Error GoTo 0
End Function

Function MoveDown()
    Dim r As Long
    Dim prevSuppress As Boolean
    prevSuppress = gSuppressSelectionEvents
    gSuppressSelectionEvents = True
    On Error GoTo CleanExit

    If gVim.Count1 = 1 Then
        keybd_event vbKeyDown, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyDown, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        r = ActiveCell.Row + gVim.Count1
        If r > ActiveSheet.Rows.Count Then
            r = ActiveSheet.Rows.Count
        End If
        ActiveSheet.Cells(r, ActiveCell.Column).Select
    End If

    gSelectionStamp = gSelectionStamp + 1

CleanExit:
    gSuppressSelectionEvents = prevSuppress
    On Error GoTo 0
End Function

Function MoveLeft()
    Dim c As Long
    Dim prevSuppress As Boolean
    prevSuppress = gSuppressSelectionEvents
    gSuppressSelectionEvents = True
    On Error GoTo CleanExit

    If gVim.Count1 = 1 Then
        keybd_event vbKeyLeft, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyLeft, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        c = ActiveCell.Column - gVim.Count1
        If c < 1 Then
            c = 1
        End If
        ActiveSheet.Cells(ActiveCell.Row, c).Select
    End If

    gSelectionStamp = gSelectionStamp + 1

CleanExit:
    gSuppressSelectionEvents = prevSuppress
    On Error GoTo 0
End Function

Function MoveRight()
    Dim c As Long
    Dim prevSuppress As Boolean
    prevSuppress = gSuppressSelectionEvents
    gSuppressSelectionEvents = True
    On Error GoTo CleanExit

    If gVim.Count1 = 1 Then
        keybd_event vbKeyRight, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyRight, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        c = ActiveCell.Column + gVim.Count1
        If c > ActiveSheet.Columns.Count Then
            c = ActiveSheet.Columns.Count
        End If
        ActiveSheet.Cells(ActiveCell.Row, c).Select
    End If

    gSelectionStamp = gSelectionStamp + 1

CleanExit:
    gSuppressSelectionEvents = prevSuppress
    On Error GoTo 0
End Function

Private Sub ExtendSelectionKey(ByVal keyCode As Long, Optional ByVal steps As Long = 1)
    Dim i As Long
    Dim prevSuppress As Boolean
    Dim shiftAlreadyDown As Boolean

    If steps < 1 Then steps = 1

    prevSuppress = gSuppressSelectionEvents
    gSuppressSelectionEvents = True
    On Error GoTo CleanExit

    shiftAlreadyDown = IsShiftPhysicallyDown()
    If Not shiftAlreadyDown Then
        keybd_event vbKeyShift, 0, 0, 0
    End If
    For i = 1 To steps
        keybd_event keyCode, 0, EXTENDED_KEY Or 0, 0
        keybd_event keyCode, 0, EXTENDED_KEY Or KEYUP, 0
    Next i
    If Not shiftAlreadyDown Then
        keybd_event vbKeyShift, 0, KEYUP, 0
    End If

    gSelectionStamp = gSelectionStamp + 1

CleanExit:
    gSuppressSelectionEvents = prevSuppress
    On Error GoTo 0
End Sub

Public Sub FastRepeatMove(ByVal arrowKey As Long, ByVal repeatKey As Long)
    Const INITIAL_DELAY_MS As Long = 160
    Const REPEAT_DELAY_MS As Long = 8
    Const YIELD_INTERVAL As Long = 4

    Dim prevSuppress As Boolean
    Dim prevEvents As Boolean
    Dim startTime As Double
    Dim iterations As Long

    If gFastRepeatActive Then
        If gFastRepeatKey = repeatKey And IsVirtualKeyDown(repeatKey) Then
            Exit Sub
        End If
    End If
    gFastRepeatActive = True
    gFastRepeatKey = repeatKey

    prevSuppress = gSuppressSelectionEvents
    prevEvents = Application.EnableEvents
    gSuppressSelectionEvents = True
    Application.EnableEvents = False
    On Error GoTo CleanExit

    ' First move happens immediately.
    keybd_event arrowKey, 0, EXTENDED_KEY Or 0, 0
    keybd_event arrowKey, 0, EXTENDED_KEY Or KEYUP, 0

    If Not IsVirtualKeyDown(repeatKey) Then GoTo CleanExit

    startTime = Timer
    Do While IsVirtualKeyDown(repeatKey)
        If ElapsedMillis(startTime) >= INITIAL_DELAY_MS Then Exit Do
        DoEvents
    Loop

    Do While IsVirtualKeyDown(repeatKey)
        keybd_event arrowKey, 0, EXTENDED_KEY Or 0, 0
        keybd_event arrowKey, 0, EXTENDED_KEY Or KEYUP, 0
        iterations = iterations + 1
        If REPEAT_DELAY_MS > 0 Then Sleep REPEAT_DELAY_MS
        If (iterations Mod YIELD_INTERVAL) = 0 Then DoEvents
    Loop

CleanExit:
    gSelectionStamp = gSelectionStamp + 1
    Application.EnableEvents = prevEvents
    gSuppressSelectionEvents = prevSuppress
    gFastRepeatActive = False
    gFastRepeatKey = 0
    On Error GoTo 0
End Sub

Private Function IsVirtualKeyDown(ByVal keyCode As Long) As Boolean
    On Error Resume Next
    IsVirtualKeyDown = ((GetAsyncKeyState(keyCode) And &H8000) <> 0)
    On Error GoTo 0
End Function

Private Function ElapsedMillis(ByVal startTime As Double) As Long
    Dim t As Double
    t = Timer
    If t < startTime Then t = t + 86400#
    ElapsedMillis = CLng((t - startTime) * 1000#)
End Function

Private Function IsShiftPhysicallyDown() As Boolean
    On Error Resume Next
    IsShiftPhysicallyDown = ((GetKeyState(ShiftLeft_) And &H8000) <> 0) _
        Or ((GetKeyState(ShiftRight_) And &H8000) <> 0) _
        Or ((GetKeyState(vbKeyShift) And &H8000) <> 0)
    On Error GoTo 0
End Function

Private Function ResizeInner(Optional Up As Long = 0, _
                             Optional Down As Long = 0, _
                             Optional Left As Long = 0, _
                             Optional Right As Long = 0)
    On Error GoTo Catch

    Dim r As Long
    Dim c As Long
    Dim firstRow As Long
    Dim firstColumn As Long
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim screenTop As Long
    Dim screenBottom As Long
    Dim screenLeft As Long
    Dim screenRight As Long

    Dim actCell As Range
    Dim baseRange As Range

    '?l??????
    r = Selection.Rows.Count
    c = Selection.Columns.Count

    firstRow = Selection.item(1).Row
    firstColumn = Selection.item(1).Column
    lastRow = Selection(Selection.Count).Row
    lastColumn = Selection(Selection.Count).Column

    screenTop = ActiveWindow.VisibleRange.item(1).Row
    screenBottom = ActiveWindow.VisibleRange.item(ActiveWindow.VisibleRange.Count).Row - 1
    screenLeft = ActiveWindow.VisibleRange.item(1).Column
    screenRight = ActiveWindow.VisibleRange.item(ActiveWindow.VisibleRange.Count).Column - 1

    '?Z????????????
    Set actCell = ActiveCell
    Set baseRange = Selection

    '?????z???????????v?Z
    If Up < 0 And -Up >= r Then
        Down = -(r + Up) + 1
        Up = 0
        Set baseRange = baseRange.offset(rowOffset:=r - 1).Resize(RowSize:=1)
    ElseIf Down < 0 And -Down >= r Then
        Up = -(r + Down) + 1
        Down = 0
        Set baseRange = baseRange.Resize(RowSize:=1)
    ElseIf Left < 0 And -Left >= c Then
        Right = -(c + Left) + 1
        Left = 0
        Set baseRange = baseRange.offset(ColumnOffset:=c - 1).Resize(ColumnSize:=1)
    ElseIf Right < 0 And -Right >= c Then
        Left = -(c + Right) + 1
        Right = 0
        Set baseRange = baseRange.Resize(ColumnSize:=1)
    End If

    '???E???z???????????}????
    If Up > 0 And firstRow <= Up Then
        Up = firstRow - 1
    ElseIf Down > 0 And lastRow + Down > ActiveSheet.Rows.Count Then
        Down = ActiveSheet.Rows.Count - lastRow
    ElseIf Left > 0 And firstColumn <= Left Then
        Left = firstColumn - 1
    ElseIf Right > 0 And lastColumn + Right > ActiveSheet.Columns.Count Then
        Right = ActiveSheet.Columns.Count - lastColumn
    End If

    '????????????????????
    If Up <> 0 Then
        baseRange.offset(rowOffset:=-Up).Resize(RowSize:=baseRange.Rows.Count + Up).Select
        SafeActivateRange actCell

        ActiveWindow.ScrollRow = screenTop
        ActiveWindow.ScrollColumn = screenLeft

        If screenTop > firstRow - Up Then
            ActiveWindow.SmallScroll Up:=screenTop - (firstRow - Up)
        ElseIf screenBottom < firstRow - Up Then
            ActiveWindow.SmallScroll Down:=(firstRow - Up) - screenBottom
        End If

    '????????????????????
    ElseIf Down <> 0 Then
        baseRange.Resize(RowSize:=baseRange.Rows.Count + Down).Select
        SafeActivateRange actCell

        ActiveWindow.ScrollRow = screenTop
        ActiveWindow.ScrollColumn = screenLeft

        If screenTop > lastRow + Down Then
            ActiveWindow.SmallScroll Up:=screenTop - (lastRow + Down)
        ElseIf screenBottom < lastRow + Down Then
            ActiveWindow.SmallScroll Down:=(lastRow + Down) - screenBottom
        End If

    '????????????????????
    ElseIf Left <> 0 Then
        baseRange.offset(ColumnOffset:=-Left).Resize(ColumnSize:=baseRange.Columns.Count + Left).Select
        SafeActivateRange actCell

        ActiveWindow.ScrollRow = screenTop
        ActiveWindow.ScrollColumn = screenLeft

        If screenLeft > firstColumn - Left Then
            ActiveWindow.SmallScroll ToLeft:=screenLeft - (firstColumn - Left)
        ElseIf screenRight < firstColumn - Left Then
            ActiveWindow.SmallScroll ToRight:=(firstColumn - Left) - screenRight
        End If

    '?E??????????????????
    ElseIf Right <> 0 Then
        baseRange.Resize(ColumnSize:=baseRange.Columns.Count + Right).Select
        SafeActivateRange actCell

        ActiveWindow.ScrollRow = screenTop
        ActiveWindow.ScrollColumn = screenLeft

        If screenLeft > lastColumn + Right Then
            ActiveWindow.SmallScroll ToLeft:=screenLeft - (lastColumn + Right)
        ElseIf screenRight < lastColumn + Right Then
            ActiveWindow.SmallScroll ToRight:=(lastColumn + Right) - screenRight
        End If

    End If

    Set actCell = Nothing
    Set baseRange = Nothing

    Exit Function

Catch:
    Call ErrorHandler("ResizeInner")
End Function

Function MoveUpWithShift()
    Dim steps As Long
    steps = gVim.Count1
    If steps < 1 Then steps = 1
    Call ExtendSelectionKey(vbKeyUp, steps)
End Function

Function MoveDownWithShift()
    Dim steps As Long
    steps = gVim.Count1
    If steps < 1 Then steps = 1
    Call ExtendSelectionKey(vbKeyDown, steps)
End Function

Function MoveLeftWithShift()
    Dim steps As Long
    steps = gVim.Count1
    If steps < 1 Then steps = 1
    Call ExtendSelectionKey(vbKeyLeft, steps)
End Function

Function MoveRightWithShift()
    Dim steps As Long
    steps = gVim.Count1
    If steps < 1 Then steps = 1
    Call ExtendSelectionKey(vbKeyRight, steps)
End Function

Function MoveToFirstRow(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    With ActiveWorkbook.ActiveSheet
        If gVim.Count = 0 Then
            .Cells(1, ActiveCell.Column).Select
        Else
            .Cells(gVim.Count1, ActiveCell.Column).Select
        End If
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToFirstRow")
End Function

Function MoveToTopRow(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    With ActiveWorkbook.ActiveSheet
        If gVim.Count = 0 Then
            .Cells(.UsedRange.item(1).Row, ActiveCell.Column).Select
        Else
            .Cells(gVim.Count1, ActiveCell.Column).Select
        End If
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToTopRow")
End Function

Function MoveToLastRow(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    With ActiveWorkbook.ActiveSheet
        If gVim.Count = 0 Then
            .Cells(.UsedRange.item(.UsedRange.Count).Row, ActiveCell.Column).Select
        Else
            .Cells(gVim.Count1, ActiveCell.Column).Select
        End If
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToLastRow")
End Function

Function MoveToNthColumn(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    If gVim.Count1 > ActiveSheet.Columns.Count Then
        gVim.Count1 = ActiveSheet.Columns.Count
    End If

    ActiveSheet.Cells(ActiveCell.Row, gVim.Count1).Select
    Exit Function

Catch:
    Call ErrorHandler("MoveToNthColumn")
End Function

Function MoveToFirstColumn(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.Row, 1).Select
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToFirstColumn")
End Function

Function MoveToLeftEnd(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.Row, .UsedRange.item(1).Column).Select
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToLeftEnd")
End Function

Function MoveToRightEnd(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(ActiveCell.Row, .UsedRange.item(.UsedRange.Count).Column).Select
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToRightEnd")
End Function

Function MoveToTopOfCurrentRegion(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    Dim targetRow As Long

    With ActiveWorkbook.ActiveSheet
        targetRow = ActiveCell.CurrentRegion.item(1).Row
        If targetRow = ActiveCell.Row Then
            targetRow = ActiveCell.End(xlUp).Row
        End If

        .Cells(targetRow, ActiveCell.Column).Select
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToTopOfCurrentRegion")
End Function

Function MoveToBottomOfCurrentRegion(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    Dim targetRow As Long

    With ActiveWorkbook.ActiveSheet
        targetRow = ActiveCell.CurrentRegion.item(ActiveCell.CurrentRegion.Count).Row
        If .Cells(targetRow, ActiveCell.Column).MergeArea.Row = ActiveCell.Row Then
            targetRow = ActiveCell.End(xlDown).Row
        End If

        .Cells(targetRow, ActiveCell.Column).Select
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToBottomOfCurrentRegion")
End Function

Function MoveToA1(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    With ActiveWorkbook.ActiveSheet
        .Cells(1, 1).Select
    End With
    Exit Function

Catch:
    Call ErrorHandler("MoveToA1")
End Function

Function MoveToSpecifiedCell(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim jumpAddress As String
    jumpAddress = UF_CmdLine.Launch("Jump to: ", "Jump to", False)

    If jumpAddress = CMDLINE_CANCELED Or Len(jumpAddress) = 0 Then
        Exit Function
    End If

    Dim jumpTarget As Range: Set jumpTarget = Nothing
    If RegExpMatch(jumpAddress, "^[0-9]{1,7}$") Then
        Set jumpTarget = ActiveSheet.Cells(CInt(jumpAddress), ActiveCell.Column)

    ElseIf RegExpMatch(jumpAddress, "^[a-z]{1,3}$", isIgnoreCase:=True) Then
        Set jumpTarget = ActiveSheet.Range(jumpAddress & ActiveCell.Row)

    ElseIf RegExpMatch(jumpAddress, "^[a-z]{1,3}[0-9]{1,7}(:[a-z]{1,3}[0-9]{1,7})?$", isIgnoreCase:=True) Then
        Set jumpTarget = ActiveSheet.Range(jumpAddress)

    ElseIf RegExpMatch(jumpAddress, "^[a-z]{1,3}:[a-z]{1,3}$", isIgnoreCase:=True) Then
        Set jumpTarget = ActiveSheet.Range(jumpAddress)

    ElseIf RegExpMatch(jumpAddress, "[0-9]{1,7}:[0-9]{1,7}") Then
        Set jumpTarget = ActiveSheet.Range(jumpAddress)

    End If

    If Not jumpTarget Is Nothing Then
        Call RecordToJumpList
        jumpTarget.Select
        Set jumpTarget = Nothing
    End If
    Exit Function

Catch:
    Call ErrorHandler("MoveToSpecifiedCell")
End Function

Function MoveToSpecifiedRow(Optional ByVal lineNum As String) As Boolean
    On Error GoTo Catch

    ' Set default return value to True (= Waiting for an argument)
    MoveToSpecifiedRow = True

    lineNum = Trim(lineNum)
    If RegExpMatch(lineNum, "^0*[1-9][0-9]{0,9}$") Then
        Dim n As Long: n = CLng(Right(lineNum, 10))

        If n > ActiveSheet.Rows.Count Then
            n = ActiveSheet.Rows.Count
        End If

        Call RecordToJumpList

        ActiveSheet.Cells(CLng(n), ActiveCell.Column).Select
        MoveToSpecifiedRow = False  ' Set return value to False (= Done)
    End If

    Exit Function

Catch:
    Call ErrorHandler("MoveToSpecifiedRow")
End Function
