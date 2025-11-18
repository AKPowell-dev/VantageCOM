Attribute VB_Name = "F_Moving"
Option Explicit
Option Private Module

Function MoveUp()
    On Error GoTo Catch

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Or IsShiftPressed() Then
        Dim r As Long
        If gVim.Count1 = 1 Then
            keybd_event vbKeyUp, 0, EXTENDED_KEY Or 0, 0
            keybd_event vbKeyUp, 0, EXTENDED_KEY Or KEYUP, 0
        Else
            r = ActiveCell.Row - gVim.Count1
            If r < 1 Then r = 1
            ActiveSheet.Cells(r, ActiveCell.Column).Select
        End If
    Else
        engine.MoveActiveCellBy -gVim.Count1, 0
    End If
    Exit Function

Catch:
    Call ErrorHandler("MoveUp")
End Function

Function MoveDown()
    On Error GoTo Catch

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Or IsShiftPressed() Then
        Dim r As Long
        If gVim.Count1 = 1 Then
            keybd_event vbKeyDown, 0, EXTENDED_KEY Or 0, 0
            keybd_event vbKeyDown, 0, EXTENDED_KEY Or KEYUP, 0
        Else
            r = ActiveCell.Row + gVim.Count1
            If r > ActiveSheet.Rows.Count Then r = ActiveSheet.Rows.Count
            ActiveSheet.Cells(r, ActiveCell.Column).Select
        End If
    Else
        engine.MoveActiveCellBy gVim.Count1, 0
    End If
    Exit Function

Catch:
    Call ErrorHandler("MoveDown")
End Function

Function MoveLeft()
    On Error GoTo Catch

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Or IsShiftPressed() Then
        Dim c As Long
        If gVim.Count1 = 1 Then
            keybd_event vbKeyLeft, 0, EXTENDED_KEY Or 0, 0
            keybd_event vbKeyLeft, 0, EXTENDED_KEY Or KEYUP, 0
        Else
            c = ActiveCell.Column - gVim.Count1
            If c < 1 Then c = 1
            ActiveSheet.Cells(ActiveCell.Row, c).Select
        End If
    Else
        engine.MoveActiveCellBy 0, -gVim.Count1
    End If
    Exit Function

Catch:
    Call ErrorHandler("MoveLeft")
End Function

Function MoveRight()
    On Error GoTo Catch

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Or IsShiftPressed() Then
        Dim c As Long
        If gVim.Count1 = 1 Then
            keybd_event vbKeyRight, 0, EXTENDED_KEY Or 0, 0
            keybd_event vbKeyRight, 0, EXTENDED_KEY Or KEYUP, 0
        Else
            c = ActiveCell.Column + gVim.Count1
            If c > ActiveSheet.Columns.Count Then c = ActiveSheet.Columns.Count
            ActiveSheet.Cells(ActiveCell.Row, c).Select
        End If
    Else
        engine.MoveActiveCellBy 0, gVim.Count1
    End If
    Exit Function

Catch:
    Call ErrorHandler("MoveRight")
End Function

Private Function ResizeInner(Optional Up As Long = 0, _
                             Optional Down As Long = 0, _
                             Optional Left As Long = 0, _
                             Optional Right As Long = 0)
    On Error GoTo CleanFail

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.ResizeSelection Up, Down, Left, Right
CleanExit:
    Exit Function
CleanFail:
    Call ErrorHandler("ResizeInner")
    Resume CleanExit
End Function
Function MoveUpWithShift()
    Dim r As Long
    If gVim.Count1 = 1 Then
        keybd_event vbKeyUp, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyUp, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        If TypeName(Selection) <> "Range" Then
            Exit Function
        End If

        If Selection.item(1).Row = ActiveCell.Row Then
            Call ResizeInner(Down:=-gVim.Count1)
        Else
            Call ResizeInner(Up:=gVim.Count1)
        End If
    End If
End Function

Function MoveDownWithShift()
    Dim r As Long
    If gVim.Count1 = 1 Then
        keybd_event vbKeyDown, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyDown, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        If TypeName(Selection) <> "Range" Then
            Exit Function
        End If

        If Selection.item(Selection.Count).Row = ActiveCell.Row Then
            Call ResizeInner(Up:=-gVim.Count1)
        Else
            Call ResizeInner(Down:=gVim.Count1)
        End If
    End If
End Function

Function MoveLeftWithShift()
    Dim c As Long
    If gVim.Count1 = 1 Then
        keybd_event vbKeyLeft, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyLeft, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        If TypeName(Selection) <> "Range" Then
            Exit Function
        End If

        If Selection.item(1).Column = ActiveCell.Column Then
            Call ResizeInner(Right:=-gVim.Count1)
        Else
            Call ResizeInner(Left:=gVim.Count1)
        End If
    End If
End Function

Function MoveRightWithShift()
    Dim c As Long
    If gVim.Count1 = 1 Then
        keybd_event vbKeyRight, 0, EXTENDED_KEY Or 0, 0
        keybd_event vbKeyRight, 0, EXTENDED_KEY Or KEYUP, 0
    Else
        If TypeName(Selection) <> "Range" Then
            Exit Function
        End If

        If Selection.item(Selection.Count).Column = ActiveCell.Column Then
            Call ResizeInner(Left:=-gVim.Count1)
        Else
            Call ResizeInner(Right:=gVim.Count1)
        End If
    End If
End Function

Private Function IsShiftPressed() As Boolean
    IsShiftPressed = ((GetAsyncKeyState(vbKeyShift) And &H8000) <> 0)
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
