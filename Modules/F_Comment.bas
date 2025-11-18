Attribute VB_Name = "F_Comment"
Option Explicit
Option Private Module

Function EditCellComment(Optional ByVal g As String) As Boolean
    Call RepeatRegister("EditCellComment")
    Call StopVisualMode

    If TypeName(Selection) = "Range" Then
        Call KeyStroke(Shift_ + F2_)
    End If
End Function

Function DeleteCellComment(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    Call RepeatRegister("DeleteCellComment")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.DeleteActiveCellComment

CleanExit:
    DeleteCellComment = False
    Exit Function

CleanFail:
    Call ErrorHandler("DeleteCellComment")
    Resume CleanExit
End Function

Function DeleteCellCommentAll(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    If ActiveSheet.Comments.Count = 0 Then GoTo CleanExit

    If MsgBox(gVim.Msg.ConfirmToDeleteAllComments, vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
        GoTo CleanExit
    End If

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.DeleteAllComments

CleanExit:
    DeleteCellCommentAll = False
    Exit Function

CleanFail:
    Call ErrorHandler("DeleteCellCommentAll")
    Resume CleanExit
End Function

Function ToggleCellComment(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    Call RepeatRegister("ToggleCellComment")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.ToggleActiveCommentVisibility

CleanExit:
    ToggleCellComment = False
    Exit Function

CleanFail:
    Call ErrorHandler("ToggleCellComment")
    Resume CleanExit
End Function

Function HideCellComment(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    Call RepeatRegister("HideCellComment")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.HideActiveComment

CleanExit:
    HideCellComment = False
    Exit Function

CleanFail:
    Call ErrorHandler("HideCellComment")
    Resume CleanExit
End Function

Function ShowCellComment(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    Call RepeatRegister("ShowCellComment")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.ShowActiveComment

CleanExit:
    ShowCellComment = False
    Exit Function

CleanFail:
    Call ErrorHandler("ShowCellComment")
    Resume CleanExit
End Function

Function ToggleCellCommentAll(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.ToggleAllCommentsVisibility

CleanExit:
    ToggleCellCommentAll = False
    Exit Function

CleanFail:
    Call ErrorHandler("ToggleCellCommentAll")
    Resume CleanExit
End Function

Function HideCellCommentAll(Optional ByVal g As String) As Boolean
    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    engine.SetCommentIndicatorMode 1
End Function

Function ShowCellCommentAll(Optional ByVal g As String) As Boolean
    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    engine.SetCommentIndicatorMode 2
End Function

Function HideCellCommentIndicator(Optional ByVal g As String) As Boolean
    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    engine.SetCommentIndicatorMode 0
End Function

Function NextComment(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.NavigateComments True, gVim.Count1

CleanExit:
    NextComment = False
    Exit Function

CleanFail:
    Call ErrorHandler("NextComment")
    Resume CleanExit
End Function

Function PrevComment(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.NavigateComments False, gVim.Count1

CleanExit:
    PrevComment = False
    Exit Function

CleanFail:
    Call ErrorHandler("PrevComment")
    Resume CleanExit
End Function
