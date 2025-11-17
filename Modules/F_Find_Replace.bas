Attribute VB_Name = "F_Find_Replace"
Option Explicit
Option Private Module

Function ShowFindFollowLang(Optional ByVal g As String) As Boolean
    Dim searchStr As String
    searchStr = UF_CmdLine.Launch("/", "Find", gVim.IsJapanese)

    If searchStr <> CMDLINE_CANCELED Then
        Call FindInner(searchStr)
    End If
End Function

Function ShowFindNotFollowLang(Optional ByVal g As String) As Boolean
    Dim searchStr As String
    searchStr = UF_CmdLine.Launch("/", "Find", Not gVim.IsJapanese)

    If searchStr <> CMDLINE_CANCELED Then
        Call FindInner(searchStr)
    End If
End Function

Private Sub FindInner(ByVal findString As String)
    Dim t As Range

    If findString = "" Then
        Call NextFoundCell
        Exit Sub
    End If

    Set t = ActiveSheet.Cells.Find(What:=findString, _
                                   LookIn:=xlValues, _
                                   LookAt:=xlPart, _
                                   SearchOrder:=xlByColumns, _
                                   MatchByte:=False)
    If Not t Is Nothing Then
        Call RecordToJumpList
        SafeActivateRange t
    End If
End Sub

Function NextFoundCell(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim t As Range
    Dim i As Integer

    If gVim.Count1 > 1 Then
        Application.ScreenUpdating = False
    End If

    Call RecordToJumpList

    For i = gVim.Count1 To 1 Step -1
        If gVim.Count1 = 1 Then
            Application.ScreenUpdating = True
        End If

        Set t = Cells.FindNext(After:=ActiveCell)
        If Not t Is Nothing Then
            SafeActivateRange t
        Else
            Application.ScreenUpdating = True
            Exit Function
        End If

    Next i
    Exit Function

Catch:
    Call ErrorHandler("NextFoundCell")
End Function

Function PreviousFoundCell(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim t As Range
    Dim i As Integer

    If gVim.Count1 > 1 Then
        Application.ScreenUpdating = False
    End If

    Call RecordToJumpList

    For i = gVim.Count1 To 1 Step -1
        If i = 1 Then
            Application.ScreenUpdating = True
        End If

        Set t = Cells.FindPrevious(After:=ActiveCell)
        If Not t Is Nothing Then
            SafeActivateRange t
        Else
            Application.ScreenUpdating = True
            Exit Function
        End If

    Next i
    Exit Function

Catch:
    Call ErrorHandler("PreviousFoundCell")
End Function

Function ShowReplaceWindow(Optional ByVal g As String) As Boolean
    Call KeyStroke(Alt_, E_, E_)
End Function

Function FindActiveValueNext(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim t As Range
    Dim findText As String

    If ActiveCell Is Nothing Then
        Exit Function
    End If

    findText = ActiveCell.value

    If findText = "" Then
        Exit Function
    End If

    Set t = ActiveSheet.Cells.Find(What:=findText, _
                                   After:=ActiveCell, _
                                   LookIn:=xlValues, _
                                   LookAt:=xlPart, _
                                   SearchOrder:=xlByColumns, _
                                   MatchByte:=False)

    If Not t Is Nothing Then
        Call RecordToJumpList
        Call NextFoundCell
    End If

    Call SetStatusBarTemporarily("/" & findText, 2000, disablePrefix:=True)
    Exit Function

Catch:
    Call ErrorHandler("FindActiveValueNext")
End Function

Function FindActiveValuePrev(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim t As Range
    Dim findText As String

    If ActiveCell Is Nothing Then
        Exit Function
    End If

    findText = ActiveCell.value

    If findText = "" Then
        Exit Function
    End If

    Set t = ActiveSheet.Cells.Find(What:=findText, _
                                   After:=ActiveCell, _
                                   LookIn:=xlValues, _
                                   LookAt:=xlPart, _
                                   SearchOrder:=xlByColumns, _
                                   MatchByte:=False)

    If Not t Is Nothing Then
        Call RecordToJumpList
        Call PreviousFoundCell
    End If

    Call SetStatusBarTemporarily("?" & findText, 2000, disablePrefix:=True)
    Exit Function

Catch:
    Call ErrorHandler("FindActiveValuePrev")
End Function

Function NextSpecialCells(ByVal TypeValue As XlCellType, Optional SearchOrder As XlSearchOrder = xlByColumns) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    engine.NavigateSpecialCells TypeValue, SearchOrder, True, gVim.Count1

    Exit Function

Catch:
    If Err.Number = 1004 Then
        Call SetStatusBarTemporarily(gVim.Msg.NoMatchingCell, 2000)
    Else
        Call ErrorHandler("NextSpecialCells")
    End If
End Function

Function PrevSpecialCells(ByVal TypeValue As XlCellType, Optional SearchOrder As XlSearchOrder = xlByColumns) As Boolean
    On Error GoTo Catch

    Call RecordToJumpList

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    engine.NavigateSpecialCells TypeValue, SearchOrder, False, gVim.Count1
    Exit Function

Catch:
    If Err.Number = 1004 Then
        Call SetStatusBarTemporarily(gVim.Msg.NoMatchingCell, 2000)
    Else
        Call ErrorHandler("PrevSpecialCells")
    End If
End Function

End Function
