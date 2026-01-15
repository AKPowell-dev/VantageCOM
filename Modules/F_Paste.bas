Attribute VB_Name = "F_Paste"
Option Explicit

Function PasteSmart(Optional ByVal PasteDirection As XlSearchDirection = xlNext) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("PasteSmart")
    Call StopVisualMode

    If Application.CutCopyMode = 0 Then 'Empty
        Set gVim.Vars.LastYanked = Nothing
    End If

    If Not IsRangeValid(gVim.Vars.LastYanked) Then
        Set gVim.Vars.LastYanked = Nothing
    End If

    If gVim.Vars.LastYanked Is Nothing Then
        Call Paste_CtrlV
        Exit Function
    End If

    If gVim.Vars.LastYanked.Rows.Count = gVim.Vars.LastYanked.Parent.Rows.Count Then
        Call PasteColumns(PasteDirection)
    ElseIf gVim.Vars.LastYanked.Columns.Count = gVim.Vars.LastYanked.Parent.Columns.Count Then
        Call PasteRows(PasteDirection)
    Else
        Call Paste_CtrlV
    End If
    Exit Function

Catch:
    Call ErrorHandler("PasteSmart")
End Function

Private Function Paste_CtrlV()
    On Error Resume Next
    Application.CommandBars.ExecuteMso "Paste"
    If Err.Number <> 0 Then
        Err.Clear
        Call KeyStroke(Ctrl_ + V_)
    End If
End Function

Private Function PasteRows(ByVal PasteDirection As XlSearchDirection)
    On Error GoTo Catch

    Dim copies As Long
    Dim engine As Object

    If Not IsRangeValid(gVim.Vars.LastYanked) Then Exit Function

    copies = gVim.Count1
    If copies < 1 Then copies = 1

    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function

    engine.PasteEntireRows gVim.Vars.LastYanked, copies, (PasteDirection = xlNext)
    Exit Function

Catch:
    Call ErrorHandler("PasteRows")
End Function

Private Function PasteColumns(ByVal PasteDirection As XlSearchDirection)
    On Error GoTo Catch

    Dim copies As Long
    Dim engine As Object

    If Not IsRangeValid(gVim.Vars.LastYanked) Then Exit Function

    copies = gVim.Count1
    If copies < 1 Then copies = 1

    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function

    engine.PasteEntireColumns gVim.Vars.LastYanked, copies, (PasteDirection = xlNext)
    Exit Function

Catch:
    Call ErrorHandler("PasteColumns")
End Function

Function PasteValue(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("PasteValue")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function

    engine.ClipboardPasteValuesSmart
    Exit Function

Catch:
    Call ErrorHandler("PasteValue")
End Function

Function PasteSpecial(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call StopVisualMode

    If Application.ClipboardFormats(1) = -1 Then
        Call SetStatusBarTemporarily(gVim.Msg.EmptyClipboard, 2000)
    Else
        On Error Resume Next
        Application.Dialogs(xlDialogPasteSpecial).Show
    End If
    Exit Function

Catch:
    Call ErrorHandler("PasteSpecial")
End Function

Function PasteExactNative(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("PasteExactNative")
    Call StopVisualMode

    Call Paste_CtrlV
    PasteExactNative = False
    Exit Function

Catch:
    Call ErrorHandler("PasteExactNative")
End Function

Function PasteValuesExact(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("PasteValuesExact")
    Call StopVisualMode

    On Error Resume Next
    Application.CommandBars.ExecuteMso "PasteValuesAndNumberFormats"
    If Err.Number <> 0 Then
        Err.Clear
        Application.CommandBars.ExecuteMso "PasteValues"
        If Err.Number <> 0 Then
            Err.Clear
            Call Paste_CtrlV
        End If
    End If
    On Error GoTo Catch

    PasteValuesExact = False
    Exit Function

Catch:
    Call ErrorHandler("PasteValuesExact")
End Function
