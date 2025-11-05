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
    Call KeyStroke(Ctrl_ + V_)
End Function

Private Function PasteRows(ByVal PasteDirection As XlSearchDirection)
    On Error GoTo Catch

    Dim yankedRows As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim copies As Long

    If Not IsRangeValid(gVim.Vars.LastYanked) Then Exit Function

    copies = gVim.Count1
    If copies < 1 Then copies = 1

    yankedRows = gVim.Vars.LastYanked.Rows.Count
    startRow = ActiveCell.Row + IIf(PasteDirection = xlNext, 1, 0)
    If startRow < 1 Then startRow = 1

    With ActiveSheet
        If startRow > .Rows.Count Then Exit Function

        endRow = startRow + yankedRows * copies - 1
        If endRow > .Rows.Count Then
            endRow = .Rows.Count
        End If
        If endRow < startRow Then Exit Function

        .Range(.Rows(startRow), .Rows(endRow)).Select

        Call KeyStroke(Ctrl_ + NumpadAdd_)
    End With

    If Application.CutCopyMode = xlCopy And IsRangeValid(gVim.Vars.LastYanked) Then
        gVim.Vars.LastYanked.Copy
    End If
    Exit Function

Catch:
    Call ErrorHandler("PasteRows")
End Function

Private Function PasteColumns(ByVal PasteDirection As XlSearchDirection)
    On Error GoTo Catch

    Dim yankedColumns As Long
    Dim startColumn As Long
    Dim endColumn As Long
    Dim copies As Long

    If Not IsRangeValid(gVim.Vars.LastYanked) Then Exit Function

    copies = gVim.Count1
    If copies < 1 Then copies = 1

    yankedColumns = gVim.Vars.LastYanked.Columns.Count
    startColumn = ActiveCell.Column + IIf(PasteDirection = xlNext, 1, 0)
    If startColumn < 1 Then startColumn = 1

    With ActiveSheet
        If startColumn > .Columns.Count Then Exit Function

        endColumn = startColumn + yankedColumns * copies - 1
        If endColumn > .Columns.Count Then
            endColumn = .Columns.Count
        End If
        If endColumn < startColumn Then Exit Function

        .Range(.Columns(startColumn), .Columns(endColumn)).Select

        Call KeyStroke(Ctrl_ + NumpadAdd_)
    End With

    If Application.CutCopyMode = xlCopy And IsRangeValid(gVim.Vars.LastYanked) Then
        gVim.Vars.LastYanked.Copy
    End If
    Exit Function

Catch:
    Call ErrorHandler("PasteColumns")
End Function

Function PasteValue(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("PasteValue")
    Call StopVisualMode

    Dim cb As Variant
    Dim cbType As Integer

    cb = Application.ClipboardFormats

    If cb(1) = -1 Then
        Exit Function
    End If
    cbType = cb(2)

    If Application.CutCopyMode > 0 Then 'Cells
        Call KeyStroke(Alt_, H_, V_, V_)

    Else
        Select Case cbType
            Case xlClipboardFormatText
                Call KeyStroke(Ctrl_ + V_)
            Case xlClipboardFormatRTF
                Call KeyStroke(Alt_, H_, V_, T_)
            Case xlHtml
                Call KeyStroke(Alt_, H_, V_, S_, End_, Enter_)
            Case Else
                Call DebugPrint("Unknown ClipboardType: " & cbType, "PasteValue")
        End Select
    End If
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
