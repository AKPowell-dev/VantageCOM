Attribute VB_Name = "F_Cell"
Option Explicit
Option Private Module

Private Const EXTENDED_SELECTION_ENABLED As Boolean = False

Private Enum eOperationType
    SelectOp
    YankOp
    CutOp
    DeleteOp
End Enum

Private Enum eSearchMode
    TopToBottom = 1
    LeftToRight
    BottomToTop
    RightToLeft
End Enum

Private Function OperateCells(ByRef target As Range, ByVal operationType As eOperationType)
    On Error GoTo Catch

    If target Is Nothing Then
        Exit Function
    End If

    Call StopVisualMode

    Select Case operationType
        Case eOperationType.SelectOp
            target.Select
        Case eOperationType.YankOp
            target.Copy
            Set gVim.Vars.LastYanked = target
        Case eOperationType.CutOp
            target.Cut
            Set gVim.Vars.LastYanked = target
        Case eOperationType.DeleteOp
            target.Select
            Call KeyStroke(Ctrl_ + Minus_)
    End Select

Catch:
    Call ErrorHandler("OperateCells")
End Function

Function SelectUsedRange(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If ActiveSheet.Type <> XlSheetType.xlWorksheet Then
        Exit Function
    End If

    Call OperateCells(ActiveSheet.UsedRange, SelectOp)
    Exit Function

Catch:
    Call ErrorHandler("SelectUsedRange")
End Function

Function SelectAllCells(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If ActiveSheet.Type <> XlSheetType.xlWorksheet Then
        Exit Function
    End If

    Call OperateCells(ActiveSheet.Cells, SelectOp)
    Exit Function

Catch:
    Call ErrorHandler("SelectAllCells")
End Function

Function CutCell(Optional ByVal g As String) As Boolean
    Call StopVisualMode
    Call KeyStroke(Ctrl_ + X_)

    If TypeName(Selection) = "Range" Then
        Set gVim.Vars.LastYanked = Selection
    End If
End Function

Function CutUsedRange(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If ActiveSheet.Type <> XlSheetType.xlWorksheet Then
        Exit Function
    End If

    Call OperateCells(ActiveSheet.UsedRange, CutOp)
    Exit Function

Catch:
    Call ErrorHandler("CutUsedRange")
End Function

Function CutAllCells(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If ActiveSheet.Type <> XlSheetType.xlWorksheet Then
        Exit Function
    End If

    Call OperateCells(ActiveSheet.Cells, CutOp)
    Exit Function

Catch:
    Call ErrorHandler("CutAllCells")
End Function

Function YankCell(Optional ByVal g As String) As Boolean
    Call StopVisualMode
    Call KeyStroke(Ctrl_ + C_)

    If TypeName(Selection) = "Range" Then
        Set gVim.Vars.LastYanked = Selection
    End If
End Function

Function YankUsedRange(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If ActiveSheet.Type <> XlSheetType.xlWorksheet Then
        Exit Function
    End If

    Call OperateCells(ActiveSheet.UsedRange, YankOp)
    Exit Function

Catch:
    Call ErrorHandler("YankUsedRange")
End Function

Function YankAllCells(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If ActiveSheet.Type <> XlSheetType.xlWorksheet Then
        Exit Function
    End If

    Call OperateCells(ActiveSheet.Cells, YankOp)
    Exit Function

Catch:
    Call ErrorHandler("YankAllCells")
End Function

Function DeleteUsedRange(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If ActiveSheet.Type <> XlSheetType.xlWorksheet Then
        Exit Function
    End If

    Call OperateCells(ActiveSheet.UsedRange, DeleteOp)
    Exit Function

Catch:
    Call ErrorHandler("DeleteUsedRange")
End Function

Function DeleteAllCells(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If ActiveSheet.Type <> XlSheetType.xlWorksheet Then
        Exit Function
    End If

    Call OperateCells(ActiveSheet.Cells, DeleteOp)
    Exit Function

Catch:
    Call ErrorHandler("DeleteAllCells")
End Function

Function YankFromUpCell(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    Call RepeatRegister("YankFromUpCell")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.FillFromAbove

CleanExit:
    YankFromUpCell = False
    Exit Function

CleanFail:
    Call ErrorHandler("YankFromUpCell")
    Resume CleanExit
End Function

Function YankFromDownCell(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    Call RepeatRegister("YankFromDownCell")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.FillFromBelow

CleanExit:
    YankFromDownCell = False
    Exit Function

CleanFail:
    Call ErrorHandler("YankFromDownCell")
    Resume CleanExit
End Function

Function YankFromLeftCell(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    Call RepeatRegister("YankFromLeftCell")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.FillFromLeft

CleanExit:
    YankFromLeftCell = False
    Exit Function

CleanFail:
    Call ErrorHandler("YankFromLeftCell")
    Resume CleanExit
End Function

Function YankFromRightCell(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    Call RepeatRegister("YankFromRightCell")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.FillFromRight

CleanExit:
    YankFromRightCell = False
    Exit Function

CleanFail:
    Call ErrorHandler("YankFromRightCell")
    Resume CleanExit
End Function

Function YankAsPlaintext(Optional ByVal ColumnSpliter As String = vbTab) As Boolean
    On Error GoTo CleanFail

    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.CopySelectionAsPlainText ColumnSpliter

CleanExit:
    YankAsPlaintext = False
    Exit Function

CleanFail:
    Call ErrorHandler("YankAsPlaintext")
    Resume CleanExit
End Function

Function IncrementText(Optional ByVal g As String) As Boolean
    Call RepeatRegister("IncrementText")
    Call StopVisualMode

    Dim i As Integer

    For i = 1 To gVim.Count1
        Call KeyStroke(Alt_, H_, k6_)
    Next i
End Function

Function DecrementText(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DecrementText")
    Call StopVisualMode

    Dim i As Integer

    For i = 1 To gVim.Count1
        Call KeyStroke(Alt_, H_, k5_)
    Next i
End Function

Function IncreaseDecimal(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    Call RepeatRegister("IncreaseDecimal")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.IncreaseDecimalPlaces gVim.Count1

CleanExit:
    IncreaseDecimal = False
    Exit Function

CleanFail:
    Call ErrorHandler("IncreaseDecimal")
    Resume CleanExit
End Function

Function DecreaseDecimal(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    Call RepeatRegister("DecreaseDecimal")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.DecreaseDecimalPlaces gVim.Count1

CleanExit:
    DecreaseDecimal = False
    Exit Function

CleanFail:
    Call ErrorHandler("DecreaseDecimal")
    Resume CleanExit
End Function


Function AddNumber(Optional ByVal g As String) As Boolean
    Call RepeatRegister("AddNumber")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    engine.AdjustNumbers gVim.Count1, False, False
End Function

Function SubtractNumber(Optional ByVal g As String) As Boolean
    Call RepeatRegister("SubtractNumber")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    engine.AdjustNumbers gVim.Count1, True, False
End Function

Function VisualAddNumber(Optional ByVal g As String) As Boolean
    Call RepeatRegister("VisualAddNumber")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    engine.AdjustNumbers gVim.Count1, False, True
End Function

Function VisualSubtractNumber(Optional ByVal g As String) As Boolean
    Call RepeatRegister("VisualSubtractNumber")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    engine.AdjustNumbers gVim.Count1, True, True
End Function

Function InsertCellsUp(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("InsertCellsUp")
    Call StopVisualMode

    Application.ScreenUpdating = False
    If gVim.Count1 > 1 Then
        Selection.Resize(gVim.Count1, Selection.Columns.Count).Select
    End If

    Call KeyStroke(Ctrl_ + Shift_ + Semicoron_JIS_, D_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("InsertCellsUp")
End Function

Function InsertCellsDown(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("InsertCellsDown")
    Call StopVisualMode

    Application.ScreenUpdating = False
    If Selection.Row < ActiveSheet.Rows.Count Then
        Selection.offset(1, 0).Select
    End If

    If gVim.Count1 > 1 Then
        Selection.Resize(gVim.Count1, Selection.Columns.Count).Select
    End If

    Call KeyStroke(Ctrl_ + Shift_ + Semicoron_JIS_, D_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("InsertCellsDown")
End Function

Function InsertCellsLeft(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("InsertCellsLeft")
    Call StopVisualMode

    Application.ScreenUpdating = False
    If gVim.Count1 > 1 Then
        Selection.Resize(Selection.Rows.Count, gVim.Count1).Select
    End If

    Call KeyStroke(Ctrl_ + Shift_ + Semicoron_JIS_, I_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("InsertCellsLeft")
End Function

Function InsertCellsRight(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("InsertCellsRight")
    Call StopVisualMode

    Application.ScreenUpdating = False
    If Selection.Column < ActiveSheet.Columns.Count Then
        Selection.offset(0, 1).Select
    End If

    If gVim.Count1 > 1 Then
        Selection.Resize(Selection.Rows.Count, gVim.Count1).Select
    End If

    Call KeyStroke(Ctrl_ + Shift_ + Semicoron_JIS_, I_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("InsertCellsRight")
End Function

Function DeleteValue(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DeleteValue")
    Call StopVisualMode
    Call KeyStroke(Delete_)
End Function

Function DeleteToUp(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("DeleteToUp")
    Call StopVisualMode

    Application.ScreenUpdating = False
    If gVim.Count1 > 1 Then
        Selection.Resize(gVim.Count1, Selection.Columns.Count).Select
    End If

    Call KeyStroke(Ctrl_ + Minus_, U_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("DeleteToUp")
End Function

Function DeleteToLeft(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Call RepeatRegister("DeleteToLeft")
    Call StopVisualMode

    Application.ScreenUpdating = False
    If gVim.Count1 > 1 Then
        Selection.Resize(Selection.Rows.Count, gVim.Count1).Select
    End If

    Call KeyStroke(Ctrl_ + Minus_, L_, Enter_)

Catch:
    Application.ScreenUpdating = True
    Call ErrorHandler("DeleteToLeft")
End Function

Function ToggleWrapText(Optional ByVal g As String) As Boolean
    Call StopVisualMode
    Call KeyStroke(Alt_, H_, W_)
End Function

Function ToggleMergeCells(Optional ByVal g As String) As Boolean
    Call RepeatRegister("ToggleMergeCells")
    Call StopVisualMode

    If TypeName(Selection) = "Range" Then
        If Not ActiveCell.MergeCells And Selection.Count = 1 Then
            Exit Function
        End If

        If ActiveCell.MergeCells Then
            Call KeyStroke(Alt_, H_, M_, U_)
        Else
            Call KeyStroke(Alt_, H_, M_, M_)
        End If
    End If
End Function

Function ApplyCommaStyle(Optional ByVal g As String) As Boolean
    Call RepeatRegister("ApplyCommaStyle")
    Call StopVisualMode

    Call KeyStroke(Alt_, H_, K_)
End Function

Function ChangeInteriorColor(Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
    End If

    If Not resultColor Is Nothing Then
        Dim engine As Object
        Set engine = NetAddin()
        If engine Is Nothing Then Exit Function
        engine.ApplyInteriorColor resultColor.IsNull, resultColor.IsThemeColor, resultColor.ThemeColor, resultColor.TintAndShade, resultColor.Color

        Call RepeatRegister("ChangeInteriorColor", resultColor)
        Call StopVisualMode
    End If
    Exit Function

Catch:
    Call ErrorHandler("ChangeInteriorColor")
End Function

Function UnionSelectCells(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If Not EXTENDED_SELECTION_ENABLED Then
        On Error Resume Next
        If Not gVim Is Nothing Then Set gVim.Vars.ExtendRange = Nothing
        Exit Function
    End If

    Dim actCell As Range

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call StopVisualMode

    If gVim.Vars.ExtendRange Is Nothing Then
        Set gVim.Vars.ExtendRange = Selection

    ElseIf Not gVim.Vars.ExtendRange.Parent Is ActiveSheet Then
        Call SetStatusBarTemporarily(gVim.Msg.InitializedExtendedSelection, 2000)
        Set gVim.Vars.ExtendRange = Selection

    Else
        Set actCell = ActiveCell
        Set gVim.Vars.ExtendRange = Union2(gVim.Vars.ExtendRange, Selection)
        SafeSelectRange gVim.Vars.ExtendRange
        SafeActivateRange actCell

    End If
    Exit Function

Catch:
    If Err.Number = 424 Then
        Set gVim.Vars.ExtendRange = Selection
    Else
        Call ErrorHandler("UnionSelectCells")
    End If
End Function

Function ExceptSelectCells(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If Not EXTENDED_SELECTION_ENABLED Then
        On Error Resume Next
        If Not gVim Is Nothing Then Set gVim.Vars.ExtendRange = Nothing
        Exit Function
    End If

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call StopVisualMode

    If Not gVim.Vars.ExtendRange Is Nothing Then
        If Selection.Address = gVim.Vars.ExtendRange.Address Then
            Set gVim.Vars.ExtendRange = Except2(gVim.Vars.ExtendRange, ActiveCell)
        Else
            Set gVim.Vars.ExtendRange = Except2(gVim.Vars.ExtendRange, Selection)
        End If

        If Not gVim.Vars.ExtendRange Is Nothing Then
            gVim.Vars.ExtendRange.Select
        Else
            Call SetStatusBarTemporarily(gVim.Msg.ClearedExtendedSelection, 2000)
        End If
    End If
    Exit Function

Catch:
    If Err.Number = 424 Then
        Set gVim.Vars.ExtendRange = Nothing
    Else
        Call ErrorHandler("ExceptSelectCells")
    End If
End Function

Function ClearSelectCells(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If Not EXTENDED_SELECTION_ENABLED Then
        On Error Resume Next
        If Not gVim Is Nothing Then Set gVim.Vars.ExtendRange = Nothing
        Exit Function
    End If

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call StopVisualMode

    If Not gVim.Vars.ExtendRange Is Nothing Then
        If Selection.Address = gVim.Vars.ExtendRange.Address Then
            Set gVim.Vars.ExtendRange = Nothing
            Call SetStatusBarTemporarily(gVim.Msg.ClearedExtendedSelection, 2000)
            Exit Function
        End If
    End If

    If Selection.Columns.Count > 1 Or Selection.Rows.Count > 1 Or Selection.Areas.Count > 1 Then
        ActiveCell.Select
    ElseIf Not gVim.Vars.ExtendRange Is Nothing Then
        gVim.Vars.ExtendRange.Select
    End If
    Exit Function

Catch:
    If Err.Number = 424 Then
        Set gVim.Vars.ExtendRange = Nothing
    Else
        Call ErrorHandler("ClearSelectCells")
    End If
End Function

Function FollowHyperlinkOfActiveCell(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    With ActiveCell
        If .Hyperlinks.Count > 0 Then
            .Hyperlinks(1).Follow
        ElseIf .Formula <> .value And InStr(.Formula, "HYPERLINK") > 0 Then
            Dim linkAddr As String
            linkAddr = Application.Evaluate(Replace(.Formula, "HYPERLINK", "IFERROR"))
            ActiveWorkbook.FollowHyperlink linkAddr
        End If
    End With
    Exit Function

Catch:
    Call ErrorHandler("FollowHyperlinkOfActiveCell")
End Function

Function ChangeSelectedCells(ByVal value As String) As Boolean
    On Error GoTo Catch

    Call StopVisualMode

    If TypeName(Selection) = "Range" Then
        Selection.value = value
    ElseIf Not ActiveCell Is Nothing Then
        ActiveCell.value = value
    End If
    Exit Function

Catch:
    Call ErrorHandler("ChangeSelectedCells")
End Function

Function ApplyFlashFill(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call RepeatRegister("ApplyFlashFill")

    Selection.FlashFill

    Call StopVisualMode

    Exit Function
Catch:
    If Err.Number = 1004 Then
        Call ApplyAutoFillInner(fallback:=True)
    Else
        Call ErrorHandler("ApplyFlashFill")
    End If
End Function

Function ApplyAutoFill(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    ElseIf Selection.Count = 1 Then
        Exit Function
    End If

    Call RepeatRegister("ApplyAutoFill")

    Call ApplyAutoFillInner

    Exit Function

Catch:
    Call ErrorHandler("ApplyAutoFill")
End Function

Function ApplyAutoFillInner(Optional fallback As Boolean = False)
    On Error GoTo Catch

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    engine.ApplyAutoFill
    Call StopVisualMode
    Exit Function

Catch:
    Call ErrorHandler("ApplyAutoFillInner")
End Function

Private Function AutoSumInner(ByVal lastKey As Long)
    On Error GoTo Catch

    If TypeName(Selection) <> "Range" Then
        Exit Function
    End If

    Call KeyStroke(Alt_, M_, U_, lastKey)

    Exit Function
Catch:
    Call ErrorHandler("AutoSumInner")
End Function

Function AutoSum(Optional ByVal g As String) As Boolean
    Call AutoSumInner(S_)
End Function

Function AutoAverage(Optional ByVal g As String) As Boolean
    Call AutoSumInner(A_)
End Function

Function AutoCount(Optional ByVal g As String) As Boolean
    Call AutoSumInner(C_)
End Function

Function AutoMax(Optional ByVal g As String) As Boolean
    Call AutoSumInner(M_)
End Function

Function AutoMin(Optional ByVal g As String) As Boolean
    Call AutoSumInner(I_)
End Function

Function InsertFunction(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    Application.CommandBars.ExecuteMso "AutoSumMoreFunctions"
    Exit Function
Catch:
    Call ErrorHandler("InsertFunction")
End Function
