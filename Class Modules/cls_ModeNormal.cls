VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_ModeNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SELECTION_LOCK_ENABLED As Boolean = False

Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Private WithEvents AppForLock As Application
Attribute AppForLock.VB_VarHelpID = -1

Private Sub App_SheetChange(ByVal Sh As Object, ByVal target As Range)
    If gVim Is Nothing Then Exit Sub
    If Not gVim.Enabled Then Exit Sub
    If target Is Nothing Then Exit Sub
    If Not IsActiveWorkbookSheet(target.Worksheet) Then Exit Sub
    If target.CountLarge > 500 Then Exit Sub

    Call DisableIME
    Call RecordToJumpList(target)

    If gVim.Vars.FromInsertCmd Then
        ' Note: Undo command will no longer be available
        Call RepeatRegister("ChangeSelectedCells", target.Item(1).Formula)
    End If
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    If gVim.Enabled Then
        On Error Resume Next

        Unload UF_Cmd
        Unload UF_CmdLine
        Unload UF_ColorPicker
        Unload UF_SheetPicker

        On Error GoTo 0
    End If
End Sub

Public Sub ApplySelectionLock()
    If Not SELECTION_LOCK_ENABLED Then
        Call ClearSelectionLock
        Exit Sub
    End If

    If AppForLock Is Nothing Then
        Set AppForLock = Application
        Set gVim.Vars.LockedSheet = ActiveSheet
    End If
End Sub

Private Sub ClearSelectionLock()
    Set AppForLock = Nothing
    With gVim.Vars
        Set .LockedSheet = Nothing
        .SetLockedRows 0, 0
        .SetLockedColumns 0, 0
        Call SetStatusBar
    End With
End Sub

Private Sub AppForLock_SheetSelectionChange(ByVal Sh As Object, ByVal target As Range)
    On Error GoTo Catch

    If Not SELECTION_LOCK_ENABLED Then Exit Sub

    If Not gVim.Vars.LockedSheet Is Sh Then
        Call ClearSelectionLock
        Exit Sub
    End If

    With gVim.Vars
        Dim lockedRange As Range
        If .LockedRowsBegin > 0 And .LockedColumnsBegin > 0 Then
            Set lockedRange = .LockedSheet.Range(.LockedSheet.Cells(.LockedRowsBegin, .LockedColumnsBegin), .LockedSheet.Cells(.LockedRowsEnd, .LockedColumnsEnd))
        ElseIf .LockedRowsBegin = 0 And .LockedColumnsBegin > 0 Then
            Set lockedRange = .LockedSheet.Range(.LockedSheet.Columns(.LockedColumnsBegin), .LockedSheet.Columns(.LockedColumnsEnd))
        ElseIf .LockedRowsBegin > 0 And .LockedColumnsBegin = 0 Then
            Set lockedRange = .LockedSheet.Range(.LockedSheet.Rows(.LockedRowsBegin), .LockedSheet.Rows(.LockedRowsEnd))
        Else
            Call ClearSelectionLock
            Exit Sub
        End If

        Dim newRange As Range
        Set newRange = Intersect2(target, lockedRange)

        If Not newRange Is Nothing Then
            Application.EnableEvents = False
            newRange.Select
            Application.EnableEvents = True
            Exit Sub
        End If

        Dim newRow As Long
        Dim newCol As Long
        If .LockedColumnsBegin > 0 Then
            If .LockedColumnsBegin > target(target.Count).Column Then
                newCol = .LockedColumnsBegin
            ElseIf .LockedColumnsEnd < target(1).Column Then
                newCol = .LockedColumnsEnd
            End If
        End If

        If .LockedRowsBegin > 0 Then
            If .LockedRowsBegin > target(target.Count).Row Then
                newRow = .LockedRowsBegin
            ElseIf .LockedRowsEnd < target(1).Row Then
                newRow = .LockedRowsEnd
            End If
        End If

        If newRow > 0 And newCol > 0 Then
            Set newRange = .LockedSheet.Cells(newRow, newCol)
        ElseIf newRow > 0 And newCol = 0 Then
            Set newRange = Intersect2(target.EntireColumn, .LockedSheet.Rows(newRow))
        ElseIf newRow = 0 And newCol > 0 Then
            Set newRange = Intersect2(target.EntireRow, .LockedSheet.Columns(newCol))
        End If

        If Not newRange Is Nothing Then
            Application.EnableEvents = False
            newRange.Select
            Application.EnableEvents = True
        End If
    End With
    Exit Sub

Catch:
    Call ErrorHandler("ModeNormal.AppForLock_SheetSelectionChange")
End Sub

Private Sub AppForLock_SheetActivate(ByVal Sh As Object)
    If Not SELECTION_LOCK_ENABLED Then Exit Sub
    If Not gVim.Vars.LockedSheet Is Sh Then
        Call ClearSelectionLock
    End If
End Sub

Private Sub AppForLock_WorkbookDeactivate(ByVal Wb As Workbook)
    If Not SELECTION_LOCK_ENABLED Then Exit Sub
    If gVim.Vars.LockedSheet.Parent Is Wb Then
        Call ClearSelectionLock
    End If
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal target As Range)
    On Error GoTo Catch

    If gSuppressSelectionEvents Then Exit Sub
    If gVim Is Nothing Then Exit Sub
    If Not gVim.Enabled Then Exit Sub
    If Not target Is Nothing Then
        If target.CountLarge > 5000 Then Exit Sub
    End If

    gSelectionStamp = gSelectionStamp + 1
    Exit Sub

Catch:
    Call ErrorHandler("ModeNormal.App_SheetSelectionChange")
End Sub

Private Sub Class_Initialize()
    Set App = Application
    If Not gVim.Vars.LockedSheet Is Nothing Then
        If gVim.Vars.LockedSheet Is ActiveSheet Then
            Call ApplySelectionLock
        End If
    End If
End Sub

Private Sub Class_Terminate()
    Set App = Nothing
    Set AppForLock = Nothing
End Sub

