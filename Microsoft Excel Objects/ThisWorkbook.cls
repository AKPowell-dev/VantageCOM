Option Explicit
Public WithEvents App As Application

Public Sub EnsureAppHook()
    On Error Resume Next
    If App Is Nothing Then
        Set App = Application
    End If
    On Error GoTo 0
End Sub

Private Sub Workbook_Open()
    On Error Resume Next
    EnsureAppHook
    Application.EnableCancelKey = xlDisabled
    C_Core.ScheduleVimStartup
    InstallFormulaNavigatorHook
End Sub

Private Sub Workbook_Activate()
    EnsureAppHook
    InstallFormulaNavigatorHook
End Sub

Private Sub Workbook_AddinInstall()
    Workbook_Open
End Sub

Private Sub Workbook_AddinUninstall()
    On Error Resume Next
    UninstallFormulaNavigatorHook
    C_Core.ExitVim
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    UninstallFormulaNavigatorHook
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal target As Range)
    On Error GoTo Done
    If target Is Nothing Then Exit Sub
    If gVim Is Nothing Then Exit Sub
    If Not gVim.Enabled Then Exit Sub
    If gMacroSafeMode Then Exit Sub
    If Not IsActiveWorkbookSheet(target.Worksheet) Then Exit Sub
    If ShouldSkipAutoColor(target) Then Exit Sub

    If target.CountLarge = 1 Then
        Dim cellVal As Variant
        cellVal = target.Value2
        If IsEmpty(cellVal) Then Exit Sub
        If Not target.HasFormula Then
            If Not IsNumeric(cellVal) Then Exit Sub
        End If
    End If

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub

    engine.AutoColorRange target, 50
Done:
End Sub

Private Sub App_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
End Sub


Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo CleanExit
    If Target Is Nothing Then Exit Sub
    If gSuppressSelectionEvents Then Exit Sub
    If gVim Is Nothing Then Exit Sub
    If Not gVim.Enabled Then Exit Sub

    If Target.CountLarge > 5000 Then Exit Sub
    gSelectionStamp = gSelectionStamp + 1

CleanExit:
End Sub

Private Function ShouldSkipAutoColor(ByVal rng As Range) As Boolean
    On Error GoTo CleanExit
    If rng Is Nothing Then Exit Function
    If gMacroSafeMode Then
        ShouldSkipAutoColor = True
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = rng.Worksheet

    If rng.Columns.Count >= ws.Columns.Count Then
        ShouldSkipAutoColor = True
        Exit Function
    End If

    If rng.Rows.Count >= ws.Rows.Count Then
        ShouldSkipAutoColor = True
        Exit Function
    End If

    ' Skip auto-color during paste to keep Ctrl+V responsive.
    If Application.CutCopyMode <> 0 Then
        ShouldSkipAutoColor = True
        Exit Function
    End If

    ' Skip large changes to keep edits snappy.
    If rng.CountLarge > 50 Then
        ShouldSkipAutoColor = True
        Exit Function
    End If

CleanExit:
End Function


