Option Explicit
Public WithEvents App As Application
Private selectionPageBreakCache As Object

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
End Sub

Private Sub Workbook_Activate()
    EnsureAppHook
End Sub

Private Sub Workbook_AddinInstall()
    Workbook_Open
End Sub

Private Sub Workbook_AddinUninstall()
    On Error Resume Next
    C_Core.ExitVim
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal target As Range)
    On Error GoTo Done
    Dim c As Range
    If target Is Nothing Then Exit Sub
    If ShouldSkipAutoColor(target) Then Exit Sub

    ' ===== Optimization (batch color without status bar flicker) =====
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ' ================================================================

    For Each c In target.Cells
        Z_ColorAux.ColorCell c
    Next c

Done:
    ' ===== Restore states safely =====
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub App_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
End Sub


Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo CleanExit
    If Target Is Nothing Then Exit Sub

    gSelectionStamp = gSelectionStamp + 1

    If Target.CountLarge <= 10000 Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub

    If selectionPageBreakCache Is Nothing Then
        Set selectionPageBreakCache = CreateObject("Scripting.Dictionary")
    End If

    Dim ws As Worksheet
    Set ws = Sh
    Dim cacheKey As String
    cacheKey = ws.CodeName

    If selectionPageBreakCache.Exists(cacheKey) Then Exit Sub

    If Not ws.DisplayPageBreaks Then
        selectionPageBreakCache.Add cacheKey, True
        Exit Sub
    End If

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    ws.DisplayPageBreaks = False
    Application.EnableEvents = prevEvents

    selectionPageBreakCache.Add cacheKey, True

CleanExit:
End Sub

Private Function ShouldSkipAutoColor(ByVal rng As Range) As Boolean
    On Error GoTo CleanExit
    If rng Is Nothing Then Exit Function

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

    ' Skip heavy colorizing on large changes to keep Excel snappy
    If rng.CountLarge > 1024 Then
        ShouldSkipAutoColor = True
        Exit Function
    End If

    ' Multi-area ranges can be especially slow to iterate
    If rng.Areas.Count > 1 And rng.CountLarge > 256 Then
        ShouldSkipAutoColor = True
        Exit Function
    End If

CleanExit:
End Function


