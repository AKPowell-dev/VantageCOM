Attribute VB_Name = "F_UsefulCmd"
Option Explicit
Option Private Module

Private Const ZOOM_STEP_PERCENT As Integer = 5 'Zoom step size in percent for manual zoom commands

Function Undo_CtrlZ(Optional ByVal g As String) As Boolean
    Dim times As Long
    Dim i As Long

    times = gVim.Count1
    If times < 1 Then times = 1

    On Error Resume Next
    For i = 1 To times
        Application.Undo
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        DoEvents
    Next i
    On Error GoTo 0

    gVim.Count1 = 1
    Undo_CtrlZ = False
End Function

Function RedoExecute(Optional ByVal g As String) As Boolean
    On Error Resume Next
    Application.CommandBars.ExecuteMso "Redo"
End Function

Function ToggleFreezePanes(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    ActiveWindow.FreezePanes = Not ActiveWindow.FreezePanes
    Exit Function

Catch:
    Call ErrorHandler("ToggleFreezePanes")
End Function

Function ZoomIn(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim afterZoomRate As Integer
    Dim zoomDelta As Integer

    zoomDelta = ZOOM_STEP_PERCENT
    If gVim.Count > 0 Then
        zoomDelta = gVim.Count * ZOOM_STEP_PERCENT
    End If

    afterZoomRate = ActiveWindow.Zoom + zoomDelta

    If afterZoomRate > 400 Then
        afterZoomRate = 400
    End If

    ActiveWindow.Zoom = afterZoomRate
    Exit Function

Catch:
    If ErrorHandler("ZoomIn") Then
        Call KeyStroke(Ctrl_ + Shift_ + Alt_ + Minus_)
    End If
End Function

Function ZoomOut(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim afterZoomRate As Integer
    Dim zoomDelta As Integer

    zoomDelta = ZOOM_STEP_PERCENT
    If gVim.Count > 0 Then
        zoomDelta = gVim.Count * ZOOM_STEP_PERCENT
    End If

    afterZoomRate = ActiveWindow.Zoom - zoomDelta

    If afterZoomRate < 10 Then
        afterZoomRate = 10
    End If

    ActiveWindow.Zoom = afterZoomRate
    Exit Function

Catch:
    If ErrorHandler("ZoomOut") Then
        Call KeyStroke(Ctrl_ + Alt_ + Minus_)
    End If
End Function

Function ZoomSpecifiedScale(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim zoomScale As Integer

    Select Case gVim.Count1
        Case 1
            zoomScale = 100
        Case 2
            zoomScale = 25
        Case 3
            zoomScale = 55
        Case 4
            zoomScale = 85
        Case 5
            zoomScale = 130
        Case 6
            zoomScale = 160
        Case 7
            zoomScale = 200
        Case 8
            zoomScale = 400
        Case 9
            ActiveWindow.Zoom = True
            Exit Function
        Case Is > 400
            zoomScale = 400
        Case Is <= 400
            zoomScale = gVim.Count1
    End Select

    ActiveWindow.Zoom = zoomScale
    Exit Function

Catch:
    Call ErrorHandler("ZoomSpecifiedScale")
End Function

Function ToggleFormulaBar(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    Application.DisplayFormulaBar = Not Application.DisplayFormulaBar
    Exit Function

Catch:
    Call ErrorHandler("ToggleFormulaBar")
End Function

Function ToggleGridlines(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    ActiveWindow.DisplayGridlines = Not ActiveWindow.DisplayGridlines
    Exit Function

Catch:
    Call ErrorHandler("ToggleGridlines")
End Function

Function ToggleHeadings(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    ActiveWindow.DisplayHeadings = Not ActiveWindow.DisplayHeadings
    Exit Function

Catch:
    Call ErrorHandler("ToggleHeadings")
End Function

Function ShowSummaryInfo(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    Application.Dialogs(xlDialogSummaryInfo).Show
    Exit Function

Catch:
    Call ErrorHandler("ShowSummaryInfo")
End Function

Function SmartFillColor(Optional ByVal g As String) As Boolean
    Call StopVisualMode

    Dim colorInfo As cls_FontColor
    Set colorInfo = UF_ColorPicker.Launch()
    If colorInfo Is Nothing Then Exit Function

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    engine.ApplySmartFillColor colorInfo.IsNull, colorInfo.IsThemeColor, colorInfo.ThemeColor, colorInfo.TintAndShade, colorInfo.Color
End Function

Function SmartFontColor(Optional ByVal g As String) As Boolean
    Call StopVisualMode

    Dim colorInfo As cls_FontColor
    Dim applied As Boolean

    Set colorInfo = UF_ColorPicker.Launch()
    If colorInfo Is Nothing Then Exit Function

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    applied = engine.ApplySmartFontColor(colorInfo.IsNull, colorInfo.IsThemeColor, colorInfo.ThemeColor, colorInfo.ObjectThemeColor, colorInfo.TintAndShade, colorInfo.Color)
    If Not applied Then
        Call SetStatusBarTemporarily("No text found to recolor.", 2000)
    End If

    SmartFontColor = False
End Function

Function ShowContextMenu(Optional ByVal g As String) As Boolean
    'Send Shift+F10
    Call KeyStroke(Shift_ + F10_)
End Function

Function ShowMacroDialog(Optional ByVal g As String) As Boolean
    'Send Alt+F8
    Call KeyStroke(Alt_ + F8_, Tab_)
End Function

Function SetPrintArea(Optional ByVal g As String) As Boolean
    Call StopVisualMode

    'Send Alt, P, R, S
    Call KeyStroke(Alt_, P_, R_, S_)
End Function

Function ClearPrintArea(Optional ByVal g As String) As Boolean
    Call StopVisualMode

    'Send Alt, P, R, C
    Call KeyStroke(Alt_, P_, R_, C_)
End Function

Function Sort(Optional ByVal sortOrder As XlSortOrder) As Boolean
    Call StopVisualMode

    If sortOrder = xlAscending Then
        'Send Alt, A, S, A
        Call KeyStroke(Alt_, A_, S_, A_)
    Else
        'Send Alt, A, S, D
        Call KeyStroke(Alt_, A_, S_, D_)
    End If
End Function

Function RemoveDuplicates(Optional ByVal g As String) As Boolean
    Call StopVisualMode

    'Send Alt, A, M, Shift + Tab, Shift + Tab
    Call KeyStroke(Alt_, A_, M_, Shift_ + Tab_, Shift_ + Tab_)
End Function

Function OpenActiveBookDir(Optional ByVal g As String) As Boolean
    On Error GoTo Catch
    ActiveWorkbook.FollowHyperlink ActiveWorkbook.Path
    Exit Function

Catch:
    Call ErrorHandler("OpenActiveBookDir")
End Function

Function YankActiveBookPath(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    'Set to clipboard
    With New DataObject
        .SetText ActiveWorkbook.FullName
        .PutInClipboard
    End With

    Call SetStatusBarTemporarily(gVim.Msg.YankDone & " (" & ActiveWorkbook.FullName & ")", 3000)
    Exit Function

Catch:
    Call ErrorHandler("YankActiveBookPath")
End Function
