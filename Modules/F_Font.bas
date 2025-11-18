Attribute VB_Name = "F_Font"
Option Explicit
Option Private Module

Function IncreaseFontSize(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    Call RepeatRegister("IncreaseFontSize")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.IncreaseFontSize gVim.Count1

CleanExit:
    IncreaseFontSize = False
    Exit Function

CleanFail:
    Call ErrorHandler("IncreaseFontSize")
    Resume CleanExit
End Function

Function DecreaseFontSize(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    Call RepeatRegister("DecreaseFontSize")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.DecreaseFontSize gVim.Count1

CleanExit:
    DecreaseFontSize = False
    Exit Function

CleanFail:
    Call ErrorHandler("DecreaseFontSize")
    Resume CleanExit
End Function

Function ChangeFontName(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    Call StopVisualMode
    Call ShowFontDialogInternal

CleanExit:
    ChangeFontName = False
    Exit Function

CleanFail:
    Call ErrorHandler("ChangeFontName")
    Resume CleanExit
End Function

Function ChangeFontSize(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    Call StopVisualMode
    Call ShowFontDialogInternal

CleanExit:
    ChangeFontSize = False
    Exit Function

CleanFail:
    Call ErrorHandler("ChangeFontSize")
    Resume CleanExit
End Function

Function AlignLeft(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    Call RepeatRegister("AlignLeft")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.AlignLeft

CleanExit:
    AlignLeft = False
    Exit Function

CleanFail:
    Call ErrorHandler("AlignLeft")
    Resume CleanExit
End Function

Function AlignCenter(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    Call RepeatRegister("AlignCenter")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.AlignCenter

CleanExit:
    AlignCenter = False
    Exit Function

CleanFail:
    Call ErrorHandler("AlignCenter")
    Resume CleanExit
End Function

Function AlignRight(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    Call RepeatRegister("AlignRight")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.AlignRight

CleanExit:
    AlignRight = False
    Exit Function

CleanFail:
    Call ErrorHandler("AlignRight")
    Resume CleanExit
End Function

Function AlignTop(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    Call RepeatRegister("AlignTop")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.AlignTop

CleanExit:
    AlignTop = False
    Exit Function

CleanFail:
    Call ErrorHandler("AlignTop")
    Resume CleanExit
End Function

Function AlignMiddle(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    Call RepeatRegister("AlignMiddle")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.AlignMiddle

CleanExit:
    AlignMiddle = False
    Exit Function

CleanFail:
    Call ErrorHandler("AlignMiddle")
    Resume CleanExit
End Function

Function AlignBottom(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    Call RepeatRegister("AlignBottom")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.AlignBottom

CleanExit:
    AlignBottom = False
    Exit Function

CleanFail:
    Call ErrorHandler("AlignBottom")
    Resume CleanExit
End Function

Function ToggleBold(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    Call RepeatRegister("ToggleBold")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.ToggleBold

CleanExit:
    ToggleBold = False
    Exit Function

CleanFail:
    Call ErrorHandler("ToggleBold")
    Resume CleanExit
End Function

Function ToggleItalic(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    Call RepeatRegister("ToggleItalic")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.ToggleItalic

CleanExit:
    ToggleItalic = False
    Exit Function

CleanFail:
    Call ErrorHandler("ToggleItalic")
    Resume CleanExit
End Function

Function ToggleUnderline(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    Call RepeatRegister("ToggleUnderline")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.ToggleUnderline

CleanExit:
    ToggleUnderline = False
    Exit Function

CleanFail:
    Call ErrorHandler("ToggleUnderline")
    Resume CleanExit
End Function

Function ToggleStrikethrough(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    Call RepeatRegister("ToggleStrikethrough")
    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.ToggleStrikethrough

CleanExit:
    ToggleStrikethrough = False
    Exit Function

CleanFail:
    Call ErrorHandler("ToggleStrikethrough")
    Resume CleanExit
End Function

Function ChangeFormat(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    Call StopVisualMode

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    engine.ShowFormatNumberDialog

CleanExit:
    ChangeFormat = False
    Exit Function

CleanFail:
    Call ErrorHandler("ChangeFormat")
    Resume CleanExit
End Function

Function showFontDialog(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail
    Call StopVisualMode
    Call ShowFontDialogInternal

CleanExit:
    showFontDialog = False
    Exit Function

CleanFail:
    Call ErrorHandler("showFontDialog")
    Resume CleanExit
End Function

Private Sub ShowFontDialogInternal()
    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub
    engine.ShowFontDialog
End Sub

Function ChangeFontColor(Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    If TypeName(Selection) = "Nothing" Then
        Exit Function
    End If

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
    End If

    If Not resultColor Is Nothing Then
        Dim engine As Object
        Set engine = NetAddin()
        If engine Is Nothing Then Exit Function
        If engine.ApplyFontColor(resultColor.IsNull, resultColor.IsThemeColor, resultColor.ThemeColor, resultColor.ObjectThemeColor, resultColor.TintAndShade, resultColor.Color) Then
            Call RepeatRegister("ChangeFontColor", resultColor)
            Call StopVisualMode
            ChangeFontColor = True
        End If
    End If
    Exit Function

Catch:
    Call ErrorHandler("ChangeFontColor")
End Function
