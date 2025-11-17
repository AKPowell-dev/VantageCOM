Attribute VB_Name = "F_Shapes"
'Attribute VB_Name = "F_Shapes"
Option Explicit
Option Private Module

Function ChangeShapeFillColor(Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    If VarType(Selection) <> vbObject Then
        Exit Function
    End If

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
    End If

    If resultColor Is Nothing Then Exit Function

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    engine.ApplyShapeFillColor resultColor.IsNull, resultColor.IsThemeColor, resultColor.ThemeColor, resultColor.TintAndShade, resultColor.Color

    Call RepeatRegister("ChangeShapeFillColor", resultColor)
    Exit Function

Catch:
    Call ErrorHandler("ChangeShapeFillColor")
End Function

Function ChangeShapeFontColor(Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    If VarType(Selection) <> vbObject Then
        Exit Function
    End If

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
    End If

    If Not resultColor Is Nothing Then
        Dim engine As Object
        Set engine = NetAddin()
        If engine Is Nothing Then Exit Function
        engine.ApplyShapeFontColor resultColor.IsNull, resultColor.IsThemeColor, resultColor.ThemeColor, resultColor.ObjectThemeColor, resultColor.TintAndShade, resultColor.Color
        Call RepeatRegister("ChangeShapeFontColor", resultColor)
        ChangeShapeFontColor = True
    End If

    Exit Function

Catch:
    Call ErrorHandler("ChangeShapeFontColor")
End Function

Function ChangeShapeBorderColor(Optional garbage As String, _
                                Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    If VarType(Selection) <> vbObject Then
        ChangeShapeBorderColor = True
        Exit Function
    End If

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
    End If

    If Not resultColor Is Nothing Then
        Dim engine As Object
        Set engine = NetAddin()
        If engine Is Nothing Then Exit Function
        engine.ApplyShapeBorderColor resultColor.IsNull, resultColor.IsThemeColor, resultColor.ThemeColor, resultColor.TintAndShade, resultColor.Color
        Call RepeatRegister("ChangeShapeBorderColor", "", resultColor)
    End If

    Exit Function

Catch:
    Call ErrorHandler("ChangeShapeBorderColor")
End Function

Function NextShape(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim i As Long
    Dim cnt As Long
    Dim shp As shape

    If VarType(Selection) = vbObject Then
        For i = 1 To gVim.Count1
            Call KeyStroke(Tab_)
        Next i
    Else
        cnt = ActiveSheet.Shapes.Count
        If cnt = 0 Then
            Exit Function
        End If
        ActiveSheet.Shapes((gVim.Count1 - 1) Mod cnt + 1).Select
    End If
    Exit Function

Catch:
    Call ErrorHandler("NextShape")
End Function

Function PrevShape(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim i As Long
    Dim cnt As Long
    Dim shp As shape

    If VarType(Selection) = vbObject Then
        For i = 1 To gVim.Count1
            Call KeyStroke(Shift_ + Tab_)
        Next i
    Else
        cnt = ActiveSheet.Shapes.Count
        If cnt = 0 Then
            Exit Function
        End If
        ActiveSheet.Shapes(cnt - (gVim.Count1 - 1) Mod cnt).Select
    End If
    Exit Function

Catch:
    Call ErrorHandler("PrevShape")
End Function


