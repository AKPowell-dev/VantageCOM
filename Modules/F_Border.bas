Attribute VB_Name = "F_Border"
Option Explicit
Option Private Module

Private Function InvokeBorderToggle(ByVal targetKey As String, _
                                    ByVal lineStyle As XlLineStyle, _
                                    ByVal weight As XlBorderWeight, _
                                    ByVal macroName As String) As Boolean
    On Error GoTo CleanFail

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    Call RepeatRegister(macroName, lineStyle, weight)
    Call engine.ToggleBorder(targetKey, lineStyle, weight)
    InvokeBorderToggle = True

CleanExit:
    Exit Function
CleanFail:
    Call ErrorHandler(macroName)
    Resume CleanExit
End Function

Private Function InvokeBorderDelete(ByVal targetKey As String, _
                                    ByVal macroName As String) As Boolean
    On Error GoTo CleanFail

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    Call RepeatRegister(macroName)
    Call engine.DeleteBorder(targetKey)
    InvokeBorderDelete = True

CleanExit:
    Exit Function
CleanFail:
    Call ErrorHandler(macroName)
    Resume CleanExit
End Function

Private Function BorderColorInner(ByVal targetKey As String, _
                                  Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo CleanFail

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then GoTo CleanExit

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
    End If

    If resultColor Is Nothing Then GoTo CleanExit

    Call RepeatRegister("BorderColorInner", targetKey, resultColor)
    Call engine.SetBorderColor(targetKey, _
                               resultColor.IsNull, _
                               resultColor.IsThemeColor, _
                               resultColor.ThemeColor, _
                               resultColor.TintAndShade, _
                               resultColor.Color)
    BorderColorInner = True

CleanExit:
    Exit Function
CleanFail:
    Call ErrorHandler("BorderColorInner")
    Resume CleanExit
End Function

Function ToggleBorderAround(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                            Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    ToggleBorderAround = InvokeBorderToggle("Around", LineStyle, Weight, "ToggleBorderAround")
End Function

Function ToggleBorderLeft(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                          Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    ToggleBorderLeft = InvokeBorderToggle("Left", LineStyle, Weight, "ToggleBorderLeft")
End Function

Function ToggleBorderTop(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                         Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    ToggleBorderTop = InvokeBorderToggle("Top", LineStyle, Weight, "ToggleBorderTop")
End Function

Function ToggleBorderBottom(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                            Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    ToggleBorderBottom = InvokeBorderToggle("Bottom", LineStyle, Weight, "ToggleBorderBottom")
End Function

Function ToggleBorderRight(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                           Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    ToggleBorderRight = InvokeBorderToggle("Right", LineStyle, Weight, "ToggleBorderRight")
End Function

Function ToggleBorderInnerHorizontal(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                     Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    ToggleBorderInnerHorizontal = InvokeBorderToggle("InsideHorizontal", LineStyle, Weight, "ToggleBorderInnerHorizontal")
End Function

Function ToggleBorderInnerVertical(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                   Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    ToggleBorderInnerVertical = InvokeBorderToggle("InsideVertical", LineStyle, Weight, "ToggleBorderInnerVertical")
End Function

Function ToggleBorderInner(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                           Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    ToggleBorderInner = InvokeBorderToggle("InsideBoth", LineStyle, Weight, "ToggleBorderInner")
End Function

Function ToggleBorderDiagonalUp(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    ToggleBorderDiagonalUp = InvokeBorderToggle("DiagonalUp", LineStyle, Weight, "ToggleBorderDiagonalUp")
End Function

Function ToggleBorderDiagonalDown(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                  Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    ToggleBorderDiagonalDown = InvokeBorderToggle("DiagonalDown", LineStyle, Weight, "ToggleBorderDiagonalDown")
End Function

Function ToggleBorderAll(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                         Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    ToggleBorderAll = InvokeBorderToggle("All", LineStyle, Weight, "ToggleBorderAll")
End Function

Function DeleteBorderAround(Optional ByVal g As String) As Boolean
    DeleteBorderAround = InvokeBorderDelete("Around", "DeleteBorderAround")
End Function

Function DeleteBorderLeft(Optional ByVal g As String) As Boolean
    DeleteBorderLeft = InvokeBorderDelete("Left", "DeleteBorderLeft")
End Function

Function DeleteBorderTop(Optional ByVal g As String) As Boolean
    DeleteBorderTop = InvokeBorderDelete("Top", "DeleteBorderTop")
End Function

Function DeleteBorderBottom(Optional ByVal g As String) As Boolean
    DeleteBorderBottom = InvokeBorderDelete("Bottom", "DeleteBorderBottom")
End Function

Function DeleteBorderRight(Optional ByVal g As String) As Boolean
    DeleteBorderRight = InvokeBorderDelete("Right", "DeleteBorderRight")
End Function

Function DeleteBorderInnerHorizontal(Optional ByVal g As String) As Boolean
    DeleteBorderInnerHorizontal = InvokeBorderDelete("InsideHorizontal", "DeleteBorderInnerHorizontal")
End Function

Function DeleteBorderInnerVertical(Optional ByVal g As String) As Boolean
    DeleteBorderInnerVertical = InvokeBorderDelete("InsideVertical", "DeleteBorderInnerVertical")
End Function

Function DeleteBorderInner(Optional ByVal g As String) As Boolean
    DeleteBorderInner = InvokeBorderDelete("InsideBoth", "DeleteBorderInner")
End Function

Function DeleteBorderDiagonalUp(Optional ByVal g As String) As Boolean
    DeleteBorderDiagonalUp = InvokeBorderDelete("DiagonalUp", "DeleteBorderDiagonalUp")
End Function

Function DeleteBorderDiagonalDown(Optional ByVal g As String) As Boolean
    DeleteBorderDiagonalDown = InvokeBorderDelete("DiagonalDown", "DeleteBorderDiagonalDown")
End Function

Function DeleteBorderAll(Optional ByVal g As String) As Boolean
    DeleteBorderAll = InvokeBorderDelete("All", "DeleteBorderAll")
End Function

Function SetBorderColorAround(Optional ByVal g As String) As Boolean
    SetBorderColorAround = BorderColorInner("Around")
End Function

Function SetBorderColorLeft(Optional ByVal g As String) As Boolean
    SetBorderColorLeft = BorderColorInner("Left")
End Function

Function SetBorderColorTop(Optional ByVal g As String) As Boolean
    SetBorderColorTop = BorderColorInner("Top")
End Function

Function SetBorderColorBottom(Optional ByVal g As String) As Boolean
    SetBorderColorBottom = BorderColorInner("Bottom")
End Function

Function SetBorderColorRight(Optional ByVal g As String) As Boolean
    SetBorderColorRight = BorderColorInner("Right")
End Function

Function SetBorderColorInnerHorizontal(Optional ByVal g As String) As Boolean
    SetBorderColorInnerHorizontal = BorderColorInner("InsideHorizontal")
End Function

Function SetBorderColorInnerVertical(Optional ByVal g As String) As Boolean
    SetBorderColorInnerVertical = BorderColorInner("InsideVertical")
End Function

Function SetBorderColorInner(Optional ByVal g As String) As Boolean
    SetBorderColorInner = BorderColorInner("InsideBoth")
End Function

Function SetBorderColorAll(Optional ByVal g As String) As Boolean
    SetBorderColorAll = BorderColorInner("All")
End Function
