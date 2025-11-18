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
    Call InvokeBorderToggle("Around", LineStyle, Weight, "ToggleBorderAround")
    ToggleBorderAround = False
End Function

Function ToggleBorderLeft(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                          Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call InvokeBorderToggle("Left", LineStyle, Weight, "ToggleBorderLeft")
    ToggleBorderLeft = False
End Function

Function ToggleBorderTop(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                         Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call InvokeBorderToggle("Top", LineStyle, Weight, "ToggleBorderTop")
    ToggleBorderTop = False
End Function

Function ToggleBorderBottom(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                            Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call InvokeBorderToggle("Bottom", LineStyle, Weight, "ToggleBorderBottom")
    ToggleBorderBottom = False
End Function

Function ToggleBorderRight(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                           Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call InvokeBorderToggle("Right", LineStyle, Weight, "ToggleBorderRight")
    ToggleBorderRight = False
End Function

Function ToggleBorderInnerHorizontal(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                     Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call InvokeBorderToggle("InsideHorizontal", LineStyle, Weight, "ToggleBorderInnerHorizontal")
    ToggleBorderInnerHorizontal = False
End Function

Function ToggleBorderInnerVertical(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                   Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call InvokeBorderToggle("InsideVertical", LineStyle, Weight, "ToggleBorderInnerVertical")
    ToggleBorderInnerVertical = False
End Function

Function ToggleBorderInner(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                           Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call InvokeBorderToggle("InsideBoth", LineStyle, Weight, "ToggleBorderInner")
    ToggleBorderInner = False
End Function

Function ToggleBorderDiagonalUp(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call InvokeBorderToggle("DiagonalUp", LineStyle, Weight, "ToggleBorderDiagonalUp")
    ToggleBorderDiagonalUp = False
End Function

Function ToggleBorderDiagonalDown(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                                  Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call InvokeBorderToggle("DiagonalDown", LineStyle, Weight, "ToggleBorderDiagonalDown")
    ToggleBorderDiagonalDown = False
End Function

Function ToggleBorderAll(Optional ByVal LineStyle As XlLineStyle = xlContinuous, _
                         Optional ByVal Weight As XlBorderWeight = xlThin) As Boolean
    Call InvokeBorderToggle("All", LineStyle, Weight, "ToggleBorderAll")
    ToggleBorderAll = False
End Function

Function DeleteBorderAround(Optional ByVal g As String) As Boolean
    Call InvokeBorderDelete("Around", "DeleteBorderAround")
    DeleteBorderAround = False
End Function

Function DeleteBorderLeft(Optional ByVal g As String) As Boolean
    Call InvokeBorderDelete("Left", "DeleteBorderLeft")
    DeleteBorderLeft = False
End Function

Function DeleteBorderTop(Optional ByVal g As String) As Boolean
    Call InvokeBorderDelete("Top", "DeleteBorderTop")
    DeleteBorderTop = False
End Function

Function DeleteBorderBottom(Optional ByVal g As String) As Boolean
    Call InvokeBorderDelete("Bottom", "DeleteBorderBottom")
    DeleteBorderBottom = False
End Function

Function DeleteBorderRight(Optional ByVal g As String) As Boolean
    Call InvokeBorderDelete("Right", "DeleteBorderRight")
    DeleteBorderRight = False
End Function

Function DeleteBorderInnerHorizontal(Optional ByVal g As String) As Boolean
    Call InvokeBorderDelete("InsideHorizontal", "DeleteBorderInnerHorizontal")
    DeleteBorderInnerHorizontal = False
End Function

Function DeleteBorderInnerVertical(Optional ByVal g As String) As Boolean
    Call InvokeBorderDelete("InsideVertical", "DeleteBorderInnerVertical")
    DeleteBorderInnerVertical = False
End Function

Function DeleteBorderInner(Optional ByVal g As String) As Boolean
    Call InvokeBorderDelete("InsideBoth", "DeleteBorderInner")
    DeleteBorderInner = False
End Function

Function DeleteBorderDiagonalUp(Optional ByVal g As String) As Boolean
    Call InvokeBorderDelete("DiagonalUp", "DeleteBorderDiagonalUp")
    DeleteBorderDiagonalUp = False
End Function

Function DeleteBorderDiagonalDown(Optional ByVal g As String) As Boolean
    Call InvokeBorderDelete("DiagonalDown", "DeleteBorderDiagonalDown")
    DeleteBorderDiagonalDown = False
End Function

Function DeleteBorderAll(Optional ByVal g As String) As Boolean
    Call InvokeBorderDelete("All", "DeleteBorderAll")
    DeleteBorderAll = False
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
