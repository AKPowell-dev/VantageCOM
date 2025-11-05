Attribute VB_Name = "F_Shapes"
'Attribute VB_Name = "F_Shapes"
Option Explicit
Option Private Module

Private Function GetFillFormatFromSelection() As Object
    Dim fillObj As Object

    If VarType(Selection) <> vbObject Then Exit Function

    On Error Resume Next
    Set fillObj = Selection.ShapeRange.Fill
    If Err.Number = 0 Then GoTo Clean
    Err.Clear

    Set fillObj = Selection.Fill
    If Err.Number = 0 Then GoTo Clean
    Err.Clear

    Set fillObj = Selection.format.Fill
    If Err.Number = 0 Then GoTo Clean
    Err.Clear

    Select Case TypeName(Selection)
        Case "Chart"
            Set fillObj = Selection.ChartArea.format.Fill
        Case "ChartArea", "PlotArea", "Series", "Trendline", "DataLabel", "DataPoint", "LegendEntry", "LegendKey"
            Set fillObj = Selection.format.Fill
        Case "ChartObject"
            Set fillObj = Selection.Chart.ChartArea.format.Fill
    End Select

    If fillObj Is Nothing And Not ActiveChart Is Nothing Then
        Set fillObj = ActiveChart.ChartArea.format.Fill
    End If

Clean:
    Set GetFillFormatFromSelection = fillObj
    On Error GoTo 0
End Function

Private Function GetInteriorFromSelection() As Object
    Dim interiorObj As Object

    If VarType(Selection) <> vbObject Then Exit Function

    On Error Resume Next
    Set interiorObj = Selection.Interior
    If Err.Number = 0 Then GoTo Clean
    Err.Clear

    If TypeName(Selection) = "ChartObject" Then
        Set interiorObj = Selection.Chart.ChartArea.Interior
    End If

Clean:
    Set GetInteriorFromSelection = interiorObj
    On Error GoTo 0
End Function

Private Function ApplyFillFormat(ByVal fillTarget As Object, ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo CleanExit

    With fillTarget
        If resultColor.IsNull Then
            On Error Resume Next
            .Visible = msoFalse
            .Transparency = 1
            On Error GoTo 0
        ElseIf resultColor.IsThemeColor Then
            On Error Resume Next
            .Visible = msoTrue
            If Not TypeName(fillTarget) = "ChartFill" Then .Solid
            On Error GoTo 0
            On Error Resume Next
            .ForeColor.ObjectThemeColor = resultColor.ObjectThemeColor
            .ForeColor.TintAndShade = resultColor.TintAndShade
            If Err.Number <> 0 Then
                .ForeColor.RGB = resultColor.Color
                Err.Clear
            End If
            .Transparency = 0
            On Error GoTo 0
        Else
            On Error Resume Next
            .Visible = msoTrue
            If Not TypeName(fillTarget) = "ChartFill" Then .Solid
            .ForeColor.RGB = resultColor.Color
            .Transparency = 0
            On Error GoTo 0
        End If
    End With

    ApplyFillFormat = True

CleanExit:
    On Error GoTo 0
End Function

Private Function ApplyInterior(ByVal interiorTarget As Object, ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo CleanExit

    With interiorTarget
        If resultColor.IsNull Then
            On Error Resume Next
            .Pattern = xlPatternNone
            .ColorIndex = xlColorIndexNone
            On Error GoTo 0
        ElseIf resultColor.IsThemeColor Then
            On Error Resume Next
            .Pattern = xlSolid
            .ThemeColor = resultColor.ThemeColor
            .TintAndShade = resultColor.TintAndShade
            If Err.Number <> 0 Then
                .Color = resultColor.Color
                Err.Clear
            End If
            On Error GoTo 0
        Else
            On Error Resume Next
            .Pattern = xlSolid
            .Color = resultColor.Color
            On Error GoTo 0
        End If
    End With

    ApplyInterior = True

CleanExit:
    On Error GoTo 0
End Function

Function ChangeShapeFillColor(Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    Dim fillTarget As Object
    Dim interiorTarget As Object
    Dim applied As Boolean

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
    End If

    If resultColor Is Nothing Then Exit Function

    Set fillTarget = GetFillFormatFromSelection()
    Set interiorTarget = GetInteriorFromSelection()

    If Not fillTarget Is Nothing Then
        applied = ApplyFillFormat(fillTarget, resultColor) Or applied
    End If

    If Not interiorTarget Is Nothing Then
        applied = ApplyInterior(interiorTarget, resultColor) Or applied
    End If

    If Not applied Then Exit Function

    Call RepeatRegister("ChangeShapeFillColor", resultColor)

    Set fillTarget = Nothing
    Set interiorTarget = Nothing
    Exit Function

Catch:
    Call ErrorHandler("ChangeShapeFillColor")
End Function

Function ChangeShapeFontColor(Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    Dim shp As ShapeRange

    If VarType(Selection) <> vbObject Then
        Exit Function
    End If

    Set shp = Selection.ShapeRange

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
    End If

    If Not resultColor Is Nothing Then
        With shp.TextFrame2.TextRange.Font.Fill.ForeColor
            If resultColor.IsNull Then
                .RGB = 0
            ElseIf resultColor.IsThemeColor Then
                .ObjectThemeColor = resultColor.ObjectThemeColor
                .TintAndShade = resultColor.TintAndShade
            Else
                .RGB = resultColor.Color
            End If

            Call RepeatRegister("ChangeShapeFontColor", resultColor)
        End With
        ChangeShapeFontColor = True
    End If

    Set shp = Nothing
    Exit Function

Catch:
    Call ErrorHandler("ChangeShapeFontColor")
End Function

Function ChangeShapeBorderColor(Optional garbage As String, _
                                Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    Dim shp As ShapeRange

    If VarType(Selection) <> vbObject Then
        ChangeShapeBorderColor = True
        Exit Function
    End If

    Set shp = Selection.ShapeRange

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
    End If

    If Not resultColor Is Nothing Then
        With shp.Line
            If resultColor.IsNull Then
                .Visible = msoFalse
            ElseIf resultColor.IsThemeColor Then
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = resultColor.ObjectThemeColor
                .ForeColor.TintAndShade = resultColor.TintAndShade
            Else
                .Visible = msoTrue
                .ForeColor.RGB = resultColor.Color
            End If

            Call RepeatRegister("ChangeShapeBorderColor", "", resultColor)
        End With
    End If

    Set shp = Nothing
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


