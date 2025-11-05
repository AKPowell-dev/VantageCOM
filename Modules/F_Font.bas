Attribute VB_Name = "F_Font"
Option Explicit
Option Private Module

Function IncreaseFontSize(Optional ByVal g As String) As Boolean
    Call RepeatRegister("IncreaseFontSize")
    Call StopVisualMode

    Dim i As Long
    For i = 1 To gVim.Count1
        Call KeyStroke(Alt_, H_, F_, G_)
    Next i
End Function

Function DecreaseFontSize(Optional ByVal g As String) As Boolean
    Call RepeatRegister("DecreaseFontSize")
    Call StopVisualMode

    Dim i As Long
    For i = 1 To gVim.Count1
        Call KeyStroke(Alt_, H_, F_, K_)
    Next i
End Function

Function ChangeFontName(Optional ByVal g As String) As Boolean
    Call KeyStroke(Alt_, H_, F_, F_)
End Function

Function ChangeFontSize(Optional ByVal g As String) As Boolean
    Call KeyStroke(Alt_, H_, F_, S_)
End Function

Function AlignLeft(Optional ByVal g As String) As Boolean
    Call RepeatRegister("AlignLeft")
    Call StopVisualMode

    'Check excel version
    On Error GoTo Excel2019
    If CDbl(Application.Version) >= 16 Then
        'Raise error in Excel 2016, 2019 (Sequence exists in Excel 2021 and later)
        WorksheetFunction.Sequence 1
    End If

    'Default
    Call KeyStroke(Alt_, H_, A_, L_)
    Exit Function

Excel2019:
    'Excel 2019 and earlier
    Call KeyStroke(Alt_, H_, L_, k1_)
End Function

Function AlignCenter(Optional ByVal g As String) As Boolean
    Call RepeatRegister("AlignCenter")
    Call StopVisualMode
    Call KeyStroke(Alt_, H_, A_, C_)
End Function

Function AlignRight(Optional ByVal g As String) As Boolean
    Call RepeatRegister("AlignRight")
    Call StopVisualMode

    'Check excel version
    On Error GoTo Excel2019
    If CDbl(Application.Version) >= 16 Then
        'Raise error in Excel 2016, 2019 (Sequence exists in Excel 2021 and later)
        WorksheetFunction.Sequence 1
    End If

    'Default
    Call KeyStroke(Alt_, H_, A_, R_)
    Exit Function

Excel2019:
    'Excel 2019 and earlier
    Call KeyStroke(Alt_, H_, R_)
End Function

Function AlignTop(Optional ByVal g As String) As Boolean
    Call RepeatRegister("AlignTop")
    Call StopVisualMode
    Call KeyStroke(Alt_, H_, A_, T_)
End Function

Function AlignMiddle(Optional ByVal g As String) As Boolean
    Call RepeatRegister("AlignMiddle")
    Call StopVisualMode
    Call KeyStroke(Alt_, H_, A_, M_)
End Function

Function AlignBottom(Optional ByVal g As String) As Boolean
    Call RepeatRegister("AlignBottom")
    Call StopVisualMode
    Call KeyStroke(Alt_, H_, A_, B_)
End Function

Function ToggleBold(Optional ByVal g As String) As Boolean
    Call RepeatRegister("ToggleBold")
    Call StopVisualMode
    Call KeyStroke(Ctrl_ + k2_)
End Function

Function ToggleItalic(Optional ByVal g As String) As Boolean
    Call RepeatRegister("ToggleItalic")
    Call StopVisualMode
    Call KeyStroke(Ctrl_ + k3_)
End Function

Function ToggleUnderline(Optional ByVal g As String) As Boolean
    Call RepeatRegister("ToggleUnderline")
    Call StopVisualMode
    Call KeyStroke(Ctrl_ + k4_)
End Function

Function ToggleStrikethrough(Optional ByVal g As String) As Boolean
    Call RepeatRegister("ToggleStrikethrough")
    Call StopVisualMode
    Call KeyStroke(Ctrl_ + k5_)
End Function

Function ChangeFormat(Optional ByVal g As String) As Boolean
    Call KeyStroke(Alt_, H_, N_, Down_, Down_)
End Function

Function showFontDialog(Optional ByVal g As String) As Boolean
    Call StopVisualMode
    Call KeyStroke(Ctrl_ + k1_)
End Function

Function ChangeFontColor(Optional ByVal resultColor As cls_FontColor) As Boolean
    On Error GoTo Catch

    If TypeName(Selection) = "Nothing" Then
        Exit Function
    End If

    If resultColor Is Nothing Then
        Set resultColor = UF_ColorPicker.Launch()
    End If

    If Not resultColor Is Nothing Then
        If ApplyFontColorToSelection(resultColor) Then
            Call RepeatRegister("ChangeFontColor", resultColor)
            Call StopVisualMode
            ChangeFontColor = True
        End If
    End If
    Exit Function

Catch:
    Call ErrorHandler("ChangeFontColor")
End Function

Private Function ApplyFontColorToSelection(ByVal colorInfo As cls_FontColor) As Boolean
    Dim selObj As Object
    Dim applied As Boolean

    On Error Resume Next
    Set selObj = Selection
    On Error GoTo 0
    If selObj Is Nothing Then Exit Function

    If TrySetFontColor(selObj, colorInfo) Then
        ApplyFontColorToSelection = True
        Exit Function
    End If

    Select Case TypeName(selObj)
        Case "DataLabels"
            Dim lbl As Object
            For Each lbl In selObj
                applied = TrySetFontColor(lbl, colorInfo) Or applied
            Next lbl
            ApplyFontColorToSelection = applied
        Case "DataLabel"
            ApplyFontColorToSelection = TrySetFontColor(selObj, colorInfo)
        Case "DataPoint"
            If selObj.HasDataLabel Then
                ApplyFontColorToSelection = TrySetFontColor(selObj.DataLabel, colorInfo)
            End If
        Case "Series"
            If selObj.HasDataLabels Then
                Dim dl As DataLabel
                For Each dl In selObj.DataLabels
                    applied = TrySetFontColor(dl, colorInfo) Or applied
                Next dl
                ApplyFontColorToSelection = applied
            End If
        Case "ShapeRange"
            Dim shp As Shape
            For Each shp In selObj
                If shp.TextFrame2.HasText Then
                    applied = TrySetFontColor(shp.TextFrame2.TextRange, colorInfo) Or applied
                End If
                If shp.HasChart Then
                    applied = ApplyFontColorToSelection(shp.Chart) Or applied
                End If
            Next shp
            ApplyFontColorToSelection = applied
        Case "Chart"
            On Error Resume Next
            If selObj.HasTitle Then applied = TrySetFontColor(selObj.ChartTitle, colorInfo) Or applied
            If selObj.HasDataTable Then applied = TrySetFontColor(selObj.DataTable, colorInfo) Or applied
            ApplyFontColorToSelection = applied
    End Select
End Function

Private Function TrySetFontColor(ByVal target As Object, ByVal colorInfo As cls_FontColor) As Boolean
    Dim fontObj As Object
    On Error GoTo Fail
    Set fontObj = CallByName(target, "Font", VbGet)
    Call ApplyFontColorToFontObject(fontObj, colorInfo)
    TrySetFontColor = True
    Exit Function
Fail:
    Err.Clear
End Function

Private Sub ApplyFontColorToFontObject(ByVal fontObj As Object, ByVal colorInfo As cls_FontColor)
    If fontObj Is Nothing Then Exit Sub
    If colorInfo Is Nothing Then Exit Sub
    On Error GoTo CleanFail
    If colorInfo.IsNull Then
        fontObj.ColorIndex = xlAutomatic
    ElseIf colorInfo.IsThemeColor Then
        On Error Resume Next
        fontObj.ThemeColor = colorInfo.ThemeColor
        If Err.Number <> 0 Then
            Err.Clear
            fontObj.Color = colorInfo.Color
        Else
            fontObj.TintAndShade = colorInfo.TintAndShade
        End If
        On Error GoTo CleanFail
    Else
        fontObj.Color = colorInfo.Color
    End If
    Exit Sub

CleanFail:
    On Error Resume Next
    fontObj.Color = colorInfo.Color
    On Error GoTo 0
End Sub
