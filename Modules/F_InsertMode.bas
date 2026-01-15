Attribute VB_Name = "F_InsertMode"
Option Explicit
Option Private Module

Function InsertWithIME(Optional ByVal g As String) As Boolean
    Dim engine As Object
    Dim startedEditing As Boolean

    On Error GoTo Catch
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function

    startedEditing = engine.InsertWithIME()
    If startedEditing Then
        Call StartEditing
    End If
    Exit Function
Catch:
    Call ErrorHandler("InsertWithIME")
End Function

Function InsertWithoutIME(Optional ByVal g As String) As Boolean
    Dim engine As Object
    Dim startedEditing As Boolean

    On Error GoTo Catch
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function

    startedEditing = engine.InsertWithoutIME()
    If startedEditing Then
        Call StartEditing
    End If
    Exit Function
Catch:
    Call ErrorHandler("InsertWithoutIME")
End Function

Function AppendWithIME(Optional ByVal g As String) As Boolean
    Dim engine As Object
    Dim startedEditing As Boolean

    On Error GoTo Catch
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function

    startedEditing = engine.AppendWithIME()
    If startedEditing Then
        Call StartEditing
    End If
    Exit Function
Catch:
    Call ErrorHandler("AppendWithIME")
End Function

Function AppendWithoutIME(Optional ByVal g As String) As Boolean
    Dim engine As Object
    Dim startedEditing As Boolean

    On Error GoTo Catch
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function

    startedEditing = engine.AppendWithoutIME()
    If startedEditing Then
        Call StartEditing
    End If
    Exit Function
Catch:
    Call ErrorHandler("AppendWithoutIME")
End Function

Function SubstituteWithIME(Optional ByVal g As String) As Boolean
    Dim engine As Object
    Dim startedEditing As Boolean

    On Error GoTo Catch
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function

    startedEditing = engine.SubstituteWithIME()
    If startedEditing Then
        Call StartEditing
    End If
    Exit Function
Catch:
    Call ErrorHandler("SubstituteWithIME")
End Function

Function SubstituteWithoutIME(Optional ByVal g As String) As Boolean
    Dim engine As Object
    Dim startedEditing As Boolean

    On Error GoTo Catch
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function

    startedEditing = engine.SubstituteWithoutIME()
    If startedEditing Then
        Call StartEditing
    End If
    Exit Function
Catch:
    Call ErrorHandler("SubstituteWithoutIME")
End Function

Function InsertFollowLangMode(Optional ByVal g As String) As Boolean
    If gVim.IsJapanese Then
        Call InsertWithIME
    Else
        Call InsertWithoutIME
    End If
End Function

Function InsertNotFollowLangMode(Optional ByVal g As String) As Boolean
    If Not gVim.IsJapanese Then
        Call InsertWithIME
    Else
        Call InsertWithoutIME
    End If
End Function

Function AppendFollowLangMode(Optional ByVal g As String) As Boolean
    If gVim.IsJapanese Then
        Call AppendWithIME
    Else
        Call AppendWithoutIME
    End If
End Function

Function AppendNotFollowLangMode(Optional ByVal g As String) As Boolean
    If Not gVim.IsJapanese Then
        Call AppendWithIME
    Else
        Call AppendWithoutIME
    End If
End Function

Function SubstituteFollowLangMode(Optional ByVal g As String) As Boolean
    If gVim.IsJapanese Then
        Call SubstituteWithIME
    Else
        Call SubstituteWithoutIME
    End If
End Function

Function SubstituteNotFollowLangMode(Optional ByVal g As String) As Boolean
    If Not gVim.IsJapanese Then
        Call SubstituteWithIME
    Else
        Call SubstituteWithoutIME
    End If
End Function

Private Sub StartEditing()
    gVim.Vars.FromInsertCmd = True
    Application.OnTime Now + (1 / 86400) * 0.02, "StopEditing"
End Sub

Private Sub StopEditing()
    Call StopVisualMode
    gVim.Vars.FromInsertCmd = False
End Sub



