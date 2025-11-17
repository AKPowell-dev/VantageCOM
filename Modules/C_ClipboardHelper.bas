Attribute VB_Name = "C_ClipboardHelper"
Option Explicit

Private gCopyRange As Range
Private gClipboardHooksActive As Boolean

Public Sub ClipboardHook()
    On Error Resume Next
    If gClipboardHooksActive Then Exit Sub

    Application.OnKey "^c", "ClipboardHandleCopy"
    Application.OnKey "^{INSERT}", "ClipboardHandleCopy"
    Application.OnKey "^x", "ClipboardHandleCut"
    Application.OnKey "^v", "ClipboardHandlePaste"
    Application.OnKey "+{INSERT}", "ClipboardHandlePaste"
    Application.OnKey "^%v", "ClipboardOpenPasteSpecial"
    Application.OnKey "^%{v}", "ClipboardOpenPasteSpecial"
    Application.OnKey "^%+v", "ClipboardHandlePasteValues"
    Application.OnKey "^%+{v}", "ClipboardHandlePasteValues"
    Application.OnKey "^%+f", "ClipboardHandlePasteFormulas"
    Application.OnKey "^%+{f}", "ClipboardHandlePasteFormulas"

    gClipboardHooksActive = True
End Sub

Public Sub ClipboardUnhook()
    On Error Resume Next
    If Not gClipboardHooksActive Then Exit Sub

    Application.OnKey "^c"
    Application.OnKey "^{INSERT}"
    Application.OnKey "^x"
    Application.OnKey "^v"
    Application.OnKey "+{INSERT}"
    Application.OnKey "^%v"
    Application.OnKey "^%{v}"
    Application.OnKey "^%+v"
    Application.OnKey "^%+{v}"
    Application.OnKey "^%+f"
    Application.OnKey "^%+{f}"

    ClipboardSetCopyRange Nothing
    gClipboardHooksActive = False
End Sub

Public Sub ClipboardHandleCopy()
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub
    engine.ClipboardHandleCopy
End Sub

Public Sub ClipboardHandleCut()
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub
    engine.ClipboardHandleCut
End Sub

Public Sub ClipboardHandlePaste()
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub
    engine.ClipboardHandlePaste
End Sub

Private Sub ClipboardHandlePasteSpecial(ByVal controlId As String)
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub
    Select Case controlId
        Case "PasteValues"
            engine.ClipboardHandlePasteValues
        Case "PasteFormulas"
            engine.ClipboardHandlePasteFormulas
        Case Else
            engine.ClipboardOpenPasteSpecial
    End Select
End Sub

Private Function ClipboardHasContent() As Boolean
    On Error Resume Next
    Dim formats As Variant
    formats = Application.ClipboardFormats
    On Error GoTo 0

    If IsEmpty(formats) Then
        ClipboardHasContent = False
    ElseIf IsArray(formats) Then
        If UBound(formats) >= LBound(formats) Then
            ClipboardHasContent = True
        End If
    Else
        ClipboardHasContent = (VarType(formats) <> vbEmpty)
    End If
End Function

Public Function ClipboardGetCopyRange() As Range
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    Set ClipboardGetCopyRange = engine.ClipboardGetCopyRange
End Function

Public Sub ClipboardSetCopyRange(ByVal rng As Range)
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub
    engine.ClipboardSetCopyRange rng
End Sub

Public Sub ClipboardOpenPasteSpecial()
    ClipboardHandlePasteSpecial "PasteSpecialDialog"
End Sub

Public Sub ClipboardHandlePasteValues()
    ClipboardHandlePasteSpecial "PasteValues"
End Sub

Public Sub ClipboardHandlePasteFormulas()
    ClipboardHandlePasteSpecial "PasteFormulas"
End Sub

Public Sub ClipboardRefresh()
    On Error Resume Next

    ' Defer clipboard hooks until core startup completes
    If gVim Is Nothing Then
        ClipboardUnhook
    ElseIf gVim.KeyMap Is Nothing Then
        ClipboardUnhook
    ElseIf Not gVim.Enabled Then
        ClipboardUnhook
    ElseIf Not gClipboardHookReady Then
        ClipboardUnhook
    Else
        ClipboardHook
    End If
End Sub
