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
    On Error GoTo FallbackCopy

    If TypeName(Selection) = "Range" Then
        ClipboardSetCopyRange Selection
    Else
        ClipboardSetCopyRange Nothing
    End If

    Application.CommandBars.ExecuteMso "Copy"
    Exit Sub

FallbackCopy:
    ClipboardSetCopyRange Nothing
    Application.CommandBars.ExecuteMso "Copy"
End Sub

Public Sub ClipboardHandleCut()
    On Error Resume Next
    ClipboardSetCopyRange Nothing
    Application.CommandBars.ExecuteMso "Cut"
End Sub

Public Sub ClipboardHandlePaste()
    On Error GoTo FallbackPaste

    Call EnsureClipboardPayload

    Application.CommandBars.ExecuteMso "Paste"
    Exit Sub

FallbackPaste:
    Application.CommandBars.ExecuteMso "Paste"
End Sub

Private Sub ClipboardHandlePasteSpecial(ByVal controlId As String)
    On Error GoTo FailPaste

    Call EnsureClipboardPayload

    Application.CommandBars.ExecuteMso controlId
    Exit Sub

FailPaste:
    Application.CommandBars.ExecuteMso controlId
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

Private Sub EnsureClipboardPayload()
    If Application.CutCopyMode <> 0 Then Exit Sub
    If ClipboardHasContent() Then Exit Sub
    If gCopyRange Is Nothing Then Exit Sub
    If Not IsRangeValid(gCopyRange) Then
        Set gCopyRange = Nothing
        Exit Sub
    End If

    Dim savedSelection As Range
    Dim savedActive As Range

    If TypeName(Selection) = "Range" Then
        Set savedSelection = Selection
        Set savedActive = ActiveCell
    End If

    gCopyRange.Copy

    If Not savedSelection Is Nothing Then
        SafeSelectRange savedSelection
        If Not savedActive Is Nothing Then SafeActivateRange savedActive
    End If
End Sub

Public Function ClipboardGetCopyRange() As Range
    On Error Resume Next
    If Not gCopyRange Is Nothing Then
        Set ClipboardGetCopyRange = gCopyRange
    End If
End Function

Public Sub ClipboardSetCopyRange(ByVal rng As Range)
    On Error Resume Next
    If rng Is Nothing Then
        Set gCopyRange = Nothing
    Else
        Set gCopyRange = rng
    End If
End Sub

Public Sub ClipboardOpenPasteSpecial()
    EnsureClipboardPayload
    Application.CommandBars.ExecuteMso "PasteSpecialDialog"
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
