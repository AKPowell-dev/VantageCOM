Attribute VB_Name = "C_HotkeyHook"
Option Explicit
Option Private Module

#If Win64 Then
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
        ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
        ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal Msg As Long, _
        ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
        ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

Private Const GWL_WNDPROC As Long = -4
Private Const WM_KEYDOWN As Long = &H100
Private Const VK_CONTROL As Long = &H11

Private gPrevWndProc As LongPtr
Private gHookHwnd As LongPtr
Private gHookInstalled As Boolean
Private gFormulaNavigatorQueued As Boolean

Public Sub InstallFormulaNavigatorHook()
    On Error GoTo CleanFail
    If gHookInstalled Then
        If gHookHwnd = Application.Hwnd And gHookHwnd <> 0 Then Exit Sub
        Call UninstallFormulaNavigatorHook
    End If

    gHookHwnd = Application.Hwnd
    If gHookHwnd = 0 Then Exit Sub

    gPrevWndProc = SetWindowLongPtr(gHookHwnd, GWL_WNDPROC, AddressOf FormulaNavigatorWndProc)
    If gPrevWndProc = 0 Then Exit Sub

    gHookInstalled = True
    Exit Sub

CleanFail:
    gHookInstalled = False
End Sub

Public Sub UninstallFormulaNavigatorHook()
    On Error Resume Next
    If Not gHookInstalled Then Exit Sub
    If gHookHwnd <> 0 And gPrevWndProc <> 0 Then
        SetWindowLongPtr gHookHwnd, GWL_WNDPROC, gPrevWndProc
    End If
    gHookInstalled = False
    gHookHwnd = 0
    gPrevWndProc = 0
End Sub

Public Function FormulaNavigatorWndProc( _
    ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr _
) As LongPtr
    On Error GoTo Passthrough

    If gVim Is Nothing Then GoTo Passthrough
    If Not gVim.Enabled Then GoTo Passthrough

    If Msg = WM_KEYDOWN Then
        Dim keyCode As Long
        keyCode = CLng(wParam)
        If keyCode = OpeningSquareBracket_ Or keyCode = Escape_ Then
            If (GetKeyState(VK_CONTROL) And &H8000) <> 0 Then
                QueueFormulaNavigator
                FormulaNavigatorWndProc = 0
                Exit Function
            End If
        End If
    End If

Passthrough:
    FormulaNavigatorWndProc = CallWindowProc(gPrevWndProc, hwnd, Msg, wParam, lParam)
End Function

Public Sub RunFormulaNavigatorHotkey()
    On Error GoTo CleanFail
    gFormulaNavigatorQueued = False
    Call SetStatusBarTemporarily("[DEBUG] Ctrl+[ -> FormulaNavigatorNext", 1000, True)
    Call FormulaNavigatorNext
    Exit Sub

CleanFail:
    gFormulaNavigatorQueued = False
End Sub

Private Sub QueueFormulaNavigator()
    On Error Resume Next
    If gFormulaNavigatorQueued Then Exit Sub
    gFormulaNavigatorQueued = True
    Call SetStatusBarTemporarily("[DEBUG] Ctrl+[ captured (VBA hook)", 1000, True)
    Application.OnTime Now, "'C_HotkeyHook.RunFormulaNavigatorHotkey'"
End Sub
