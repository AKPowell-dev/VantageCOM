Attribute VB_Name = "C_FormResize"
Option Explicit
Option Private Module

' Makes a (modal) UserForm resizable and reflows its controls on resize.
'
' Adds the WS_THICKFRAME / WS_MAXIMIZEBOX window styles to the form's window and
' subclasses its window procedure so WM_SIZE drives the form's Relayout method.
' WM_GETMINMAXINFO enforces a sensible minimum size. Follows the same
' SetWindowLongPtr / CallWindowProc / AddressOf pattern as C_HotkeyHook.
'
' Only one form is hooked at a time (econs is modal), which keeps the state
' simple. Always paired: EnableFormResize on activate, DisableFormResize on
' close/terminate, so the original window procedure is restored before the
' window is destroyed.

#If Win64 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
        ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
        ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
        ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal Msg As Long, _
        ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, _
        ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        Destination As Any, Source As Any, ByVal Length As Long)
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
        ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" ( _
        ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
        ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        Destination As Any, Source As Any, ByVal Length As Long)
#End If

Private Const GWL_STYLE As Long = -16
Private Const GWL_WNDPROC As Long = -4
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WM_SIZE As Long = &H5
Private Const WM_GETMINMAXINFO As Long = &H24
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20

' Minimum window track size in pixels.
Private Const MIN_TRACK_X As Long = 500
Private Const MIN_TRACK_Y As Long = 560

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private gPrevWndProc As LongPtr
Private gHwnd As LongPtr
Private gInstalled As Boolean
Private gInLayout As Boolean
Private gForm As Object

Public Sub EnableFormResize(ByVal frm As Object, ByVal caption As String)
    On Error GoTo CleanFail

    Dim h As LongPtr
    h = FindWindow("ThunderDFrame", caption)        ' modal MSForms window class
    If h = 0 Then h = FindWindow("ThunderXFrame", caption)
    If h = 0 Then h = FindWindow(vbNullString, caption)
    If h = 0 Then Exit Sub

    If gInstalled Then
        If gHwnd = h Then Exit Sub                  ' already hooked this window
        Call DisableFormResize
    End If

    Dim st As LongPtr
    st = GetWindowLongPtr(h, GWL_STYLE)
    SetWindowLongPtr h, GWL_STYLE, st Or WS_THICKFRAME Or WS_MAXIMIZEBOX
    SetWindowPos h, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED

    Set gForm = frm
    gPrevWndProc = SetWindowLongPtr(h, GWL_WNDPROC, AddressOf FormResizeWndProc)
    If gPrevWndProc = 0 Then
        Set gForm = Nothing
        Exit Sub
    End If

    gHwnd = h
    gInstalled = True
    Exit Sub

CleanFail:
    gInstalled = False
End Sub

Public Sub DisableFormResize()
    On Error Resume Next
    If Not gInstalled Then Exit Sub
    If gHwnd <> 0 And gPrevWndProc <> 0 Then
        SetWindowLongPtr gHwnd, GWL_WNDPROC, gPrevWndProc
    End If
    gInstalled = False
    gHwnd = 0
    gPrevWndProc = 0
    Set gForm = Nothing
End Sub

Public Function FormResizeWndProc( _
    ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr _
) As LongPtr
    On Error GoTo Passthrough

    If Msg = WM_GETMINMAXINFO Then
        Dim mmi As MINMAXINFO
        CopyMemory mmi, ByVal lParam, LenB(mmi)
        If mmi.ptMinTrackSize.X < MIN_TRACK_X Then mmi.ptMinTrackSize.X = MIN_TRACK_X
        If mmi.ptMinTrackSize.Y < MIN_TRACK_Y Then mmi.ptMinTrackSize.Y = MIN_TRACK_Y
        CopyMemory ByVal lParam, mmi, LenB(mmi)
        FormResizeWndProc = 0
        Exit Function
    ElseIf Msg = WM_SIZE Then
        If gInstalled And Not gInLayout And Not gForm Is Nothing Then
            gInLayout = True
            gForm.Relayout
            gInLayout = False
        End If
    End If

Passthrough:
    On Error Resume Next
    FormResizeWndProc = CallWindowProc(gPrevWndProc, hwnd, Msg, wParam, lParam)
End Function
