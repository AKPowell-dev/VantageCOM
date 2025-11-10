VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelUiGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_prevScreenUpdating As Boolean
Private m_prevEnableEvents As Boolean
Private m_prevStatusBar As Boolean
Private m_statusBarManaged As Boolean
Private m_restored As Boolean

' We still capture these, in case you later want conditional restore,
' but we won't use them in Restore() below.
Private m_originalWindow As Window
Private m_originalWorkbook As Workbook
Private m_originalSheet As Worksheet
Private m_originalSelection As Object
Private m_originalCell As Range

Private Function ResolveSelectionWindow() As Window
    Dim sel As Object
    Dim candidateSheet As Worksheet
    Dim candidateWorkbook As Workbook
    Dim win As Window
    Dim activeWin As Window

    On Error Resume Next
    Set activeWin = Application.ActiveWindow
    On Error GoTo 0

    On Error Resume Next
    Set sel = Selection
    On Error GoTo 0

    If Not sel Is Nothing Then
        Select Case TypeName(sel)
            Case "Range"
                On Error Resume Next
                Set candidateSheet = sel.Worksheet
                On Error GoTo 0
            Case "ChartObject"
                On Error Resume Next
                Set candidateSheet = sel.Parent.Parent
                On Error GoTo 0
            Case "Chart"
                Dim parentObj As Object
                On Error Resume Next
                Set parentObj = sel.Parent
                On Error GoTo 0
                If Not parentObj Is Nothing Then
                    Select Case TypeName(parentObj)
                        Case "ChartObject"
                            Set candidateSheet = parentObj.Parent
                        Case "Worksheet"
                            Set candidateSheet = parentObj
                    End Select
                End If
            Case "Worksheet"
                Set candidateSheet = sel
        End Select
    End If

    If candidateSheet Is Nothing Then
        If Not activeWin Is Nothing Then
            Set ResolveSelectionWindow = activeWin
        Else
            On Error Resume Next
            Set ResolveSelectionWindow = Application.ActiveWindow
            On Error GoTo 0
        End If
        Exit Function
    End If

    Set candidateWorkbook = candidateSheet.Parent
    If candidateWorkbook Is Nothing Then
        If Not activeWin Is Nothing Then
            Set ResolveSelectionWindow = activeWin
        Else
            On Error Resume Next
            Set ResolveSelectionWindow = Application.ActiveWindow
            On Error GoTo 0
        End If
        Exit Function
    End If

    If Not activeWin Is Nothing Then
        On Error Resume Next
        If activeWin.Visible Then
            If Not activeWin.ActiveSheet Is Nothing Then
                If activeWin.ActiveSheet Is candidateSheet Then
                    Set ResolveSelectionWindow = activeWin
                    On Error GoTo 0
                    Exit Function
                End If
            End If
        End If
        On Error GoTo 0
    End If

    On Error Resume Next
    For Each win In candidateWorkbook.Windows
        If Not win Is Nothing Then
            If win.Visible Then
                If Not win.ActiveSheet Is Nothing Then
                    If win.ActiveSheet Is candidateSheet Then
                        If TypeName(sel) = "Range" Then
                            Dim cmp As Range
                            Set cmp = Nothing
                            On Error Resume Next
                            Set cmp = win.RangeSelection
                            On Error GoTo 0
                            If Not cmp Is Nothing Then
                                If cmp.Address(True, True, xlA1, True) = sel.Address(True, True, xlA1, True) Then
                                    Set ResolveSelectionWindow = win
                                    Exit Function
                                End If
                            End If
                        Else
                            Set ResolveSelectionWindow = win
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next win
    On Error GoTo 0

    On Error Resume Next
    If candidateWorkbook.Windows.Count = 1 Then
        Set ResolveSelectionWindow = candidateWorkbook.Windows(1)
    End If
    If ResolveSelectionWindow Is Nothing Then
        Set ResolveSelectionWindow = Application.ActiveWindow
    End If
    On Error GoTo 0
End Function

Public Sub EnsureWindow()
    On Error Resume Next
    If Not m_originalWindow Is Nothing Then
        If m_originalWindow.Visible Then
            If Not Application.ActiveWindow Is m_originalWindow Then
                m_originalWindow.Activate
            End If
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub Class_Initialize()
    ' Snapshot current UI flags and disable expensive UI updates
    m_prevScreenUpdating = Application.ScreenUpdating
    m_prevEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error Resume Next
    Set m_originalWindow = ResolveSelectionWindow()
    Set m_originalWorkbook = Application.ActiveWorkbook
    If Not m_originalWindow Is Nothing Then
        Set m_originalSheet = m_originalWindow.ActiveSheet
    Else
        Set m_originalSheet = ActiveSheet
    End If
    If Not Selection Is Nothing Then
        Set m_originalSelection = Selection
        If TypeName(m_originalSelection) = "Range" Then
            Set m_originalCell = ActiveCell
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub DisableStatusBar()
    If m_statusBarManaged Then Exit Sub
    On Error Resume Next
    m_prevStatusBar = Application.DisplayStatusBar
    On Error GoTo 0
    m_statusBarManaged = True
End Sub

Public Sub Restore()
    If m_restored Then Exit Sub
    On Error Resume Next

    If Not m_originalWindow Is Nothing Then
        If m_originalWindow.Visible Then
            If Not Application.ActiveWindow Is m_originalWindow Then
                m_originalWindow.Activate
            End If
        End If
    End If

    ' Release object references
    Set m_originalSelection = Nothing
    Set m_originalCell = Nothing
    Set m_originalSheet = Nothing
    Set m_originalWindow = Nothing
    Set m_originalWorkbook = Nothing

    ' Restore application flags
    Application.ScreenUpdating = m_prevScreenUpdating
    Application.EnableEvents = m_prevEnableEvents

    If m_statusBarManaged Then
        On Error Resume Next
        Application.DisplayStatusBar = m_prevStatusBar
        On Error GoTo 0
    End If

    m_restored = True
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Restore
End Sub

Private Sub TrySelectObject(ByVal candidate As Object)
    On Error Resume Next
    If candidate Is Nothing Then Exit Sub
    CallByName candidate, "Select", VbMethod
    On Error GoTo 0
End Sub
