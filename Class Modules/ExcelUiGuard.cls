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

Private Sub Class_Initialize()
    ' Snapshot current UI flags and disable expensive UI updates
    m_prevScreenUpdating = Application.ScreenUpdating
    m_prevEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error Resume Next
    Set m_originalWindow = Application.ActiveWindow
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
    m_prevStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = False
    m_statusBarManaged = True
End Sub

Public Sub Restore()
    If m_restored Then Exit Sub
    On Error Resume Next

    ' ----------------------------------------------------------
    ' NO window/sheet activation and NO selection restoration.
    ' This prevents Excel from switching to a different window.
    ' Whatever selection is active at macro end remains as-is.
    ' ----------------------------------------------------------
    If Not m_originalWindow Is Nothing Then
        If Not Application.ActiveWindow Is m_originalWindow Then
            If m_originalWindow.Visible Then
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
        Application.DisplayStatusBar = m_prevStatusBar
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
