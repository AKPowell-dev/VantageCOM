VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_AppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_SheetChange(ByVal Sh As Object, ByVal target As Range)
    On Error GoTo CleanExit
    Dim c As Range
    If target Is Nothing Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For Each c In target.Cells
        ColorCell c
    Next c

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


