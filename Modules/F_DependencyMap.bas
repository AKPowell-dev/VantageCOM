Attribute VB_Name = "F_DependencyMap"
Option Explicit

Sub DrawDependencyMap()
    Dim engine As Object
    On Error GoTo CleanFail

    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub
    engine.DrawDependencyMap
    Exit Sub

CleanFail:
    Call ErrorHandler("DrawDependencyMap")
End Sub
