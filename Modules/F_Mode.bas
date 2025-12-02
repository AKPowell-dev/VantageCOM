Attribute VB_Name = "F_Mode"
Option Explicit
Option Private Module

Function ChangeToNormalMode(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If gVim Is Nothing Then
        ChangeToNormalMode = True
        Exit Function
    End If

    Call gVim.Mode.Change(MODE_NORMAL)
    Call SetStatusBar
    Exit Function

Catch:
    Call ErrorHandler("ChangeToNormalMode")
End Function

Function ChangeToShapeInsertMode(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    If gVim Is Nothing Then
        ChangeToShapeInsertMode = True
        Exit Function
    End If

    Call gVim.Mode.Change(MODE_SHAPEINSERT)
    Call SetStatusBar("-- SHAPE INSERT (ESC to exit) --")
    Exit Function

Catch:
    Call ErrorHandler("ChangeToShapeInsertMode")
End Function

Function StopVisualMode(Optional ByVal g As String) As Boolean
    ' Visual mode removed; keep as harmless no-op for callers
End Function
