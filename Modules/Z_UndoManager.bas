Attribute VB_Name = "Z_UndoManager"
Option Explicit

Private gUndoActive As Boolean
Private gUndoSkipNext As Boolean

Public Sub UndoPrepareForCommand(ByVal cmd As String)
    If gUndoSkipNext Then
        gUndoSkipNext = False
        Exit Sub
    End If

    If Not ShouldCaptureUndo(cmd) Then Exit Sub
    StartCustomUndo "macro: " & cmd
End Sub

Public Sub UndoFinalizeForCommand()
    FinalizeCustomUndo
End Sub

Public Sub UndoAbortForCommand()
    FinalizeCustomUndo
End Sub

Public Sub UndoClearSnapshot()
    gUndoActive = False
    gUndoSkipNext = False
End Sub

Public Sub UndoSuppressForNextCommand()
    gUndoSkipNext = True
End Sub

Private Function ShouldCaptureUndo(ByVal cmd As String) As Boolean
    Dim lower As String
    Dim token As Variant
    lower = LCase$(cmd)
    If lower = "" Then Exit Function

    If Left$(lower, 1) = "'" Then Exit Function
    If InStr(lower, "undo") > 0 Or InStr(lower, "redo") > 0 Then Exit Function

    Dim skipTokens As Variant
    skipTokens = Array("move", "scroll", "toggle", "show", "jump", "focus", "select", "start", "stop", "center", "undo_c", "keystroke")
    For Each token In skipTokens
        If Left$(lower, Len(token)) = token Then Exit Function
    Next token

    ShouldCaptureUndo = True
End Function

Private Sub StartCustomUndo(ByVal description As String)
    On Error GoTo CleanFail
    Dim undoRec As Object
    Set undoRec = GetUndoRecord()
    If undoRec Is Nothing Then Exit Sub

    undoRec.StartCustomRecord description
    gUndoActive = True
    Exit Sub

CleanFail:
    gUndoActive = False
End Sub

Private Sub FinalizeCustomUndo()
    On Error Resume Next
    If Not gUndoActive Then
        gUndoSkipNext = False
        Exit Sub
    End If

    Dim undoRec As Object
    Set undoRec = GetUndoRecord()
    If Not undoRec Is Nothing Then
        undoRec.EndCustomRecord
    End If

    gUndoActive = False
    gUndoSkipNext = False
    On Error GoTo 0
End Sub

Private Function GetUndoRecord() As Object
    On Error Resume Next
    Set GetUndoRecord = CallByName(Application, "UndoRecord", VbGet)
    On Error GoTo 0
End Function
