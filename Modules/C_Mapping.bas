Attribute VB_Name = "C_Mapping"
Option Explicit
Option Private Module

'/*
' * Dummy function used for unallocated keymap.
' */
Function Dummy()
    If gVim.DebugMode Then
        Call SetStatusBarTemporarily(gVim.Msg.NoKeyAllocation, 1000)
    End If
End Function

'/*
' * LazyLoad function to dynamically load procedures based on key mappings.
' *
' * @param {String} key - The key for which the procedure needs to be loaded.
' * @returns {Boolean} - True if the procedure is successfully loaded and executed, False otherwise.
' */
Function LazyLoad(ByVal key As String) As Boolean
    ' Start Vim if not already started
    If gVim Is Nothing Then
        Call StartVim
    End If

    ' Get procedure from the key map manager
    Dim cmd As String
    cmd = gVim.KeyMap.Get_(key)

    ' Clear mapping if command is empty string
    If cmd = "" Then
        Application.OnKey key

        Call KeyUpControlKeys
        Dim numLockWasOn As Boolean
        numLockWasOn = ((GetKeyState(NumLock_) And &H1) <> 0)
        Application.SendKeys key
        If numLockWasOn <> ((GetKeyState(NumLock_) And &H1) <> 0) Then
            keybd_event NumLock_, 0, 0, 0
            keybd_event NumLock_, 0, KEYUP, 0
        End If
        Call UnkeyUpControlKeys
        Exit Function
    End If

    On Error GoTo Catch

    Dim currentMode As String: currentMode = gVim.Mode.Current

    ' Run the procedure using Application.Run
    Dim ret As Variant
    Call UndoPrepareForCommand(cmd)
    ret = Application.Run(cmd)
    LazyLoad = CBool(ret)

    ' Prevent to register the dummy command
    If cmd = DUMMY_PROCEDURE Then
        GoTo Finally

    ' Prevent registration if the mode is changed by execution
    ElseIf currentMode <> gVim.Mode.Current Then
        GoTo Finally

    End If

    ' If result is not succeeded, show command form and exit
    If ret Then
        Call ShowCmdForm(key)
        GoTo Finally
    End If

    ' Register the key with the command
    If UndoShouldCapture(cmd) Then
        Application.OnKey key, BuildUndoWrappedProcedure(cmd)
    Else
        Application.OnKey key, cmd
    End If
    GoTo Finally

Catch:
    If Err.Number = 1004 Then
        Call SetStatusBarTemporarily(gVim.Msg.MissingMacro & cmd, 3000)
    Else
        Call ErrorHandler("LazyLoad")
    End If
    Call UndoAbortForCommand
    Resume Finally

Finally:
    Call UndoFinalizeForCommand
    Exit Function
End Function

Public Function RunMappedCommand(ByVal cmd As String) As Boolean
    On Error GoTo Catch

    Call UndoPrepareForCommand(cmd)
    RunMappedCommand = CBool(Application.Run(cmd))
    Call UndoFinalizeForCommand
    Exit Function

Catch:
    Call UndoAbortForCommand
    If Err.Number = 1004 Then
        Call SetStatusBarTemporarily(gVim.Msg.MissingMacro & cmd, 3000)
    Else
        Call ErrorHandler("RunMappedCommand")
    End If
End Function

Private Function BuildUndoWrappedProcedure(ByVal cmd As String) As String
    Dim safeCmd As String
    safeCmd = Replace(cmd, """", """""")
    BuildUndoWrappedProcedure = "'RunMappedCommand """ & safeCmd & """'"
End Function
