Attribute VB_Name = "F_NameScrubber"
Option Explicit
Option Private Module

Public Function NameScrubber(Optional ByVal g As String) As Boolean
    Dim engine As Object
    On Error GoTo CleanFail

    HideCmdLineIfVisible

    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    engine.NameScrubber
    Exit Function

CleanFail:
    Call ErrorHandler("NameScrubber")
End Function

Private Sub HideCmdLineIfVisible()
    On Error Resume Next
    If UF_Cmd.Visible Then
        UF_Cmd.Hide
    End If
    Unload UF_Cmd
    If UF_CmdLine.Visible Then
        UF_CmdLine.Hide
    End If
    Unload UF_CmdLine
    On Error GoTo 0

    On Error Resume Next
    If Not gVim Is Nothing Then
        If gVim.Mode.Current = MODE_CMDLINE Then
            gVim.Mode.Change MODE_NORMAL
        End If
    End If
    On Error GoTo 0
End Sub
