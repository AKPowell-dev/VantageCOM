Attribute VB_Name = "C_Core"
Option Explicit
Option Private Module

Public gVim As cls_Vim              ' Core vim instance
Public gSelectionStamp As Long      ' Incremented on every selection change
Private gHelpKeySuppressed As Boolean

Private gStartupScheduled As Boolean
Private gStartupTime As Date
Public gClipboardHookReady As Boolean   ' Defer clipboard hooks until post-init

Public Sub SuppressExcelHelpKey()
    On Error Resume Next
    Application.OnKey "{F1}", ""
    If Err.Number = 0 Then
        gHelpKeySuppressed = True
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub RestoreExcelHelpKey()
    On Error Resume Next
    If gHelpKeySuppressed Then
        Application.OnKey "{F1}"
        gHelpKeySuppressed = False
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Sub ScheduleVimStartup()
    On Error Resume Next
    If gStartupScheduled Then Exit Sub
    ' Defer heavy init ~1s after open so Excel UI settles
    gStartupTime = Now + TimeSerial(0, 0, 1)
    Application.OnTime EarliestTime:=gStartupTime, Procedure:="'C_Core.StartVimDelayed'", Schedule:=True
    gStartupScheduled = True
End Sub

Sub StartVimDelayed()
    gStartupScheduled = False
    Call StartVim
End Sub

Sub CancelScheduledVimStartup()
    On Error Resume Next
    If gStartupScheduled Then
        Application.OnTime EarliestTime:=gStartupTime, Procedure:="'C_Core.StartVimDelayed'", Schedule:=False
        gStartupScheduled = False
    End If
End Sub

Function StartVim(Optional ByVal g As String) As Boolean
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean

    Call CancelScheduledVimStartup
    Call TimeClear

    ' Suppress UI/events during initialization
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    On Error GoTo Finally

    If gVim Is Nothing Then
        ' Create vim instance
        Set gVim = New cls_Vim

        ' Load default setting
        Call DefaultConfig

        ' Load custom config
        Call gVim.Config.LoadCustomConfig
    End If

    ' Defer clipboard hooking until post-init
    gClipboardHookReady = False

    ' Enable Vim addin (bind keys for current mode)
    gVim.Enabled = True

    ' Schedule post-init (hooks and any optional work) shortly after
    On Error Resume Next
    Application.OnTime EarliestTime:=Now + TimeSerial(0, 0, 1), Procedure:="'C_Core.StartVimPostInit'", Schedule:=True
    On Error GoTo 0

    Call SetStatusBarTemporarily(gVim.Msg.VimStarted & "(Load time: " & format(GetQueryPerformanceTime(), "0.000") & "s)", 3000)

    Application.OnKey ("=")

Finally:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
End Function

Sub StartVimPostInit()
    On Error Resume Next
    gClipboardHookReady = True
    ' Attach clipboard hooks now that startup is finished
    ClipboardRefresh
    On Error GoTo 0
End Sub

Function StopVim(Optional ByVal g As String) As Boolean
    If Not gVim Is Nothing Then
        gVim.Enabled = False
    End If
End Function

Function ExitVim(Optional ByVal g As String) As Boolean
    Call CancelScheduledVimStartup
    If Not gVim Is Nothing Then
        Call gVim.Quit
        Set gVim = Nothing
    End If
    Call RestoreExcelHelpKey
End Function

Function ReloadVim(Optional ByVal g As String) As Boolean
    If Not gVim Is Nothing Then
        Call ExitVim
    End If
    Call StartVim
End Function

Function ToggleVim(Optional ByVal g As String) As Boolean
    If gVim Is Nothing Then
        Call StartVim
    Else
        gVim.Enabled = Not gVim.Enabled
    End If
End Function

Function ToggleLang(Optional ByVal g As String) As Boolean
    If Not gVim Is Nothing Then
        gVim.IsJapanese = Not gVim.IsJapanese
    End If
End Function

Function ToggleDebugMode(Optional ByVal g As String) As Boolean
    If Not gVim Is Nothing Then
        gVim.DebugMode = Not gVim.DebugMode
    End If
End Function

Function EnterCmdlineMode(Optional ByVal g As String) As Boolean
    Dim cmdResult As String
    cmdResult = UF_CmdLine.Launch()

    ' Canceled
    If cmdResult = CMDLINE_CANCELED Or Len(cmdResult) = 0 Then
        Exit Function

    ' Only numbers
    ElseIf Not cmdResult Like "*[!0-9]*" And Len(cmdResult) > 0 Then
        Dim lineNum As Long
        If Len(cmdResult) > 7 Then
            lineNum = Cells.Rows.Count
        Else
            lineNum = CLng(cmdResult)
        End If

        If lineNum > Cells.Rows.Count Then
            lineNum = Cells.Rows.Count
        End If

        ' Go to line
        Call MoveToSpecifiedRow(CStr(lineNum))

        Exit Function
    End If

    Dim cmdAndArg() As String
    cmdAndArg = Split(cmdResult, " ", 2)

    Dim cmdSuggests() As String
    Dim prefix As String
    Dim isExcFlag As Boolean
    Dim i As Long

    If Right(cmdAndArg(0), 1) = "!" Then
        prefix = Left(cmdAndArg(0), Len(cmdAndArg(0)) - 1)
        isExcFlag = True
    Else
        prefix = cmdAndArg(0)
    End If
    cmdSuggests = gVim.KeyMap.Suggest(prefix, True)

    For i = LBound(cmdSuggests) To UBound(cmdSuggests)
        If EndsWith(cmdSuggests(i), "!") Then
            If Not isExcFlag Then
                cmdSuggests(i) = ""
            End If
        Else
            If isExcFlag Then
                cmdSuggests(i) = ""
            End If
        End If
    Next i
    cmdSuggests = Filter(cmdSuggests, prefix)

    If UBound(cmdSuggests) < 0 Then
        Call SetStatusBarTemporarily(gVim.Msg.NoCommandAvailable & cmdResult, 3000)
        Exit Function
    End If

    Dim cmd As String
    cmd = gVim.KeyMap.Get_(cmdAndArg(0), True)

    If cmd = "" Then
        If UBound(cmdSuggests) = 0 Then
            cmd = gVim.KeyMap.Get_(cmdSuggests(0), True)
        ElseIf UBound(cmdSuggests) > 0 Then
            Call SetStatusBarTemporarily(gVim.Msg.AmbiguousCommand & cmdResult, 3000)
            Exit Function
        End If
    End If

    If UBound(cmdAndArg) > 0 Then
        Application.Run cmd, Trim(cmdAndArg(1))
    Else
        Application.Run cmd
    End If
End Function

Function ShowCmdForm(ByVal prefixStr As String) As Boolean
ShowCmdForm = UF_Cmd.Launch(prefixStr)
End Function

Function ShowVersion(Optional ByVal g As String) As Boolean
    Dim versionStr As String
    versionStr = ThisWorkbook.BuiltinDocumentProperties("Comments")

    MsgBox versionStr, vbInformation
End Function

