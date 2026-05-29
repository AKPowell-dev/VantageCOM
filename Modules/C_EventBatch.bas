Attribute VB_Name = "C_EventBatch"
Option Explicit
Option Private Module

Private mAutoColorRange As Range
Private mAutoColorScheduled As Boolean
Private mAutoColorBusy As Boolean
Private mBindAllScheduled As Boolean

'/*
' * Queues a cell range for deferred AutoColor processing.
' * Multiple rapid SheetChange events (e.g. CapIQ bulk refresh) collapse into
' * a single FlushAutoColor call on the next idle tick.
' */
Public Sub QueueAutoColor(ByVal rng As Range)
    On Error Resume Next
    If mAutoColorRange Is Nothing Then
        Set mAutoColorRange = rng
    ElseIf mAutoColorRange.Worksheet.Name = rng.Worksheet.Name And _
           mAutoColorRange.Worksheet.Parent.Name = rng.Worksheet.Parent.Name Then
        Set mAutoColorRange = Application.Union(mAutoColorRange, rng)
    Else
        Set mAutoColorRange = rng
    End If
    On Error GoTo 0

    If Not mAutoColorScheduled Then
        mAutoColorScheduled = True
        Application.OnTime Now, "'C_EventBatch.FlushAutoColor'"
    End If
End Sub

'/*
' * Processes accumulated AutoColor changes in one batched COM call.
' * Called by Application.OnTime after QueueAutoColor schedules it.
' */
Public Sub FlushAutoColor()
    On Error GoTo Done
    mAutoColorScheduled = False
    If mAutoColorRange Is Nothing Then Exit Sub
    If mAutoColorBusy Then Exit Sub

    Dim rng As Range
    Set rng = mAutoColorRange
    Set mAutoColorRange = Nothing

    Dim engine As Object
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub

    mAutoColorBusy = True
    engine.AutoColorRange rng, 50
Done:
    mAutoColorBusy = False
End Sub

'/*
' * Schedules a single deferred BindAll. Rapid successive window activations
' * (e.g. switching workbooks quickly) collapse into one rebind.
' */
Public Sub QueueBindAll()
    If gVim Is Nothing Then Exit Sub
    If Not gVim.Enabled Then Exit Sub
    If mBindAllScheduled Then Exit Sub
    mBindAllScheduled = True
    Application.OnTime Now, "'C_EventBatch.FlushBindAll'"
End Sub

'/*
' * Runs the actual BindAll on the next idle tick.
' */
Public Sub FlushBindAll()
    On Error Resume Next
    mBindAllScheduled = False
    If gVim Is Nothing Then Exit Sub
    If Not gVim.Enabled Then Exit Sub
    gVim.KeyMap.BindAll
End Sub
