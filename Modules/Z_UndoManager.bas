Attribute VB_Name = "Z_UndoManager"
Option Explicit

Private gUndoActive As Boolean
Private gUndoSkipNext As Boolean
Private Const UNDO_STACK_MAX As Long = 2
Private Const UNDO_CACHE_SLOTS As Long = 3
' 0 = no limit (can be heavy for very large selections)
Private Const UNDO_MAX_CELLS As Long = 0
Private Const UNDO_CAPTURE_ENABLED As Boolean = False

Private mUndoStack As Collection
Private mRedoStack As Collection
Private mPendingSnapshot As Object

Public Sub UndoPrepareForCommand(ByVal cmd As String)
    If Not UNDO_CAPTURE_ENABLED Then
        gUndoSkipNext = False
        Exit Sub
    End If

    If gUndoSkipNext Then
        gUndoSkipNext = False
        Exit Sub
    End If

    If Not ShouldCaptureUndo(cmd) Then Exit Sub

    EnsureUndoStacks
    If Not mRedoStack Is Nothing Then
        If mRedoStack.Count > 0 Then ClearRedoStack
    End If

    Set mPendingSnapshot = Nothing
    Dim target As Range
    Set target = GetUndoTargetRange()
    If target Is Nothing Then Exit Sub

    Dim slot As Long
    slot = GetFreeCacheSlot()
    If slot = 0 Then Exit Sub

    Set mPendingSnapshot = CaptureSnapshotFromRange(target, slot)
    If mPendingSnapshot Is Nothing Then Exit Sub

    StartCustomUndo "macro: " & cmd
End Sub

Public Sub UndoPrepareForRange(ByVal target As Range, Optional ByVal cmd As String = "")
    If Not UNDO_CAPTURE_ENABLED Then
        gUndoSkipNext = False
        Exit Sub
    End If

    If gUndoSkipNext Then
        gUndoSkipNext = False
        Exit Sub
    End If

    If target Is Nothing Then Exit Sub

    EnsureUndoStacks
    If Not mRedoStack Is Nothing Then
        If mRedoStack.Count > 0 Then ClearRedoStack
    End If

    Set mPendingSnapshot = Nothing

    Dim slot As Long
    slot = GetFreeCacheSlot()
    If slot = 0 Then Exit Sub

    Set mPendingSnapshot = CaptureSnapshotFromRange(target, slot)
    If mPendingSnapshot Is Nothing Then Exit Sub

    If Not gUndoActive Then
        If Len(cmd) > 0 Then
            StartCustomUndo "macro: " & cmd
        Else
            StartCustomUndo "macro"
        End If
    End If
End Sub

Public Sub UndoFinalizeForCommand()
    If Not mPendingSnapshot Is Nothing Then
        PushUndoSnapshot mPendingSnapshot, True
        Set mPendingSnapshot = Nothing
    End If
    FinalizeCustomUndo
End Sub

Public Sub UndoAbortForCommand()
    Set mPendingSnapshot = Nothing
    FinalizeCustomUndo
End Sub

Public Sub UndoClearSnapshot()
    gUndoActive = False
    gUndoSkipNext = False
    Set mPendingSnapshot = Nothing
    ClearUndoRedoStacks
End Sub

Public Sub UndoSuppressForNextCommand()
    gUndoSkipNext = True
End Sub

Public Function UndoShouldCapture(ByVal cmd As String) As Boolean
    UndoShouldCapture = ShouldCaptureUndo(cmd)
End Function

Private Function ShouldCaptureUndo(ByVal cmd As String) As Boolean
    Dim lower As String
    Dim token As Variant
    lower = Trim$(LCase$(cmd))
    If lower = "" Then Exit Function

    If Left$(lower, 1) = "'" Then
        lower = Mid$(lower, 2)
        lower = Trim$(lower)
        If lower = "" Then Exit Function
    End If
    If InStr(lower, "undo") > 0 Or InStr(lower, "redo") > 0 Then Exit Function

    Dim skipTokens As Variant
    skipTokens = Array("move", "scroll", "toggle", "show", "jump", "focus", "select", "start", "stop", "center", "undo_c", "keystroke")
    For Each token In skipTokens
        If Left$(lower, Len(token)) = token Then Exit Function
    Next token

    ShouldCaptureUndo = True
End Function

Public Function UndoPerform(Optional ByVal times As Long = 1) As Boolean
    Dim i As Long
    Dim ok As Boolean

    If times < 1 Then times = 1
    For i = 1 To times
        ok = UndoOnce()
        If Not ok Then
            On Error Resume Next
            Application.Undo
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            On Error GoTo 0
        End If
    Next i

    UndoPerform = False
End Function

Public Function RedoPerform(Optional ByVal times As Long = 1) As Boolean
    Dim i As Long
    Dim ok As Boolean

    If times < 1 Then times = 1
    For i = 1 To times
        ok = RedoOnce()
        If Not ok Then
            On Error Resume Next
            Application.CommandBars.ExecuteMso "Redo"
            If Err.Number <> 0 Then
                Err.Clear
                Exit For
            End If
            On Error GoTo 0
        End If
    Next i

    RedoPerform = False
End Function

Private Function UndoOnce() As Boolean
    EnsureUndoStacks
    If mUndoStack.Count = 0 Then Exit Function

    Dim snap As Object
    Set snap = mUndoStack.Item(1)

    Dim slot As Long
    slot = GetFreeCacheSlot()
    If slot = 0 Then Exit Function

    Dim redoSnap As Object
    Set redoSnap = CaptureSnapshotFromSnapshot(snap, slot)
    If redoSnap Is Nothing Then Exit Function

    Set snap = PopUndoSnapshot()
    If snap Is Nothing Then Exit Function
    If Not RestoreSnapshot(snap) Then Exit Function

    PushRedoSnapshot redoSnap
    UndoOnce = True
End Function

Private Function RedoOnce() As Boolean
    EnsureUndoStacks
    If mRedoStack.Count = 0 Then Exit Function

    Dim snap As Object
    Set snap = mRedoStack.Item(1)

    Dim slot As Long
    slot = GetFreeCacheSlot()
    If slot = 0 Then Exit Function

    Dim undoSnap As Object
    Set undoSnap = CaptureSnapshotFromSnapshot(snap, slot)
    If undoSnap Is Nothing Then Exit Function

    Set snap = PopRedoSnapshot()
    If snap Is Nothing Then Exit Function
    If Not RestoreSnapshot(snap) Then Exit Function

    PushUndoSnapshot undoSnap, False
    RedoOnce = True
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

Private Sub EnsureUndoStacks()
    If mUndoStack Is Nothing Then Set mUndoStack = New Collection
    If mRedoStack Is Nothing Then Set mRedoStack = New Collection
End Sub

Private Sub ClearUndoRedoStacks()
    Set mUndoStack = New Collection
    Set mRedoStack = New Collection
End Sub

Private Sub ClearRedoStack()
    Set mRedoStack = New Collection
End Sub

Private Function GetUndoTargetRange() As Range
    On Error Resume Next
    If TypeName(Selection) = "Range" Then
        Set GetUndoTargetRange = Selection
    Else
        Dim win As Window
        Dim rangeSel As Range
        Set win = Application.ActiveWindow
        If Not win Is Nothing Then
            Set rangeSel = win.RangeSelection
            If Not rangeSel Is Nothing Then
                Set GetUndoTargetRange = rangeSel
                GoTo CleanExit
            End If
        End If

        If Not ActiveCell Is Nothing Then
            Set GetUndoTargetRange = ActiveCell
        End If
    End If
CleanExit:
    On Error GoTo 0
End Function

Private Sub PushUndoSnapshot(ByVal snap As Object, ByVal clearRedo As Boolean)
    EnsureUndoStacks
    If clearRedo Then ClearRedoStack

    mUndoStack.Add snap, , 1
    If mUndoStack.Count > UNDO_STACK_MAX Then
        mUndoStack.Remove mUndoStack.Count
    End If
End Sub

Private Sub PushRedoSnapshot(ByVal snap As Object)
    EnsureUndoStacks
    mRedoStack.Add snap, , 1
    If mRedoStack.Count > UNDO_STACK_MAX Then
        mRedoStack.Remove mRedoStack.Count
    End If
End Sub

Private Function PopUndoSnapshot() As Object
    EnsureUndoStacks
    If mUndoStack.Count = 0 Then Exit Function
    Set PopUndoSnapshot = mUndoStack.Item(1)
    mUndoStack.Remove 1
End Function

Private Function PopRedoSnapshot() As Object
    EnsureUndoStacks
    If mRedoStack.Count = 0 Then Exit Function
    Set PopRedoSnapshot = mRedoStack.Item(1)
    mRedoStack.Remove 1
End Function

Private Function GetFreeCacheSlot() As Long
    Dim used(1 To UNDO_CACHE_SLOTS) As Boolean
    Dim i As Long

    EnsureUndoStacks
    MarkSlotsUsed used

    For i = 1 To UNDO_CACHE_SLOTS
        If Not used(i) Then
            GetFreeCacheSlot = i
            Exit Function
        End If
    Next i
End Function

Private Sub MarkSlotsUsed(ByRef used() As Boolean)
    Dim snap As Object
    Dim slot As Long
    Dim i As Long

    If Not mPendingSnapshot Is Nothing Then
        slot = GetSnapshotSlot(mPendingSnapshot)
        If slot >= 1 And slot <= UNDO_CACHE_SLOTS Then used(slot) = True
    End If

    If Not mUndoStack Is Nothing Then
        For i = 1 To mUndoStack.Count
            Set snap = mUndoStack.Item(i)
            slot = GetSnapshotSlot(snap)
            If slot >= 1 And slot <= UNDO_CACHE_SLOTS Then used(slot) = True
        Next i
    End If

    If Not mRedoStack Is Nothing Then
        For i = 1 To mRedoStack.Count
            Set snap = mRedoStack.Item(i)
            slot = GetSnapshotSlot(snap)
            If slot >= 1 And slot <= UNDO_CACHE_SLOTS Then used(slot) = True
        Next i
    End If
End Sub

Private Function GetSnapshotSlot(ByVal snap As Object) As Long
    On Error Resume Next
    GetSnapshotSlot = CLng(snap("CacheSlot"))
    On Error GoTo 0
End Function

Private Function CaptureSnapshotFromRange(ByVal target As Range, ByVal slot As Long) As Object
    Dim prevScreen As Boolean
    Dim prevEvents As Boolean
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo CleanFail
    If target Is Nothing Then Exit Function

    If UNDO_MAX_CELLS > 0 Then
        If target.CountLarge > UNDO_MAX_CELLS Then Exit Function
    End If

    Dim ws As Worksheet
    Dim wb As Workbook
    Set ws = target.Worksheet
    Set wb = ws.Parent

    Dim cacheWs As Worksheet
    Set cacheWs = EnsureUndoCacheSheet(slot)
    cacheWs.Cells.Clear

    Dim areas As Collection
    Set areas = New Collection

    Dim nextRow As Long
    nextRow = 1

    Dim area As Range
    For Each area In target.Areas
        Dim rows As Long
        Dim cols As Long
        rows = area.Rows.Count
        cols = area.Columns.Count

        Dim cacheRange As Range
        Set cacheRange = cacheWs.Cells(nextRow, 1).Resize(rows, cols)
        area.Copy Destination:=cacheRange

        Dim areaInfo As Object
        Set areaInfo = CreateObject("Scripting.Dictionary")
        areaInfo.Add "TargetAddress", area.Address(True, True, xlA1, False)
        areaInfo.Add "CacheAddress", cacheRange.Address(True, True, xlA1, False)
        areaInfo.Add "Rows", rows
        areaInfo.Add "Cols", cols
        areas.Add areaInfo

        nextRow = nextRow + rows + 1
    Next area

    Dim snap As Object
    Set snap = CreateObject("Scripting.Dictionary")
    snap.Add "WorkbookName", wb.Name
    snap.Add "WorkbookFullName", wb.FullName
    snap.Add "SheetName", ws.Name
    snap.Add "Areas", areas
    snap.Add "CacheSlot", slot

    Application.CutCopyMode = False
    Set CaptureSnapshotFromRange = snap
    GoTo CleanExit

CleanFail:
    Application.CutCopyMode = False
CleanExit:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
End Function

Private Function CaptureSnapshotFromSnapshot(ByVal src As Object, ByVal slot As Long) As Object
    Dim prevScreen As Boolean
    Dim prevEvents As Boolean
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo CleanFail
    If src Is Nothing Then Exit Function

    Dim wb As Workbook
    Set wb = ResolveSnapshotWorkbook(src)
    If wb Is Nothing Then Exit Function

    Dim ws As Worksheet
    Set ws = wb.Worksheets(CStr(src("SheetName")))
    If ws Is Nothing Then Exit Function

    Dim cacheWs As Worksheet
    Set cacheWs = EnsureUndoCacheSheet(slot)
    cacheWs.Cells.Clear

    Dim areas As Collection
    Set areas = New Collection

    Dim nextRow As Long
    nextRow = 1

    Dim areaInfo As Object
    For Each areaInfo In src("Areas")
        Dim addr As String
        addr = CStr(areaInfo("TargetAddress"))

        Dim target As Range
        Set target = ws.Range(addr)

        Dim rows As Long
        Dim cols As Long
        rows = target.Rows.Count
        cols = target.Columns.Count

        Dim cacheRange As Range
        Set cacheRange = cacheWs.Cells(nextRow, 1).Resize(rows, cols)
        target.Copy Destination:=cacheRange

        Dim newInfo As Object
        Set newInfo = CreateObject("Scripting.Dictionary")
        newInfo.Add "TargetAddress", addr
        newInfo.Add "CacheAddress", cacheRange.Address(True, True, xlA1, False)
        newInfo.Add "Rows", rows
        newInfo.Add "Cols", cols
        areas.Add newInfo

        nextRow = nextRow + rows + 1
    Next areaInfo

    Dim snap As Object
    Set snap = CreateObject("Scripting.Dictionary")
    snap.Add "WorkbookName", CStr(src("WorkbookName"))
    snap.Add "WorkbookFullName", CStr(src("WorkbookFullName"))
    snap.Add "SheetName", CStr(src("SheetName"))
    snap.Add "Areas", areas
    snap.Add "CacheSlot", slot

    Application.CutCopyMode = False
    Set CaptureSnapshotFromSnapshot = snap
    GoTo CleanExit

CleanFail:
    Application.CutCopyMode = False
CleanExit:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
End Function

Private Function RestoreSnapshot(ByVal snap As Object) As Boolean
    On Error GoTo CleanFail
    If snap Is Nothing Then Exit Function

    Dim wb As Workbook
    Set wb = ResolveSnapshotWorkbook(snap)
    If wb Is Nothing Then Exit Function

    Dim ws As Worksheet
    Set ws = wb.Worksheets(CStr(snap("SheetName")))
    If ws Is Nothing Then Exit Function

    Dim cacheWs As Worksheet
    Dim slot As Long
    slot = GetSnapshotSlot(snap)
    Set cacheWs = EnsureUndoCacheSheet(slot)

    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim areaInfo As Object
    For Each areaInfo In snap("Areas")
        Dim target As Range
        Dim cacheRange As Range
        Set target = ws.Range(CStr(areaInfo("TargetAddress")))
        Set cacheRange = cacheWs.Range(CStr(areaInfo("CacheAddress")))
        cacheRange.Copy Destination:=target
    Next areaInfo

    Application.CutCopyMode = False
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen

    RestoreSnapshot = True
    Exit Function

CleanFail:
    Application.CutCopyMode = False
End Function

Private Function ResolveSnapshotWorkbook(ByVal snap As Object) As Workbook
    On Error Resume Next
    Dim fullName As String
    Dim nameOnly As String

    fullName = CStr(snap("WorkbookFullName"))
    nameOnly = CStr(snap("WorkbookName"))

    If Len(fullName) > 0 Then
        Set ResolveSnapshotWorkbook = Application.Workbooks(fullName)
        If Not ResolveSnapshotWorkbook Is Nothing Then Exit Function
    End If

    If Len(nameOnly) > 0 Then
        Set ResolveSnapshotWorkbook = Application.Workbooks(nameOnly)
    End If
    On Error GoTo 0
End Function

Private Function EnsureUndoCacheSheet(ByVal slot As Long) As Worksheet
    Dim name As String
    name = "VantageUndoCache" & CStr(slot)

    On Error Resume Next
    Set EnsureUndoCacheSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0

    If EnsureUndoCacheSheet Is Nothing Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = name
        ws.Visible = xlSheetVeryHidden
        Set EnsureUndoCacheSheet = ws
    Else
        EnsureUndoCacheSheet.Visible = xlSheetVeryHidden
    End If
End Function
