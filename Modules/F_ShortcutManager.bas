Attribute VB_Name = "F_ShortcutManager"
Option Explicit
Option Private Module

Private Const SHORTCUTS_FILE As String = "_vimxlamrc"
Private Const OVERRIDE_START As String = "'-- Vantage Shortcut Overrides (auto) --"
Private Const OVERRIDE_END As String = "'-- End Vantage Shortcut Overrides --"

Public Function ShortcutManager(Optional ByVal g As String) As Boolean
    On Error GoTo Fail

    If gVim Is Nothing Then
        Call StartVim
    End If

    If Not gVim Is Nothing Then
        If gVim.Mode.Current = MODE_CMDLINE Then
            gVim.Mode.Change MODE_NORMAL
        End If
    End If

    UF_ShortcutManager.ShowManager
    Exit Function

Fail:
    Call ErrorHandler("ShortcutManager")
End Function

Public Function ShortcutManager_BuildItems() As Collection
    Dim items As New Collection
    Dim overrides As Object
    Dim disabled As Object
    Dim catalog As Variant
    Dim i As Long

    If gVim Is Nothing Then
        Call StartVim
    End If

    ShortcutManager_LoadOverrides overrides, disabled
    catalog = ShortcutCatalog()
    For i = LBound(catalog, 1) To UBound(catalog, 1)
        Dim category As String
        Dim mapStr As String
        Dim lhs As String
        Dim rhs As String

        category = CStr(catalog(i, 1))
        mapStr = CStr(catalog(i, 2))

        If Not ShortcutManager_ParseMap(mapStr, lhs, rhs) Then
            GoTo ContinueRow
        End If

        Dim item As cls_ShortcutItem
        Set item = New cls_ShortcutItem
        item.Category = category
        item.Action = rhs
        item.DefaultKey = lhs
        item.CurrentKey = ShortcutManager_ResolveKey(lhs, rhs, overrides, disabled)
        items.Add item

ContinueRow:
    Next i

    ShortcutManager_AppendCmdItems items

    Set ShortcutManager_BuildItems = items
End Function

Private Sub ShortcutManager_AppendCmdItems(ByVal items As Collection)
    On Error Resume Next
    If gVim Is Nothing Then Exit Sub

    Dim cmdEntries As Collection
    Set cmdEntries = ShortcutManager_GetCmdEntries()

    Dim entry As Variant
    For Each entry In cmdEntries
        Dim keyText As String
        Dim actionText As String
        keyText = CStr(entry(0))
        actionText = CStr(entry(1))

        If Not ShortcutManager_KeyExists(items, keyText) Then
            Dim item As cls_ShortcutItem
            Set item = New cls_ShortcutItem
            item.Category = "Commands"
            item.Action = actionText
            item.DefaultKey = keyText
            item.CurrentKey = keyText
            items.Add item
        End If
    Next entry
End Sub

Private Function ShortcutManager_GetCmdEntries() As Collection
    Dim results As New Collection
    Dim cmdKeys As Variant
    Dim hadSuggest As Boolean
    Dim keyText As String
    Dim actionText As String
    Dim rawKey As String
    Dim displayKey As String
    Dim keyVal As Variant
    Dim i As Long

    Dim cmdMappings As Variant
    On Error Resume Next
    cmdMappings = gVim.KeyMap.ExportCommandMappings()
    If Err.Number = 0 Then
        If Not IsEmpty(cmdMappings) Then
            For i = LBound(cmdMappings, 1) To UBound(cmdMappings, 1)
                keyText = CStr(cmdMappings(i, 1))
                actionText = ShortcutManager_CleanAction(CStr(cmdMappings(i, 2)))
                If Len(keyText) > 0 And Len(actionText) > 0 Then
                    results.Add Array(keyText, actionText)
                End If
            Next i
        End If
    End If
    Err.Clear
    On Error GoTo 0

    If results.Count > 0 Then
        Set ShortcutManager_GetCmdEntries = results
        Exit Function
    End If

    On Error Resume Next
    cmdKeys = gVim.KeyMap.Suggest("", True)
    hadSuggest = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    If hadSuggest Then
        For Each keyVal In cmdKeys
            rawKey = CStr(keyVal)
            If Len(rawKey) = 0 Then GoTo ContinueKey

            actionText = ShortcutManager_CleanAction(gVim.KeyMap.Get_(rawKey, True))
            If Len(actionText) = 0 Then GoTo ContinueKey
            If LCase$(Left$(actionText, Len(SHOW_CMD_PROCEDURE))) = LCase$(SHOW_CMD_PROCEDURE) Then GoTo ContinueKey
            If actionText = DUMMY_PROCEDURE Then GoTo ContinueKey

            If Left$(rawKey, 1) = ":" Then
                displayKey = rawKey
            Else
                displayKey = ":" & rawKey
            End If
            results.Add Array(displayKey, actionText)
ContinueKey:
        Next keyVal
    End If

    If results.Count = 0 Then
        Dim mappings As Variant
        mappings = gVim.KeyMap.ExportMappings(True, True)
        If Not IsEmpty(mappings) Then
            For i = LBound(mappings, 1) To UBound(mappings, 1)
                keyText = CStr(mappings(i, 1))
                actionText = ShortcutManager_CleanAction(CStr(mappings(i, 2)))
                If Left$(keyText, 1) = ":" And Len(actionText) > 0 Then
                    results.Add Array(keyText, actionText)
                End If
            Next i
        End If
    End If

    If results.Count = 0 Then
        Dim catalog As Variant
        catalog = ShortcutCatalog()

        Dim j As Long
        For j = LBound(catalog, 1) To UBound(catalog, 1)
            Dim mapStr As String
            Dim lhs As String
            Dim rhs As String
            mapStr = CStr(catalog(j, 2))
            If ShortcutManager_ParseMap(mapStr, lhs, rhs) Then
                If Left$(lhs, 1) = ":" Then
                    results.Add Array(lhs, rhs)
                End If
            End If
        Next j
    End If

    Set ShortcutManager_GetCmdEntries = results
End Function

Private Function ShortcutManager_CleanAction(ByVal rawValue As String) As String
    Dim text As String
    text = Trim$(rawValue)
    If Len(text) = 0 Then Exit Function

    If (Left$(text, 1) = "'" And Right$(text, 1) = "'") Or (Left$(text, 1) = """" And Right$(text, 1) = """") Then
        text = Mid$(text, 2, Len(text) - 2)
    End If

    ShortcutManager_CleanAction = Trim$(text)
End Function

Private Function ShortcutManager_KeyExists(ByVal items As Collection, ByVal keyText As String) As Boolean
    Dim item As cls_ShortcutItem
    For Each item In items
        If StrComp(item.CurrentKey, keyText, vbBinaryCompare) = 0 Then
            ShortcutManager_KeyExists = True
            Exit Function
        End If
        If StrComp(item.DefaultKey, keyText, vbBinaryCompare) = 0 Then
            ShortcutManager_KeyExists = True
            Exit Function
        End If
    Next item
End Function

Private Function ShortcutManager_ItemExists(ByVal items As Collection, ByVal actionText As String) As Boolean
    Dim item As cls_ShortcutItem
    For Each item In items
        If StrComp(item.Action, actionText, vbBinaryCompare) = 0 Then
            ShortcutManager_ItemExists = True
            Exit Function
        End If
    Next item
End Function

Private Function ShortcutManager_ResolveKey(ByVal defaultKey As String, ByVal actionText As String, ByVal overrides As Object, ByVal disabled As Object) As String
    Dim keyText As String
    keyText = defaultKey

    If Not overrides Is Nothing Then
        If overrides.Exists(actionText) Then
            keyText = CStr(overrides(actionText))
        End If
    End If

    If Not disabled Is Nothing Then
        If disabled.Exists(defaultKey) Then
            keyText = ""
        End If
    End If

    ShortcutManager_ResolveKey = keyText
End Function

Public Sub ShortcutManager_ApplyChange(ByVal item As cls_ShortcutItem, ByVal newKey As String)
    On Error GoTo Fail
    If gVim Is Nothing Then Exit Sub
    If IsCmdItem(item) Then Exit Sub

    Dim oldKey As String
    oldKey = item.CurrentKey

    If Len(Trim$(newKey)) = 0 Then
        ShortcutManager_Disable item
        Exit Sub
    End If

    If Len(oldKey) > 0 Then
        If StrComp(oldKey, newKey, vbBinaryCompare) <> 0 Then
            gVim.KeyMap.Map "nunmap " & oldKey
        End If
    End If

    gVim.KeyMap.Map "nmap " & newKey & " " & item.Action
    item.CurrentKey = newKey
    gVim.KeyMap.BindAll
    ShortcutManager_SaveOverrides
    Exit Sub

Fail:
    Call ErrorHandler("ShortcutManager_ApplyChange")
End Sub

Public Sub ShortcutManager_Disable(ByVal item As cls_ShortcutItem)
    On Error GoTo Fail
    If gVim Is Nothing Then Exit Sub
    If IsCmdItem(item) Then Exit Sub

    If Len(Trim$(item.CurrentKey)) > 0 Then
        gVim.KeyMap.Map "nunmap " & item.CurrentKey
        item.CurrentKey = ""
    End If

    gVim.KeyMap.BindAll
    ShortcutManager_SaveOverrides
    Exit Sub

Fail:
    Call ErrorHandler("ShortcutManager_Disable")
End Sub

Public Sub ShortcutManager_ResetItem(ByVal item As cls_ShortcutItem)
    On Error GoTo Fail
    If gVim Is Nothing Then Exit Sub
    If IsCmdItem(item) Then Exit Sub

    If Len(Trim$(item.CurrentKey)) > 0 Then
        gVim.KeyMap.Map "nunmap " & item.CurrentKey
    End If

    gVim.KeyMap.Map "nmap " & item.DefaultKey & " " & item.Action
    item.CurrentKey = item.DefaultKey
    gVim.KeyMap.BindAll
    ShortcutManager_SaveOverrides
    Exit Sub

Fail:
    Call ErrorHandler("ShortcutManager_ResetItem")
End Sub

Public Sub ShortcutManager_ResetAll(ByVal items As Collection)
    On Error GoTo Fail
    If gVim Is Nothing Then Exit Sub

    Dim item As cls_ShortcutItem
    For Each item In items
        ShortcutManager_ResetItem item
    Next item

    gVim.KeyMap.BindAll
    ShortcutManager_SaveOverrides
    Exit Sub

Fail:
    Call ErrorHandler("ShortcutManager_ResetAll")
End Sub

Public Function ShortcutManager_FindByKey(ByVal items As Collection, ByVal keyText As String, Optional ByVal ignoreAction As String = "") As cls_ShortcutItem
    Dim item As cls_ShortcutItem
    For Each item In items
        If StrComp(item.CurrentKey, keyText, vbBinaryCompare) = 0 Then
            If Len(ignoreAction) = 0 Or StrComp(item.Action, ignoreAction, vbBinaryCompare) <> 0 Then
                Set ShortcutManager_FindByKey = item
                Exit Function
            End If
        End If
    Next item
End Function

Public Function ShortcutManager_ValidateKey(ByVal keyText As String, ByRef errMsg As String) As Boolean
    Dim trimmed As String
    trimmed = Trim$(keyText)

    If Len(trimmed) = 0 Then
        ShortcutManager_ValidateKey = True
        Exit Function
    End If

    If InStr(trimmed, " ") > 0 Then
        errMsg = "Keys cannot contain spaces."
        Exit Function
    End If

    If LCase$(Left$(trimmed, Len(KEY_CMD))) = KEY_CMD Then
        errMsg = "Command-line mappings are not edited here."
        Exit Function
    End If

    If gVim Is Nothing Then
        errMsg = "Vantage is not initialized."
        Exit Function
    End If

    On Error GoTo Fail
    Call gVim.KeyMap.VimToVBA(trimmed, KEY_SEPARATOR, True, VBASendkeys)
    ShortcutManager_ValidateKey = True
    Exit Function

Fail:
    errMsg = "Invalid key syntax."
    ShortcutManager_ValidateKey = False
End Function

Private Function ShortcutManager_ParseMap(ByVal mapStr As String, ByRef lhs As String, ByRef rhs As String) As Boolean
    Dim lower As String
    lower = LCase$(Trim$(mapStr))

    If Left$(lower, 4) <> "nmap" Then Exit Function

    Dim rest As String
    rest = Trim$(Mid$(mapStr, 5))
    If Len(rest) = 0 Then Exit Function

    Dim spacePos As Long
    spacePos = InStr(rest, " ")
    If spacePos = 0 Then Exit Function

    lhs = Trim$(Left$(rest, spacePos - 1))
    rhs = Trim$(Mid$(rest, spacePos + 1))

    If Len(lhs) = 0 Or Len(rhs) = 0 Then Exit Function

    If LCase$(Left$(lhs, Len(KEY_CMD))) = KEY_CMD Then
        lhs = ":" & Mid$(lhs, Len(KEY_CMD) + 1)
    End If

    ShortcutManager_ParseMap = True
End Function

Private Sub ShortcutManager_LoadOverrides(ByRef overrides As Object, ByRef disabled As Object)
    Dim filePath As String
    filePath = ShortcutManager_ConfigPath

    Set overrides = CreateObject("Scripting.Dictionary")
    Set disabled = CreateObject("Scripting.Dictionary")

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Sub

    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1, False)

    Dim inBlock As Boolean
    inBlock = False

    Do While Not ts.AtEndOfStream
        Dim line As String
        line = Trim$(ts.ReadLine)
        If Len(line) = 0 Then GoTo ContinueLine
        If line = OVERRIDE_START Then
            inBlock = True
            GoTo ContinueLine
        End If
        If line = OVERRIDE_END Then
            inBlock = False
            GoTo ContinueLine
        End If
        If Not inBlock Then GoTo ContinueLine
        If Left$(line, 1) = "'" Then GoTo ContinueLine

        Dim lcaseLine As String
        lcaseLine = LCase$(line)
        If Left$(lcaseLine, 5) = "nmap " Then
            Dim lhs As String
            Dim rhs As String
            If ShortcutManager_ParseMap(line, lhs, rhs) Then
                overrides(rhs) = lhs
            End If
        ElseIf Left$(lcaseLine, 7) = "nunmap " Then
            Dim keyText As String
            keyText = Trim$(Mid$(line, 8))
            If Len(keyText) > 0 Then
                disabled(keyText) = True
            End If
        End If

ContinueLine:
    Loop

    ts.Close
End Sub

Private Function ShortcutManager_ConfigPath() As String
    Dim basePath As String
    basePath = ThisWorkbook.Path

    If basePath Like "https://*" Then
        Dim sepIndex As Long
        If basePath Like "*my.sharepoint.com/*" Then
            sepIndex = InStr(basePath, "/Documents/")
            basePath = Mid$(basePath, sepIndex + 10)
        Else
            sepIndex = InStr(10, basePath, "/")
            sepIndex = InStr(sepIndex + 1, basePath, "/")
            basePath = Mid$(basePath, sepIndex)
        End If
        basePath = Environ$("OneDrive") & Replace(basePath, "/", "\\")
    End If

    ShortcutManager_ConfigPath = basePath & "\\" & SHORTCUTS_FILE
End Function

Public Sub ShortcutManager_SaveOverrides()
    On Error GoTo Fail

    Dim items As Collection
    Set items = UF_ShortcutManager.Items
    If items Is Nothing Then Exit Sub

    Dim filePath As String
    filePath = ShortcutManager_ConfigPath

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim lines As Collection
    Set lines = New Collection

    If fso.FileExists(filePath) Then
        Dim tsRead As Object
        Set tsRead = fso.OpenTextFile(filePath, 1, False)
        Do While Not tsRead.AtEndOfStream
            lines.Add tsRead.ReadLine
        Loop
        tsRead.Close
    End If

    Dim cleaned As Collection
    Set cleaned = New Collection

    Dim inBlock As Boolean
    inBlock = False

    Dim i As Long
    For i = 1 To lines.Count
        Dim line As String
        line = CStr(lines(i))
        If Trim$(line) = OVERRIDE_START Then
            inBlock = True
            GoTo ContinueLine
        End If
        If Trim$(line) = OVERRIDE_END Then
            inBlock = False
            GoTo ContinueLine
        End If
        If Not inBlock Then cleaned.Add line
ContinueLine:
    Next i

    cleaned.Add OVERRIDE_START

    Dim item As cls_ShortcutItem
    For Each item In items
        If StrComp(item.CurrentKey, item.DefaultKey, vbBinaryCompare) <> 0 Then
            cleaned.Add "nunmap " & item.DefaultKey
            If Len(Trim$(item.CurrentKey)) > 0 Then
                cleaned.Add "nmap " & item.CurrentKey & " " & item.Action
            End If
        End If
    Next item

    cleaned.Add OVERRIDE_END

    Dim tsWrite As Object
    Set tsWrite = fso.OpenTextFile(filePath, 2, True)
    For i = 1 To cleaned.Count
        tsWrite.WriteLine CStr(cleaned(i))
    Next i
    tsWrite.Close
    Exit Sub

Fail:
    Call ErrorHandler("ShortcutManager_SaveOverrides")
End Sub

Private Function IsCmdItem(ByVal item As cls_ShortcutItem) As Boolean
    If item Is Nothing Then Exit Function
    IsCmdItem = (Left$(item.DefaultKey, 1) = ":")
End Function
