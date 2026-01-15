Attribute VB_Name = "C_Info"
Option Explicit
Option Private Module

Public Function ShowCommandInfo(Optional ByVal g As String) As Boolean
    On Error GoTo CleanFail

    If gVim Is Nothing Then
        Call StartVim
    End If

    Dim mappings As Variant
    mappings = gVim.KeyMap.ExportMappings(True, True)
    If IsEmpty(mappings) Then
        Call SetStatusBarTemporarily("No command mappings found.", 2000)
        Exit Function
    End If

    Dim infoText As String
    infoText = BuildCommandInfoText(mappings)
    If Len(infoText) = 0 Then
        Call SetStatusBarTemporarily("No command mappings found.", 2000)
        Exit Function
    End If

    Dim errText As String
    On Error Resume Next
    UF_Info.ShowInfo infoText, "Vantage Commands"
    If Err.Number <> 0 Then
        errText = Err.Description
        Err.Clear
        On Error GoTo CleanFail
        MsgBox infoText, vbInformation, "Vantage Commands"
        If Len(errText) > 0 Then
            Call SetStatusBarTemporarily("Info form failed: " & errText, 3000)
        End If
        Exit Function
    End If
    On Error GoTo CleanFail

    ShowCommandInfo = False
    Exit Function

CleanFail:
    Call ErrorHandler("ShowCommandInfo")
End Function

Private Function BuildCommandInfoText(ByRef mappings As Variant) As String
    Dim items() As String
    Dim itemCount As Long

    Dim maxKey As Long
    Dim maxAction As Long
    Dim i As Long

    For i = LBound(mappings, 1) To UBound(mappings, 1)
        Dim keyText As String
        Dim actionText As String
        Dim actionDisplay As String
        Dim descText As String

        keyText = CStr(mappings(i, 1))
        actionText = CStr(mappings(i, 2))
        actionDisplay = StripOuterQuotes(actionText)
        descText = gVim.Help.GetText(actionText)
        If descText = actionText Then
            descText = ""
        End If
        descText = Replace(descText, vbCr, " ")
        descText = Replace(descText, vbLf, " ")

        itemCount = itemCount + 1
        ReDim Preserve items(1 To itemCount)
        items(itemCount) = keyText & vbTab & actionDisplay & vbTab & descText

        If Len(keyText) > maxKey Then maxKey = Len(keyText)
        If Len(actionDisplay) > maxAction Then maxAction = Len(actionDisplay)
    Next i

    If itemCount = 0 Then
        Exit Function
    End If

    Call SortInfoItems(items, 1, itemCount)

    Dim keyWidth As Long
    Dim actionWidth As Long
    keyWidth = ClampWidth(maxKey, 10, 24)
    actionWidth = ClampWidth(maxAction, 12, 34)

    Dim output As String
    output = PadRight("Key", keyWidth) & "  " & PadRight("Action", actionWidth) & "  " & "Description" & vbCrLf
    output = output & String$(keyWidth + actionWidth + 2 + 12, "-") & vbCrLf

    Dim entry As String
    For i = 1 To itemCount
        entry = items(i)
        Dim parts() As String
        parts = Split(entry, vbTab)

        keyText = parts(0)
        actionText = parts(1)
        descText = ""
        If UBound(parts) >= 2 Then
            descText = parts(2)
        End If

        output = output & PadRight(TruncateText(keyText, keyWidth), keyWidth)
        output = output & "  " & PadRight(TruncateText(actionText, actionWidth), actionWidth)
        output = output & "  " & descText & vbCrLf
    Next i

    BuildCommandInfoText = output
End Function

Private Sub SortInfoItems(ByRef items() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As String
    Dim temp As String

    i = first
    j = last
    pivot = items((first + last) \ 2)

    Do While i <= j
        Do While StrComp(items(i), pivot, vbTextCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(items(j), pivot, vbTextCompare) > 0
            j = j - 1
        Loop
        If i <= j Then
            temp = items(i)
            items(i) = items(j)
            items(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then
        Call SortInfoItems(items, first, j)
    End If
    If i < last Then
        Call SortInfoItems(items, i, last)
    End If
End Sub

Private Function ClampWidth(ByVal value As Long, ByVal minValue As Long, ByVal maxValue As Long) As Long
    If value < minValue Then value = minValue
    If value > maxValue Then value = maxValue
    ClampWidth = value
End Function

Private Function PadRight(ByVal text As String, ByVal width As Long) As String
    If Len(text) >= width Then
        PadRight = text
    Else
        PadRight = text & Space$(width - Len(text))
    End If
End Function

Private Function TruncateText(ByVal text As String, ByVal maxLen As Long) As String
    If Len(text) <= maxLen Then
        TruncateText = text
    ElseIf maxLen <= 3 Then
        TruncateText = Left$(text, maxLen)
    Else
        TruncateText = Left$(text, maxLen - 3) & "..."
    End If
End Function

Private Function StripOuterQuotes(ByVal text As String) As String
    Dim trimmed As String
    trimmed = Trim$(text)
    If Len(trimmed) >= 2 Then
        If Left$(trimmed, 1) = "'" And Right$(trimmed, 1) = "'" Then
            trimmed = Mid$(trimmed, 2, Len(trimmed) - 2)
        End If
    End If
    StripOuterQuotes = trimmed
End Function
