Attribute VB_Name = "Z_ColorAux"
Option Explicit

Public Sub ColorCell(ByVal c As Range)
    Const DEFAULT_FONT_COLOR As Long = vbBlack
    On Error GoTo CleanExit

    If c Is Nothing Then Exit Sub
    If IsError(c.value) Then Exit Sub

    Dim cellText As String
    cellText = CStr(c.value)

    If Len(Trim$(cellText)) = 0 Then
        If c.Font.Color <> DEFAULT_FONT_COLOR Then c.Font.Color = DEFAULT_FONT_COLOR
        Exit Sub
    End If

    If Not IsNumeric(c.value) Then Exit Sub

    Dim f As String
    Dim desiredColor As Long

    If c.hasFormula Then
        f = c.Formula2
        If IsExternalReference(f) Then
            desiredColor = RGB(120, 33, 112)    ' external workbook (#782170)
        ElseIf IsOtherSheetReference(f, c) Then
            desiredColor = RGB(0, 128, 0)       ' other sheet
        Else
            desiredColor = RGB(0, 0, 0)         ' same sheet
        End If
    Else
        desiredColor = RGB(0, 0, 255)           ' hard-coded numeric
    End If

    ' Avoid redundant redraws
    If c.Font.Color <> desiredColor Then
        c.Font.Color = desiredColor
    End If
    Exit Sub

CleanExit:
    On Error Resume Next
    If Not c Is Nothing Then
        If c.Font.Color <> DEFAULT_FONT_COLOR Then
            c.Font.Color = DEFAULT_FONT_COLOR
        End If
    End If
End Sub

Private Function IsExternalReference(ByVal f As String) As Boolean
    IsExternalReference = (InStr(1, f, "[") > 0 And InStr(1, f, "]") > 0)
End Function

Private Function IsOtherSheetReference(ByVal f As String, ByVal cell As Range) As Boolean
    On Error GoTo CleanExit
    If cell Is Nothing Then Exit Function

    Dim formulaText As String
    formulaText = f
    If Left$(formulaText, 1) = "=" Then formulaText = Mid$(formulaText, 2)
    If InStr(1, formulaText, "!") = 0 Then Exit Function

    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = "('([^']+)'|[A-Za-z0-9_]+)!"
    reg.Global = True
    reg.IgnoreCase = True

    Dim matches As Object
    Set matches = reg.Execute(formulaText)
    If matches.Count = 0 Then Exit Function

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = cell.Parent.Parent

    Dim matchItem As Object
    Dim candidate As String
    For Each matchItem In matches
        candidate = NormalizeSheetToken(CStr(matchItem.value))
        If Len(candidate) = 0 Then GoTo ContinueLoop
        If InStr(1, candidate, "[") > 0 Or InStr(1, candidate, "]") > 0 Then GoTo ContinueLoop

        On Error Resume Next
        Set ws = wb.Worksheets(candidate)
        On Error GoTo 0

        If Not ws Is Nothing Then
            If StrComp(ws.Name, cell.Parent.Name, vbTextCompare) <> 0 Then
                IsOtherSheetReference = True
                Exit Function
            End If
        End If

ContinueLoop:
        Set ws = Nothing
    Next matchItem

CleanExit:
End Function

Private Function NormalizeSheetToken(ByVal token As String) As String
    token = Trim$(token)
    If Len(token) = 0 Then Exit Function

    If Right$(token, 1) = "!" Then
        token = Left$(token, Len(token) - 1)
    End If
    token = Trim$(token)
    If Len(token) = 0 Then Exit Function

    If Left$(token, 1) = "'" And Right$(token, 1) = "'" Then
        token = Mid$(token, 2, Len(token) - 2)
        token = Replace$(token, "''", "'")
    End If

    NormalizeSheetToken = token
End Function

