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
    On Error Resume Next
    If Left$(f, 1) = "=" Then f = Mid$(f, 2)
    If InStr(1, f, "!") = 0 Then Exit Function

    Dim Wb As Workbook
    Dim ws As Worksheet
    Dim parts() As String, candidate As String
    Set Wb = cell.Parent.Parent

    parts = Split(f, "!")
    candidate = Replace$(parts(0), "'", "")
    candidate = Trim$(candidate)
    If InStr(1, candidate, "[") > 0 Then Exit Function

    If Len(candidate) = 0 Then Exit Function

    On Error Resume Next
    Set ws = Wb.Worksheets(candidate)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    If StrComp(ws.Name, cell.Parent.Name, vbTextCompare) <> 0 Then
        IsOtherSheetReference = True
    End If
End Function

