Attribute VB_Name = "C_Util1"
Option Explicit
Option Private Module

Public Sub TimeClear()
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub
    engine.TimeClear
End Sub

Public Function GetQueryPerformanceTime(Optional ByVal vFormat As String = "0.0000") As Double
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    GetQueryPerformanceTime = engine.GetQueryPerformanceTime(vFormat)
End Function

Sub SetStatusBar(Optional ByVal str As String = "", _
                 Optional ByVal currentCount As Long = -1, _
                 Optional ByVal maximumCount As Long = -1, _
                 Optional ByVal percent As Double = -1, _
                 Optional ByVal numDigitsAfterDecimal As Byte = 0, _
                 Optional ByVal progressBar As Boolean = False, _
                 Optional ByVal countPerMax As Boolean = False)

    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub
    engine.SetStatusBar str, currentCount, maximumCount, percent, numDigitsAfterDecimal, progressBar, countPerMax
End Sub

Sub SetStatusBarTemporarily(ByVal str As String, _
                            ByVal miliseconds As Long, _
                   Optional ByVal disablePrefix As Boolean = False)

    Dim engine As Object
    Dim statusPrefix As String

    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub

    If Not gVim Is Nothing Then
        statusPrefix = gVim.Config.StatusPrefix
    End If

    engine.SetStatusBarTemporarily str, miliseconds, disablePrefix, statusPrefix
End Sub

Function RegExpMatch(ByVal str As String, ByVal matchPattern As String, _
            Optional ByVal isIgnoreCase As Boolean = False, _
            Optional ByVal isGlobal As Boolean = True, _
            Optional ByVal isMultiline As Boolean = False) As Boolean

    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    RegExpMatch = engine.RegExpMatch(str, matchPattern, isIgnoreCase, isGlobal, isMultiline)
End Function

Function RegExpSearch(ByVal str As String, ByVal matchPattern As String, _
             Optional ByVal isIgnoreCase As Boolean = False, _
             Optional ByVal isGlobal As Boolean = True, _
             Optional ByVal isMultiline As Boolean = False) As String

    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    RegExpSearch = engine.RegExpSearch(str, matchPattern, isIgnoreCase, isGlobal, isMultiline)
End Function

Function RegExpReplace(ByVal str As String, ByVal matchPattern As String, ByVal replaceStr As String, _
              Optional ByVal isIgnoreCase As Boolean = False, _
              Optional ByVal isGlobal As Boolean = True, _
              Optional ByVal isMultiline As Boolean = False) As String

    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    RegExpReplace = engine.RegExpReplace(str, matchPattern, replaceStr, isIgnoreCase, isGlobal, isMultiline)
End Function

Function StartsWith(ByRef str As String, ByVal prefixes As Variant) As Boolean
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    StartsWith = engine.StartsWith(str, prefixes)
End Function

Function EndsWith(ByRef str As String, ByVal suffixes As Variant) As Boolean
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    EndsWith = engine.EndsWith(str, suffixes)
End Function

Function GetWorkbookIndex(ByVal targetWorkbook As Workbook) As Long
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    GetWorkbookIndex = engine.GetWorkbookIndex(targetWorkbook)
End Function

Function IsSheetExists(ByVal targetSheetName As String) As Boolean
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    IsSheetExists = engine.IsSheetExists(targetSheetName)
End Function

Function GetVisibleSheetsCount() As Long
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    GetVisibleSheetsCount = engine.GetVisibleSheetsCount()
End Function

Function DirGrob(ByVal folderPath As String) As Collection
    Dim engine As Object
    Dim items As Variant
    Dim result As New Collection
    Dim i As Long

    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then
        Set DirGrob = result
        Exit Function
    End If

    items = engine.DirGrob(folderPath)

    If IsArray(items) Then
        For i = LBound(items) To UBound(items)
            result.Add items(i)
        Next i
    End If

    Set DirGrob = result
End Function

Function GetAbsolutePath(ByRef cwd As String, ByRef relativePath As String) As String
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    GetAbsolutePath = engine.GetAbsolutePath(cwd, relativePath)
End Function

Function ResolvePath(ByVal strPath As String) As String
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    ResolvePath = engine.ResolvePath(strPath)
End Function

Function HexColorCodeToLong(ByVal colorCode As String) As Long
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    HexColorCodeToLong = engine.HexColorCodeToLong(colorCode)
End Function

Function ColorCodeToHex(ByVal colorCode As Long) As String
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    ColorCodeToHex = engine.ColorCodeToHex(colorCode)
End Function

Function IsJISKeyboardLayout() As Boolean
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    IsJISKeyboardLayout = engine.IsJISKeyboardLayout()
End Function

Public Function Union2(ParamArray ArgList() As Variant) As Range
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    Set Union2 = engine.Union2(ArgList)
End Function

Public Function Intersect2(ParamArray ArgList() As Variant) As Range
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    Set Intersect2 = engine.Intersect2(ArgList)
End Function

Public Function Except2(ByRef SourceRange As Variant, ParamArray ArgList() As Variant) As Range
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    Set Except2 = engine.Except2(SourceRange, ArgList)
End Function

Public Function Invert2(ByRef SourceRange As Variant) As Range
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    Set Invert2 = engine.Invert2(SourceRange)
End Function

Public Function IsRangeValid(ByVal candidate As Range) As Boolean
    Dim engine As Object
    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function
    IsRangeValid = engine.IsRangeValid(candidate)
End Function

Sub DebugPrint(ByVal str As String, Optional ByVal funcName As String = "")
    Dim engine As Object
    Dim debugMode As Boolean
    Dim statusPrefix As String

    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Sub

    If Not gVim Is Nothing Then
        debugMode = gVim.Config.DebugMode
        statusPrefix = gVim.Config.StatusPrefix
    End If

    If Not debugMode Then Exit Sub
    engine.DebugPrint str, funcName, debugMode, statusPrefix
End Sub

Function ErrorHandler(Optional ByVal funcName As String = "") As Boolean
    Dim engine As Object
    Dim statusPrefix As String

    If Err.Number = 0 Then Exit Function

    On Error Resume Next
    Set engine = NetAddin()
    If engine Is Nothing Then Exit Function

    If Not gVim Is Nothing Then
        statusPrefix = gVim.Config.StatusPrefix
    End If

    If engine.ErrorHandler(Err.Number, Err.Description, funcName, statusPrefix) Then
        Err.Clear
        ErrorHandler = True
    End If
End Function
