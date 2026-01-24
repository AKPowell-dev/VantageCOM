Attribute VB_Name = "F_FunctionalMap"
Option Explicit
Option Private Module

Private gFunctionalMapActive As Boolean
Private gFunctionalMapOriginal As Object
Private gFunctionalMapBaseColor As Object

Private Const MAX_MAP_CELLS As Long = 125000
Private Const AUTO_COLOR_DATES As Boolean = True

Public Function FunctionalMap(Optional ByVal g As String) As Boolean
    Dim rng As Range
    On Error GoTo CleanFail

    HideCmdLineIfVisible

    If TypeName(Selection) <> "Range" Then Exit Function
    Set rng = Selection
    If rng Is Nothing Then Exit Function

    If gFunctionalMapActive Then
        FunctionalMapClear
        FunctionalMap = False
        Exit Function
    End If

    FunctionalMapApply rng
    FunctionalMap = False
    Exit Function

CleanFail:
    Call ErrorHandler("FunctionalMap")
End Function

Private Sub FunctionalMapApply(ByVal rng As Range)
    Dim app As Application
    Dim prevScreen As Boolean
    Dim prevEvents As Boolean
    Dim prevCalc As XlCalculation
    Dim area As Range
    Dim cell As Range
    Dim baseColor As Variant

    On Error GoTo CleanFail
    Set app = Application
    prevScreen = app.ScreenUpdating
    prevEvents = app.EnableEvents
    prevCalc = app.Calculation

    app.ScreenUpdating = False
    app.EnableEvents = False

    If rng.Cells.CountLarge > MAX_MAP_CELLS Then
        If MsgBox("Selection is large. Continue with Functional Map?", vbOKCancel + vbExclamation, "Functional Map") = vbCancel Then
            GoTo CleanUp
        End If
    End If

    Set gFunctionalMapOriginal = CreateObject("Scripting.Dictionary")
    Set gFunctionalMapBaseColor = CreateObject("Scripting.Dictionary")

    For Each area In rng.Areas
        For Each cell In area.Cells
            If CellHasValue(cell) Then
                SaveOriginalCellFormat cell
                baseColor = ApplyFunctionalMapToCell(cell)
                If Not IsEmpty(baseColor) Then
                    gFunctionalMapBaseColor(CellKey(cell)) = baseColor
                End If
            End If
        Next cell
    Next area

    ApplyRowFormulaDifferences rng

    gFunctionalMapActive = True

CleanUp:
    app.EnableEvents = prevEvents
    app.Calculation = prevCalc
    app.ScreenUpdating = prevScreen
    Exit Sub

CleanFail:
    Call ErrorHandler("FunctionalMapApply")
    Resume CleanUp
End Sub

Private Sub FunctionalMapClear()
    On Error Resume Next
    If Not gFunctionalMapOriginal Is Nothing Then
        Dim key As Variant
        Dim arr As Variant
        Dim cell As Range
        For Each key In gFunctionalMapOriginal.Keys
            Set cell = RangeFromKey(CStr(key))
            If Not cell Is Nothing Then
                arr = gFunctionalMapOriginal(key)
                cell.Interior.Pattern = arr(0)
                cell.Interior.PatternColor = arr(1)
                cell.Interior.Color = arr(2)
            End If
        Next key
    End If
    Set gFunctionalMapOriginal = Nothing
    Set gFunctionalMapBaseColor = Nothing
    gFunctionalMapActive = False
    HideCmdLineIfVisible
End Sub

Private Function ApplyFunctionalMapToCell(ByVal cell As Range) As Variant
    Dim clr As Variant
    Dim bSet As Boolean
    Dim colorVal As Variant
    Dim formulaText As String

    ApplyFunctionalMapToCell = Empty

    If cell.HasFormula Then
        formulaText = CStr(cell.Formula)
        colorVal = GetAutoColor(6)
        If Not IsEmpty(colorVal) Then
            If HasDataFunction(formulaText) Then
                SetPattern cell, xlPatternGray50, CLng(colorVal)
                ApplyFunctionalMapToCell = colorVal
                Exit Function
            End If
        End If
    End If

    If HasHyperlink(cell) Then
        colorVal = GetAutoColor(5)
        If Not IsEmpty(colorVal) Then
            SetPattern cell, xlPatternGray50, CLng(colorVal)
            ApplyFunctionalMapToCell = colorVal
            Exit Function
        End If
    End If

    CheckForOffSheetReferences cell, clr, bSet
    If Not bSet Then
        CheckForPartialInputs cell, clr, bSet
    End If
    If bSet Then
        SetPattern cell, xlPatternGray50, CLng(clr)
        ApplyFunctionalMapToCell = clr
        Exit Function
    End If

    If cell.HasFormula Then
        colorVal = GetAutoColor(2)
        If Not IsEmpty(colorVal) Then
            SetPattern cell, xlPatternGray50, CLng(colorVal)
            ApplyFunctionalMapToCell = colorVal
        End If
        Exit Function
    End If

    If IsNumeric(cell.Value2) Or (AUTO_COLOR_DATES And IsDate(cell.Value2)) Then
        colorVal = GetAutoColor(0)
        If Not IsEmpty(colorVal) Then
            SetPattern cell, xlPatternGray50, CLng(colorVal)
            ApplyFunctionalMapToCell = colorVal
        End If
        Exit Function
    End If

    cell.Interior.Pattern = xlPatternGray75
    cell.Interior.PatternColor = RGB(255, 255, 255)
    ApplyFunctionalMapToCell = RGB(255, 255, 255)
End Function

Private Sub ApplyRowFormulaDifferences(ByVal rng As Range)
    Dim area As Range
    Dim rowRng As Range
    Dim cell As Range
    Dim counts As Object
    Dim cells As Collection
    Dim formulas As Collection
    Dim f As String
    Dim key As Variant
    Dim maxCount As Long
    Dim maxKey As String
    Dim isTie As Boolean
    Dim i As Long

    For Each area In rng.Areas
        For Each rowRng In area.Rows
            Set counts = CreateObject("Scripting.Dictionary")
            Set cells = New Collection
            Set formulas = New Collection

            For Each cell In rowRng.Cells
                If cell.HasFormula Then
                    f = CStr(cell.FormulaR1C1)
                    cells.Add cell
                    formulas.Add f
                    If counts.Exists(f) Then
                        counts(f) = counts(f) + 1
                    Else
                        counts.Add f, 1
                    End If
                End If
            Next cell

            If counts.Count > 1 Then
                maxCount = 0
                maxKey = ""
                isTie = False
                For Each key In counts.Keys
                    If counts(key) > maxCount Then
                        maxCount = counts(key)
                        maxKey = CStr(key)
                        isTie = False
                    ElseIf counts(key) = maxCount Then
                        isTie = True
                    End If
                Next key

                If maxCount > 1 And Not isTie Then
                    For i = 1 To cells.Count
                        If formulas(i) <> maxKey Then
                            AddRedDots cells(i)
                        End If
                    Next i
                Else
                    Dim lastFormula As String
                    lastFormula = ""
                    For i = 1 To cells.Count
                        If lastFormula <> "" Then
                            If formulas(i) <> lastFormula Then
                                AddRedDots cells(i)
                            End If
                        End If
                        lastFormula = formulas(i)
                    Next i
                End If
            End If
        Next rowRng
    Next area
End Sub

Private Sub AddRedDots(ByVal cell As Range)
    Dim key As String
    Dim baseColor As Variant

    key = CellKey(cell)
    baseColor = Empty
    If Not gFunctionalMapBaseColor Is Nothing Then
        If gFunctionalMapBaseColor.Exists(key) Then
            baseColor = gFunctionalMapBaseColor(key)
        End If
    End If

    If IsEmpty(baseColor) Then
        On Error Resume Next
        baseColor = cell.Interior.Color
        On Error GoTo 0
    End If

    With cell.Interior
        .Color = baseColor
        .Pattern = xlPatternGray75
        .PatternColor = RGB(255, 0, 0)
    End With
End Sub

Private Sub SaveOriginalCellFormat(ByVal cell As Range)
    Dim key As String
    Dim arr(0 To 2) As Variant

    key = CellKey(cell)
    If gFunctionalMapOriginal.Exists(key) Then Exit Sub

    On Error Resume Next
    arr(0) = cell.Interior.Pattern
    arr(1) = cell.Interior.PatternColor
    arr(2) = cell.Interior.Color
    On Error GoTo 0

    gFunctionalMapOriginal.Add key, arr
End Sub

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

Private Function CellHasValue(ByVal cell As Range) As Boolean
    On Error Resume Next
    CellHasValue = (Len(CStr(cell.Text)) > 0)
    On Error GoTo 0
End Function

Private Function HasDataFunction(ByVal formulaText As String) As Boolean
    Dim funcs As Variant
    Dim funcName As Variant
    Dim pattern As String

    funcs = GetDataFunctions()
    If IsEmpty(funcs) Then Exit Function

    For Each funcName In funcs
        pattern = "(^|[^A-Z0-9_])" & EscapeRegex(CStr(funcName)) & "\\s*\\("
        If RegexTest(formulaText, pattern) Then
            HasDataFunction = True
            Exit Function
        End If
    Next funcName
End Function

Private Sub CheckForOffSheetReferences(ByVal cell As Range, ByRef clr As Variant, ByRef bSet As Boolean)
    Dim formulaText As String
    Dim colorVal As Variant
    Dim openPos As Long
    Dim closePos As Long
    Dim namedClr As Variant

    bSet = False
    clr = Empty

    If Not cell.HasFormula Then Exit Sub
    formulaText = CStr(cell.Formula)

    colorVal = GetAutoColor(4)
    If Not IsEmpty(colorVal) Then
        openPos = InStr(1, formulaText, "[", vbTextCompare)
        If openPos > 0 Then
            closePos = InStr(openPos + 1, formulaText, "]", vbTextCompare)
            If closePos > openPos Then
                If InStr(closePos + 1, formulaText, "!", vbTextCompare) > closePos Then
                    clr = colorVal
                    bSet = True
                    Exit Sub
                End If
            End If
        End If

        If RegexTest(formulaText, "\\[[^\\]]+\\][^!]*!") Then
            clr = colorVal
            bSet = True
            Exit Sub
        End If
    End If

    colorVal = GetAutoColor(3)
    If Not IsEmpty(colorVal) Then
        formulaText = RemoveExtraneousSheetName(formulaText, cell.Worksheet.Name)
        If InStr(1, formulaText, "!", vbTextCompare) > 0 Then
            clr = colorVal
            bSet = True
        End If
    End If

    If bSet Then Exit Sub

    If HasOffSheetNamedReference(cell, formulaText, namedClr) Then
        clr = namedClr
        bSet = True
    End If
End Sub

Private Sub CheckForPartialInputs(ByVal cell As Range, ByRef clr As Variant, ByRef bSet As Boolean)
    Dim colorVal As Variant

    bSet = False
    clr = Empty

    colorVal = GetAutoColor(1)
    If IsEmpty(colorVal) Then Exit Sub
    If Not cell.HasFormula Then Exit Sub

    If ContainsPartialInput(cell) Then
        clr = colorVal
        bSet = True
    End If
End Sub

Private Function ContainsPartialInput(ByVal cell As Range) As Boolean
    Dim formulaText As String
    Dim stripped As String
    Dim matches As Object
    Dim matchItem As Object
    Dim numVal As Double
    Dim intVal As Double

    formulaText = CStr(cell.Formula)
    stripped = StripCellRefs(formulaText)
    stripped = RegexReplace(stripped, """[^""]*""", "")

    Set matches = RegexExecute(stripped, "(-?\\d+(?:\\.\\d+)?)")
    If matches Is Nothing Then Exit Function

    For Each matchItem In matches
        On Error Resume Next
        numVal = CDbl(matchItem.Value)
        On Error GoTo 0
        intVal = Fix(numVal)
        If numVal <> intVal Then
            ContainsPartialInput = True
            Exit Function
        End If
        If Not IsAllowedPartialNumber(CLng(intVal)) Then
            ContainsPartialInput = True
            Exit Function
        End If
    Next matchItem
End Function

Private Function IsAllowedPartialNumber(ByVal numVal As Long) As Boolean
    Select Case numVal
        Case 0, 1, 10, 100, 1000, 10000, 100000, 1000000, 10000000, 100000000, 1000000000
            IsAllowedPartialNumber = True
    End Select
End Function

Private Function RemoveExtraneousSheetName(ByVal formulaText As String, ByVal sheetName As String) As String
    Dim safeSheet As String
    safeSheet = Replace(sheetName, "'", "''")

    If InStr(1, formulaText, "'" & safeSheet & "'!", vbTextCompare) > 0 Then
        formulaText = Replace(formulaText, "'" & safeSheet & "'!", "")
    End If
    If InStr(1, formulaText, sheetName & "!", vbTextCompare) > 0 Then
        formulaText = Replace(formulaText, sheetName & "!", "")
    End If

    RemoveExtraneousSheetName = formulaText
End Function

Private Function HasOffSheetNamedReference(ByVal cell As Range, ByVal formulaText As String, ByRef outColor As Variant) As Boolean
    Dim tokens As Collection
    Dim token As Variant
    Dim nm As Name
    Dim nmRange As Range
    Dim foundSheet As Boolean
    Dim foundBook As Boolean
    Dim refersToText As String

    outColor = Empty
    If cell Is Nothing Then Exit Function
    If Len(formulaText) = 0 Then Exit Function

    Set tokens = ExtractNameTokens(formulaText)
    If tokens Is Nothing Then Exit Function

    For Each token In tokens
        Set nm = ResolveNameToken(cell, CStr(token))
        If Not nm Is Nothing Then
            Set nmRange = Nothing
            On Error Resume Next
            Set nmRange = nm.RefersToRange
            On Error GoTo 0

            If Not nmRange Is Nothing Then
                If Not SameWorkbook(cell, nmRange) Then
                    foundBook = True
                ElseIf Not SameWorksheet(cell, nmRange) Then
                    foundSheet = True
                End If
            Else
                refersToText = ""
                On Error Resume Next
                refersToText = CStr(nm.RefersTo)
                On Error GoTo 0
                If Len(refersToText) > 0 Then
                    If InStr(1, refersToText, "[", vbTextCompare) > 0 And InStr(1, refersToText, "]", vbTextCompare) > 0 Then
                        foundBook = True
                    ElseIf InStr(1, refersToText, "!", vbTextCompare) > 0 Then
                        If Not RefersToSameSheet(cell, refersToText) Then
                            foundSheet = True
                        End If
                    End If
                End If
            End If
        End If

        If foundBook Then Exit For
    Next token

    If foundBook Then
        outColor = GetAutoColor(4)
        HasOffSheetNamedReference = Not IsEmpty(outColor)
    ElseIf foundSheet Then
        outColor = GetAutoColor(3)
        HasOffSheetNamedReference = Not IsEmpty(outColor)
    End If
End Function

Private Function ExtractNameTokens(ByVal formulaText As String) As Collection
    Dim reg As Object
    Dim matches As Object
    Dim matchItem As Object
    Dim token As String
    Dim dict As Object
    Dim tokens As Collection

    Set tokens = New Collection
    Set dict = CreateObject("Scripting.Dictionary")

    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = "[A-Za-z_][A-Za-z0-9_.]*"
    reg.Global = True
    reg.IgnoreCase = True

    Set matches = reg.Execute(formulaText)
    For Each matchItem In matches
        If IsIndexInStringLiteral(formulaText, matchItem.FirstIndex) Then GoTo ContinueLoop
        If IsFunctionLikeToken(formulaText, matchItem) Then GoTo ContinueLoop
        token = matchItem.Value
        If IsPlainAddressToken(token) Then GoTo ContinueLoop
        If IsReservedNameToken(token) Then GoTo ContinueLoop
        If Not dict.Exists(token) Then
            dict.Add token, True
            tokens.Add token
        End If
ContinueLoop:
    Next matchItem

    Set ExtractNameTokens = tokens
End Function

Private Function ResolveNameToken(ByVal cell As Range, ByVal token As String) As Name
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nm As Name

    If Len(token) = 0 Then Exit Function
    If cell Is Nothing Then Exit Function

    Set wb = cell.Worksheet.Parent
    On Error Resume Next
    Set nm = wb.Names.Item(token)
    On Error GoTo 0
    If Not nm Is Nothing Then
        Set ResolveNameToken = nm
        Exit Function
    End If

    Set ws = cell.Worksheet
    On Error Resume Next
    Set nm = ws.Names.Item(token)
    On Error GoTo 0
    If Not nm Is Nothing Then
        Set ResolveNameToken = nm
    End If
End Function

Private Function SameWorksheet(ByVal cell As Range, ByVal otherRange As Range) As Boolean
    On Error Resume Next
    SameWorksheet = (Not otherRange Is Nothing) And (Not cell Is Nothing) And _
        (LCase$(otherRange.Worksheet.Name) = LCase$(cell.Worksheet.Name))
    On Error GoTo 0
End Function

Private Function SameWorkbook(ByVal cell As Range, ByVal otherRange As Range) As Boolean
    Dim wb1 As Workbook
    Dim wb2 As Workbook
    On Error Resume Next
    Set wb1 = cell.Worksheet.Parent
    Set wb2 = otherRange.Worksheet.Parent
    If wb1 Is Nothing Or wb2 Is Nothing Then Exit Function
    SameWorkbook = (LCase$(wb1.FullName) = LCase$(wb2.FullName))
    On Error GoTo 0
End Function

Private Function RefersToSameSheet(ByVal cell As Range, ByVal refersToText As String) As Boolean
    Dim cleaned As String
    If cell Is Nothing Then Exit Function
    If Len(refersToText) = 0 Then Exit Function
    cleaned = RemoveExtraneousSheetName(refersToText, cell.Worksheet.Name)
    RefersToSameSheet = (InStr(1, cleaned, "!", vbTextCompare) = 0)
End Function

Private Function IsFunctionLikeToken(ByVal formulaText As String, ByVal matchItem As Object) As Boolean
    Dim idx As Long
    Dim ch As String

    idx = matchItem.FirstIndex + matchItem.Length + 1
    Do While idx <= Len(formulaText)
        ch = Mid$(formulaText, idx, 1)
        If ch = " " Or ch = vbTab Then
            idx = idx + 1
        Else
            If ch = "(" Then IsFunctionLikeToken = True
            Exit Function
        End If
    Loop
End Function

Private Function IsPlainAddressToken(ByVal token As String) As Boolean
    Static reg As Object
    If reg Is Nothing Then
        Set reg = CreateObject("VBScript.RegExp")
        reg.Pattern = "^\$?[A-Za-z]{1,3}\$?\d{1,7}(:\$?[A-Za-z]{1,3}\$?\d{1,7})?$|^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$|^\$?\d{1,7}:\$?\d{1,7}$"
        reg.IgnoreCase = True
    End If
    IsPlainAddressToken = reg.Test(token)
End Function

Private Function IsReservedNameToken(ByVal token As String) As Boolean
    Dim upperToken As String
    upperToken = UCase$(token)
    If upperToken = "TRUE" Or upperToken = "FALSE" Then
        IsReservedNameToken = True
    End If
End Function

Private Function IsIndexInStringLiteral(ByVal text As String, ByVal zeroBasedIndex As Long) As Boolean
    Dim i As Long
    Dim inString As Boolean
    Dim target As Long

    If zeroBasedIndex < 0 Then Exit Function

    target = zeroBasedIndex + 1
    i = 1
    Do While i <= Len(text) And i <= target
        If Mid$(text, i, 1) = """" Then
            If inString Then
                If i < Len(text) And Mid$(text, i + 1, 1) = """" Then
                    i = i + 1
                Else
                    inString = False
                End If
            Else
                inString = True
            End If
        End If
        If i >= target Then Exit Do
        i = i + 1
    Loop

    IsIndexInStringLiteral = inString
End Function

Private Function GetAutoColor(ByVal idx As Long) As Variant
    Select Case idx
        Case 0
            GetAutoColor = RGB(0, 0, 255) ' inputs
        Case 1
            ' partial inputs (not set in default config)
        Case 2
            GetAutoColor = RGB(0, 0, 0) ' formulas
        Case 3
            GetAutoColor = RGB(0, 128, 0) ' sheet links
        Case 4
            GetAutoColor = RGB(128, 0, 128) ' book links
        Case 5
            GetAutoColor = RGB(255, 102, 0) ' hyperlinks
        Case 6
            ' data source (not set in default config)
    End Select
End Function

Private Function GetDataFunctions() As Variant
    GetDataFunctions = Array("RDP.Data", "BDH", "BDP", "BDS", "CIQ", "FDS", "SNL")
End Function

Private Function StripCellRefs(ByVal formulaText As String) As String
    StripCellRefs = RegexReplace(formulaText, "\$?[A-Za-z]{1,3}\$?\d{1,7}", "")
End Function

Private Sub SetPattern(ByVal cell As Range, ByVal pattern As XlPattern, ByVal colorVal As Long)
    With cell.Interior
        .Pattern = pattern
        .PatternColor = colorVal
    End With
End Sub

Private Function HasHyperlink(ByVal cell As Range) As Boolean
    On Error Resume Next
    HasHyperlink = (cell.Hyperlinks.Count > 0)
    On Error GoTo 0
End Function

Private Function CellKey(ByVal cell As Range) As String
    CellKey = cell.Address(True, True, xlA1, True)
End Function

Private Function RangeFromKey(ByVal key As String) As Range
    Dim bangPos As Long
    Dim leftPart As String
    Dim addr As String
    Dim cleaned As String
    Dim wbName As String
    Dim wsName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wbStart As Long
    Dim wbEnd As Long

    On Error Resume Next
    bangPos = InStrRev(key, "!")
    If bangPos = 0 Then Exit Function

    leftPart = Left$(key, bangPos - 1)
    addr = Mid$(key, bangPos + 1)
    cleaned = leftPart
    If Len(cleaned) >= 2 And Left$(cleaned, 1) = "'" And Right$(cleaned, 1) = "'" Then
        cleaned = Mid$(cleaned, 2, Len(cleaned) - 2)
        cleaned = Replace(cleaned, "''", "'")
    End If

    wbStart = InStr(cleaned, "[")
    wbEnd = InStr(cleaned, "]")
    If wbStart > 0 And wbEnd > wbStart Then
        wbName = Mid$(cleaned, wbStart + 1, wbEnd - wbStart - 1)
        wsName = Mid$(cleaned, wbEnd + 1)
    Else
        wsName = cleaned
    End If

    If Len(wbName) > 0 Then
        Set wb = Workbooks(wbName)
    Else
        Set wb = ActiveWorkbook
    End If

    If Not wb Is Nothing Then
        Set ws = wb.Worksheets(wsName)
        If Not ws Is Nothing Then
            Set RangeFromKey = ws.Range(addr)
        End If
    End If
    On Error GoTo 0
End Function

Private Function EscapeRegex(ByVal text As String) As String
    Dim specials As Variant
    Dim i As Long
    specials = Array("\", ".", "+", "*", "?", "^", "$", "(", ")", "[", "]", "{", "}", "|")
    EscapeRegex = text
    For i = LBound(specials) To UBound(specials)
        EscapeRegex = Replace(EscapeRegex, specials(i), "\" & specials(i))
    Next i
End Function

Private Function RegexTest(ByVal text As String, ByVal pattern As String) As Boolean
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = pattern
    reg.IgnoreCase = True
    reg.Global = False
    RegexTest = reg.Test(text)
End Function

Private Function RegexExecute(ByVal text As String, ByVal pattern As String) As Object
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = pattern
    reg.IgnoreCase = True
    reg.Global = True
    Set RegexExecute = reg.Execute(text)
End Function

Private Function RegexReplace(ByVal text As String, ByVal pattern As String, ByVal replacement As String) As String
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = pattern
    reg.IgnoreCase = True
    reg.Global = True
    RegexReplace = reg.Replace(text, replacement)
End Function
