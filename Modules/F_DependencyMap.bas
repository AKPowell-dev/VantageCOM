Attribute VB_Name = "F_DependencyMap"
Option Explicit

Private Const EDGE_DELIM As String = "~~>"

Sub DrawDependencyMap()
    On Error GoTo CleanFail

    Dim wb As Workbook
    Dim edges As Object
    Dim levels As Object
    Dim mapSheet As Worksheet
    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean

    Set wb = ThisWorkbook

    prevCalc = Application.Calculation
    prevScreenUpdating = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set edges = GetSheetDependencies(wb)
    Set levels = BuildSheetLevels(wb, edges)
    Set mapSheet = CreateMapSheet(wb)

    DrawNodesAndEdges mapSheet, levels, edges
    SafeActivateWorksheet mapSheet

CleanExit:
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

CleanFail:
    Call ErrorHandler("DrawDependencyMap")
    Resume CleanExit
End Sub

Private Function GetSheetDependencies(ByVal wb As Workbook) As Object
    Dim edges As Object
    Dim ws As Worksheet
    Dim formulaCells As Range
    Dim cell As Range
    Dim refs As Collection
    Dim refName As Variant
    Dim key As String

    Set edges = CreateObject("Scripting.Dictionary")

    For Each ws In wb.Worksheets
        On Error Resume Next
        Set formulaCells = Nothing
        Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0

        If Not formulaCells Is Nothing Then
            For Each cell In formulaCells.Cells
                Set refs = ExtractSheetRefs(cell.Formula, ws.Name, wb)
                If Not refs Is Nothing Then
                    For Each refName In refs
                        key = refName & EDGE_DELIM & ws.Name
                        If Not edges.Exists(key) Then
                            edges.Add key, True
                        End If
                    Next refName
                End If
            Next cell
        End If
    Next ws

    Set GetSheetDependencies = edges
End Function

Private Function ExtractSheetRefs(ByVal formula As String, _
                                  ByVal currentSheet As String, _
                                  ByVal wb As Workbook) As Collection
    Dim reg As Object
    Dim matches As Object
    Dim match As Object
    Dim candidate As String
    Dim references As New Collection
    Dim sheetName As String

    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True
    reg.IgnoreCase = False
    reg.Pattern = "('([^']+)'|[A-Za-z0-9_]+)!"

    Set matches = reg.Execute(formula)

    For Each match In matches
        candidate = match.Value
        candidate = Left$(candidate, Len(candidate) - 1) ' remove !

        If Left$(candidate, 1) = "'" Then
            sheetName = Mid$(candidate, 2, Len(candidate) - 2)
            sheetName = Replace(sheetName, "''", "'")
        Else
            sheetName = candidate
        End If

        If InStr(sheetName, "]") > 0 Then
            sheetName = Mid$(sheetName, InStrRev(sheetName, "]") + 1)
        End If

        If SheetExists(wb, sheetName) Then
            If StrComp(sheetName, currentSheet, vbTextCompare) <> 0 Then
                If Not ContainsInCollection(references, sheetName) Then
                    references.Add sheetName
                End If
            End If
        End If
    Next match

    If references.Count > 0 Then
        Set ExtractSheetRefs = references
    End If
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    Set ws = Nothing
    Err.Clear
    On Error GoTo 0
End Function

Private Function ContainsInCollection(col As Collection, value As String) As Boolean
    Dim item As Variant
    For Each item In col
        If StrComp(CStr(item), value, vbTextCompare) = 0 Then
            ContainsInCollection = True
            Exit Function
        End If
    Next item
End Function

Private Function CreateMapSheet(ByVal wb As Workbook) As Worksheet
    Dim mapSheet As Worksheet
    Dim shp As Shape

    On Error Resume Next
    Set mapSheet = wb.Worksheets("Dependency Map")
    On Error GoTo 0

    If mapSheet Is Nothing Then
        Set mapSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        mapSheet.Name = "Dependency Map"
    Else
        mapSheet.Cells.Clear
        For Each shp In mapSheet.Shapes
            shp.Delete
        Next shp
    End If

    mapSheet.Cells.VerticalAlignment = xlCenter
    mapSheet.Cells.HorizontalAlignment = xlCenter
    mapSheet.Columns("A:Z").ColumnWidth = 3

    Set CreateMapSheet = mapSheet
End Function

Private Function BuildSheetLevels(ByVal wb As Workbook, ByVal edges As Object) As Object
    Dim levels As Object
    Dim indegree As Object
    Dim adjacency As Object
    Dim incoming As Object
    Dim ws As Worksheet
    Dim key As Variant
    Dim parts() As String
    Dim src As String
    Dim tgt As String
    Dim queue As Collection
    Dim order As Collection
    Dim inOrder As Object
    Dim i As Long
    Dim sheetName As String
    Dim node As String
    Dim lvl As Long
    Dim srcName As Variant

    Set levels = CreateObject("Scripting.Dictionary")
    Set indegree = CreateObject("Scripting.Dictionary")
    Set adjacency = CreateObject("Scripting.Dictionary")
    Set incoming = CreateObject("Scripting.Dictionary")

    For Each ws In wb.Worksheets
        levels(ws.Name) = 0
        indegree(ws.Name) = 0
    Next ws

    For Each key In edges.Keys
        parts = Split(CStr(key), EDGE_DELIM)
        If UBound(parts) = 1 Then
            src = parts(0)
            tgt = parts(1)

            If Not adjacency.Exists(src) Then
                Set adjacency(src) = New Collection
            End If
            adjacency(src).Add tgt

            indegree(tgt) = indegree(tgt) + 1

            If Not incoming.Exists(tgt) Then
                Set incoming(tgt) = New Collection
            End If
            incoming(tgt).Add src
        End If
    Next key

    Set queue = New Collection
    Set order = New Collection
    Set inOrder = CreateObject("Scripting.Dictionary")

    For Each sheetName In indegree.Keys
        If CLng(indegree(sheetName)) = 0 Then
            queue.Add sheetName
        End If
    Next sheetName

    Do While queue.Count > 0
        node = queue(1)
        queue.Remove 1

        If Not inOrder.Exists(node) Then
            order.Add node
            inOrder(node) = True
        End If

        If adjacency.Exists(node) Then
            For Each tgt In adjacency(node)
                indegree(tgt) = indegree(tgt) - 1
                If CLng(indegree(tgt)) = 0 Then
                    queue.Add tgt
                End If
            Next tgt
        End If
    Loop

    For Each sheetName In indegree.Keys
        If Not inOrder.Exists(sheetName) Then
            order.Add sheetName
            inOrder(sheetName) = True
        End If
    Next sheetName

    For i = 1 To order.Count
        node = order(i)
        lvl = 0

        If incoming.Exists(node) Then
            For Each srcName In incoming(node)
                If levels.Exists(srcName) Then
                    If levels(srcName) + 1 > lvl Then
                        lvl = levels(srcName) + 1
                    End If
                End If
            Next srcName
        End If

        levels(node) = lvl
    Next i

    Set BuildSheetLevels = levels
End Function

Private Sub DrawNodesAndEdges(ByVal mapSheet As Worksheet, _
                              ByVal levels As Object, _
                              ByVal edges As Object)
    Dim levelBuckets As Object
    Dim levelKey As Variant
    Dim bucket As Collection
    Dim sheetName As Variant
    Dim nodeShapes As Object
    Dim level As Long
    Dim maxCount As Long
    Dim horizontalSpacing As Double
    Dim verticalSpacing As Double
    Dim nodeWidth As Double
    Dim nodeHeight As Double
    Dim leftMargin As Double
    Dim topMargin As Double
    Dim columnHeight As Double
    Dim totalHeight As Double
    Dim x As Double
    Dim y As Double
    Dim shape As Shape
    Dim appearanceColor As Long
    Dim key As Variant
    Dim parts() As String
    Dim src As String
    Dim tgt As String
    Dim connector As Shape

    Set levelBuckets = CreateObject("Scripting.Dictionary")
    Set nodeShapes = CreateObject("Scripting.Dictionary")

    For Each sheetName In levels.Keys
        level = CLng(levels(sheetName))
        If Not levelBuckets.Exists(level) Then
            Set levelBuckets(level) = New Collection
        End If
        levelBuckets(level).Add CStr(sheetName)
    Next sheetName

    For Each levelKey In levelBuckets.Keys
        If levelBuckets(levelKey).Count > maxCount Then
            maxCount = levelBuckets(levelKey).Count
        End If
    Next levelKey

    nodeWidth = 140
    nodeHeight = 60
    horizontalSpacing = 110
    verticalSpacing = 40
    leftMargin = 60
    topMargin = 60
    columnHeight = Application.WorksheetFunction.Max(maxCount * (nodeHeight + verticalSpacing), nodeHeight + verticalSpacing)
    appearanceColor = RGB(84, 130, 53)

    Dim levelKeys As Variant
    Dim names As Variant
    Dim idx As Long
    Dim nameIndex As Long

    levelKeys = levelBuckets.Keys
    SortNumericArray levelKeys

    For idx = LBound(levelKeys) To UBound(levelKeys)
        levelKey = levelKeys(idx)
        Set bucket = levelBuckets(levelKey)
        names = CollectionToArray(bucket)
        SortTextArray names

        x = leftMargin + CLng(levelKey) * (nodeWidth + horizontalSpacing)
        totalHeight = (UBound(names) - LBound(names) + 1) * nodeHeight
        totalHeight = totalHeight + (UBound(names) - LBound(names)) * verticalSpacing
        y = topMargin + (columnHeight - totalHeight) / 2

        For nameIndex = LBound(names) To UBound(names)
            sheetName = names(nameIndex)
            Set shape = mapSheet.Shapes.AddShape(msoShapeRoundedRectangle, x, y, nodeWidth, nodeHeight)
            shape.TextFrame2.TextRange.Text = CStr(sheetName)
            shape.TextFrame2.TextRange.Font.Size = 12
            shape.TextFrame2.TextRange.Font.Bold = msoTrue
            shape.Fill.ForeColor.RGB = RGB(221, 235, 247)
            shape.Line.ForeColor.RGB = RGB(79, 129, 189)
            shape.Line.Weight = 1.5
            shape.Name = "node_" & CStr(sheetName)
            nodeShapes(CStr(sheetName)) = shape
            y = y + nodeHeight + verticalSpacing
        Next nameIndex
    Next idx

    For Each key In edges.Keys
        parts = Split(CStr(key), EDGE_DELIM)
        If UBound(parts) = 1 Then
            src = parts(0)
            tgt = parts(1)
            If nodeShapes.Exists(src) And nodeShapes.Exists(tgt) Then
                Set connector = mapSheet.Shapes.AddConnector(msoConnectorStraight, 0, 0, 0, 0)
                connector.Line.ForeColor.RGB = appearanceColor
                connector.Line.EndArrowheadStyle = msoArrowheadTriangle
                connector.Line.Weight = 1.5
                connector.ConnectorFormat.BeginConnect nodeShapes(src), 2
                connector.ConnectorFormat.EndConnect nodeShapes(tgt), 1
                connector.RerouteConnections
                connector.ZOrder msoSendToBack
            End If
        End If
    Next key

    AddLegend mapSheet, appearanceColor

    mapSheet.Range("A1").Value = "Dependency Map"
    mapSheet.Range("A1").Font.Bold = True
    mapSheet.Range("A1").Font.Size = 16
End Sub

Private Sub AddLegend(ByVal mapSheet As Worksheet, ByVal lineColor As Long)
    Dim legendBox As Shape
    Dim legendLine As Shape
    Dim topPos As Double
    Dim leftPos As Double

    topPos = 20
    leftPos = 20

    Set legendBox = mapSheet.Shapes.AddShape(msoShapeRectangle, leftPos, topPos, 220, 60)
    legendBox.Fill.ForeColor.RGB = RGB(242, 242, 242)
    legendBox.Line.ForeColor.RGB = RGB(191, 191, 191)
    legendBox.TextFrame2.TextRange.Text = "Legend" & vbCrLf & "Arrow: upstream sheet feeds downstream sheet"
    legendBox.TextFrame2.TextRange.Font.Size = 10
    legendBox.TextFrame2.TextRange.Font.Bold = msoTrue
    legendBox.TextFrame2.VerticalAnchor = msoAnchorMiddle

    Set legendLine = mapSheet.Shapes.AddConnector(msoConnectorStraight, leftPos + 20, topPos + 40, leftPos + 100, topPos + 40)
    legendLine.Line.ForeColor.RGB = lineColor
    legendLine.Line.EndArrowheadStyle = msoArrowheadTriangle
    legendLine.Line.Weight = 1.5
End Sub

Private Sub SortNumericArray(arr As Variant)
    Dim i As Long, j As Long, temp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CLng(arr(i)) > CLng(arr(j)) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

Private Sub SortTextArray(arr As Variant)
    Dim i As Long, j As Long, temp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(CStr(arr(i)), CStr(arr(j)), vbTextCompare) > 0 Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

Private Function CollectionToArray(col As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    If col.Count = 0 Then Exit Function
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count
        arr(i) = col(i)
    Next i
    CollectionToArray = arr
End Function
