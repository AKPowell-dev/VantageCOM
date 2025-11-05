Attribute VB_Name = "F_Scrolling"
Option Explicit
Option Private Module

Private Enum eRowSearchMode
    modeTop = -1
    modeMiddle = 0
    modeBottom = 1
End Enum

Private Enum eColumnSearchMode
    modeLeft = -1
    modeCenter = 0
    modeRight = 1
End Enum

Private Function ActivateCellInVisibleRange()
    On Error GoTo Catch

    Dim targetRow As Long
    Dim targetColumn As Long
    Dim visibleTop As Long, visibleBottom As Long
    Dim visibleLeft As Long, visibleRight As Long

    targetRow = ActiveCell.Row
    targetColumn = ActiveCell.Column

    With ActiveWindow.VisibleRange
        visibleTop = .item(1).Row
        visibleBottom = PointToRow(.item(.Count).Top - 1, xlNone)
        visibleLeft = .item(1).Column
        visibleRight = PointToColumn(.item(.Count).Left - 1, xlNone)
    End With

    If targetRow < visibleTop Then
        targetRow = visibleTop
    ElseIf targetRow > visibleBottom Then
        targetRow = visibleBottom
    End If

    If targetColumn < visibleLeft Then
        targetColumn = visibleLeft
    ElseIf targetColumn > visibleRight Then
        targetColumn = visibleRight
    End If

    If TypeName(Selection) = "Range" Then
        If ActiveCell.Row <> targetRow Or ActiveCell.Column <> targetColumn Then
            SafeActivateRange Cells(targetRow, targetColumn)
            ActiveWindow.ScrollRow = visibleTop
            ActiveWindow.ScrollColumn = visibleLeft
        End If
    End If
    Exit Function

Catch:
    Call ErrorHandler("ActivateCellInVisibleRange")
End Function

Function ScrollUpHalf(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim topRowVisible As Long
    Dim scrollWidth As Integer
    Dim targetRow As Long

    If gVim.Count1 > 1 Then
        Application.ScreenUpdating = False
        ActiveWindow.LargeScroll Up:=gVim.Count1 \ 2
        Application.ScreenUpdating = True
    End If

    If (gVim.Count1 And 1) = 1 Then
        topRowVisible = ActiveWindow.VisibleRange.Row

        scrollWidth = ActiveWindow.VisibleRange.Rows.Count / 2
        targetRow = topRowVisible - scrollWidth

        If targetRow < 1 Then
            targetRow = 1
        End If

        ActiveWindow.SmallScroll Up:=scrollWidth
    End If

    Call ActivateCellInVisibleRange
    Exit Function

Catch:
    Call ErrorHandler("ScrollUpHalf")
End Function

Function ScrollDownHalf(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim topRowVisible As Long
    Dim scrollWidth As Integer
    Dim targetRow As Long

    If gVim.Count1 > 1 Then
        Application.ScreenUpdating = False
        ActiveWindow.LargeScroll Down:=gVim.Count1 \ 2
        Application.ScreenUpdating = True
    End If

    If (gVim.Count1 And 1) = 1 Then
        topRowVisible = ActiveWindow.VisibleRange.Row

        scrollWidth = ActiveWindow.VisibleRange.Rows.Count / 2
        targetRow = topRowVisible + scrollWidth

        If targetRow > ActiveSheet.Rows.Count Then
            targetRow = ActiveSheet.Rows.Count
        End If

        ActiveWindow.SmallScroll Down:=scrollWidth
    End If

    Call ActivateCellInVisibleRange
    Exit Function

Catch:
    Call ErrorHandler("ScrollDownHalf")
End Function

Function ScrollLeftHalf(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim leftColVisible As Long
    Dim scrollWidth As Integer
    Dim targetCol As Long

    If gVim.Count1 > 1 Then
        Application.ScreenUpdating = False
        ActiveWindow.LargeScroll ToLeft:=gVim.Count1 \ 2
        Application.ScreenUpdating = True
    End If

    If (gVim.Count1 And 1) = 1 Then
        leftColVisible = ActiveWindow.VisibleRange.Column

        scrollWidth = ActiveWindow.VisibleRange.Columns.Count / 2
        targetCol = leftColVisible - scrollWidth

        If targetCol < 1 Then
            targetCol = 1
        End If

        ActiveWindow.SmallScroll ToLeft:=scrollWidth
    End If

    Call ActivateCellInVisibleRange
    Exit Function

Catch:
    Call ErrorHandler("ScrollLeftHalf")
End Function

Function ScrollRightHalf(Optional ByVal g As String) As Boolean
    On Error GoTo Catch

    Dim leftColVisible As Long
    Dim scrollWidth As Integer
    Dim targetCol As Long

    If gVim.Count1 > 1 Then
        Application.ScreenUpdating = False
        ActiveWindow.LargeScroll ToRight:=gVim.Count1 \ 2
        Application.ScreenUpdating = True
    End If

    If (gVim.Count1 And 1) = 1 Then
        leftColVisible = ActiveWindow.VisibleRange.Column

        scrollWidth = ActiveWindow.VisibleRange.Columns.Count / 2
        targetCol = leftColVisible + scrollWidth

        If targetCol > ActiveSheet.Columns.Count Then
            targetCol = ActiveSheet.Columns.Count
        End If

        ActiveWindow.SmallScroll ToRight:=scrollWidth
    End If

    Call ActivateCellInVisibleRange
    Exit Function

Catch:
    Call ErrorHandler("ScrollRightHalf")
End Function


Function ScrollUp(Optional ByVal g As String) As Boolean
    Application.ScreenUpdating = False
    ActiveWindow.LargeScroll Up:=gVim.Count1
    Application.ScreenUpdating = True
    Call ActivateCellInVisibleRange
End Function

Function ScrollDown(Optional ByVal g As String) As Boolean
    Application.ScreenUpdating = False
    ActiveWindow.LargeScroll Down:=gVim.Count1
    Application.ScreenUpdating = True
    Call ActivateCellInVisibleRange
End Function

Function ScrollLeft(Optional ByVal g As String) As Boolean
    Application.ScreenUpdating = False
    ActiveWindow.LargeScroll ToLeft:=gVim.Count1
    Application.ScreenUpdating = True
    Call ActivateCellInVisibleRange
End Function

Function ScrollRight(Optional ByVal g As String) As Boolean
    Application.ScreenUpdating = False
    ActiveWindow.LargeScroll ToRight:=gVim.Count1
    Application.ScreenUpdating = True
    Call ActivateCellInVisibleRange
End Function

Function ScrollUp1Row(Optional ByVal g As String) As Boolean
    ActiveWindow.SmallScroll Up:=gVim.Count1
    Call ActivateCellInVisibleRange
End Function

Function ScrollDown1Row(Optional ByVal g As String) As Boolean
    ActiveWindow.SmallScroll Down:=gVim.Count1
    Call ActivateCellInVisibleRange
End Function

Function ScrollLeft1Column(Optional ByVal g As String) As Boolean
    ActiveWindow.SmallScroll ToLeft:=gVim.Count1
    Call ActivateCellInVisibleRange
End Function

Function ScrollRight1Column(Optional ByVal g As String) As Boolean
    ActiveWindow.SmallScroll ToRight:=gVim.Count1
    Call ActivateCellInVisibleRange
End Function

Private Function PointToRow(ByVal point As Double, ByVal searchMode As eRowSearchMode) As Long
    On Error GoTo Catch

    Dim avg As Double
    Dim pred As Long
    Dim diff As Double
    Dim predTop As Double
    Dim i As Integer
    Dim l As Long
    Dim m As Long
    Dim h As Long
    Dim tmp As Long

    '?????O???P?[?X??????
    If point > Rows(Rows.Count).Top Then
        PointToRow = Rows.Count
        Exit Function
    ElseIf point <= 0 Then
        PointToRow = 1
        Exit Function
    End If

    '?????????????????????????????Z?o
    avg = ActiveWindow.VisibleRange.Height / ActiveWindow.VisibleRange.Rows.Count

    '?????????s??????
    pred = CLng(point / avg) + 1
    If pred > Rows.Count Then
        pred = Rows.Count
    ElseIf pred < 1 Then
        pred = 1
    End If
    predTop = Rows(pred).Top

    '??????????
    diff = point - predTop

    '???????L?????????s????????????????????
    i = 0
    l = pred
    h = pred
    Do Until diff = 0
        tmp = CLng(diff / avg + 0.5) * 2 ^ i
        If tmp = 0 Then
            tmp = Sgn(diff) * 2 ^ i
        End If

        If tmp > Rows.Count Then
            tmp = Rows.Count
        Else
            tmp = pred + tmp
        End If

        If diff < 0 Then
            h = l
            If tmp < 1 Then
                l = 1
            Else
                l = tmp
            End If
        Else
            l = h
            If tmp > Rows.Count Then
                h = Rows.Count
            Else
                h = tmp
            End If
        End If

        If Rows(l).Top <= point And point < Rows(h).Top Then
            Exit Do
        End If

        i = i + 1
    Loop

    '?????T?????s??????
    Do
        m = Round(l + (h - l) / 2 - 0.25)
        If h - l < 2 Then
            Exit Do
        End If

        predTop = Rows(m).Top
        If point < predTop Then
            h = m
        Else
            l = m
        End If
    Loop

    '???[?h????????????????
    Select Case searchMode
        '????????
        Case modeMiddle
            '1?s?????????????????????????l???????????????I??
            If (point - Rows(m).Top) >= Rows(m).Height / 2 Then
                PointToRow = m + 1
            Else
                PointToRow = m
            End If

        '??????
        Case modeTop
            '?s?b?^????????????1?s????
            If point > Rows(m).Top Then
                PointToRow = m + 1
            Else
                PointToRow = m
            End If

        '??????
        Case modeBottom
            'gVim.Config.ScrollOffset ????????????????????1?s????
            If point - gVim.Config.ScrollOffset > Rows(m).Top Then
                PointToRow = m + 1
            Else
                PointToRow = m
            End If

        '???O
        Case Else
            PointToRow = m

    End Select
    Exit Function

Catch:
    Call ErrorHandler("pointToRow")
End Function

Private Function PointToColumn(ByVal point As Double, ByVal searchMode As eColumnSearchMode) As Long
    On Error GoTo Catch

    Dim avg As Double
    Dim pred As Long
    Dim diff As Double
    Dim predLeft As Double
    Dim i As Integer
    Dim l As Long
    Dim m As Long
    Dim h As Long
    Dim tmp As Long

    '?????O???P?[?X??????
    If point > Columns(Columns.Count).Left Then
        PointToColumn = Columns.Count
        Exit Function
    ElseIf point <= 0 Then
        PointToColumn = 1
        Exit Function
    End If

    '???????????????????????????Z?o
    avg = ActiveWindow.VisibleRange.Width / ActiveWindow.VisibleRange.Columns.Count

    '????????????????
    pred = CLng(point / avg) + 1
    If pred > Columns.Count Then
        pred = Columns.Count
    ElseIf pred < 1 Then
        pred = 1
    End If
    predLeft = Columns(pred).Left

    '??????????
    diff = point - predLeft

    '???????L??????????????????????????????
    i = 0
    l = pred
    h = pred
    Do Until diff = 0
        tmp = CLng(diff / avg + 0.5) * 2 ^ i
        If tmp = 0 Then
            tmp = Sgn(diff) * 2 ^ i
        End If

        If tmp > Columns.Count Then
            tmp = Columns.Count
        Else
            tmp = pred + tmp
        End If

        If diff < 0 Then
            h = l
            If tmp < 1 Then
                l = 1
            Else
                l = tmp
            End If
        Else
            l = h
            If tmp > Columns.Count Then
                h = Columns.Count
            Else
                h = tmp
            End If
        End If

        If Columns(l).Left <= point And point < Columns(h).Left Then
            Exit Do
        End If

        i = i + 1
    Loop

    '?????T????????????
    Do
        m = Round(l + (h - l) / 2 - 0.25)
        If h - l < 2 Then
            Exit Do
        End If

        predLeft = Columns(m).Left
        If point < predLeft Then
            h = m
        Else
            l = m
        End If
    Loop

    '???[?h????????????????
    Select Case searchMode
        '????????
        Case modeCenter
            '1???????????????????????????l???????????????I??
            If (point - Columns(m).Left) >= Columns(m).Width / 2 Then
                PointToColumn = m + 1
            Else
                PointToColumn = m
            End If

        '??????, ?E????
        Case modeLeft, modeRight
            '?s?b?^????????????1??????
            If point > Columns(m).Left Then
                PointToColumn = m + 1
            Else
                PointToColumn = m
            End If

        '???O
        Case Else
            PointToColumn = m
    End Select
    Exit Function

Catch:
    Call ErrorHandler("pointToColumn")
End Function

Private Function GetLengthWithZoomConsidered(ByVal Length As Double) As Double
    'Zoom???l????????????????
    Dim rate As Double

    If 90 < ActiveWindow.Zoom And ActiveWindow.Zoom < 110 Then
        rate = 1
    Else
        rate = 103.32 / ActiveWindow.Zoom - 0.05
    End If
    GetLengthWithZoomConsidered = Length * rate
End Function

Private Function GetRealUsableHeight() As Double
    If ActiveWindow.DisplayHeadings Then
        GetRealUsableHeight = ActiveWindow.UsableHeight - ActiveSheet.StandardHeight
    Else
        GetRealUsableHeight = ActiveWindow.UsableHeight
    End If
End Function

Private Function GetRealUsableWidth() As Double
    Dim maxVisibleRow As Long
    Dim headingWidth As Double

    If ActiveWindow.DisplayHeadings Then
        maxVisibleRow = ActiveWindow.VisibleRange.item(ActiveWindow.VisibleRange.Count).Row
        headingWidth = 25

        If maxVisibleRow >= 1000 Then
            headingWidth = headingWidth + 6.75 * (Len(CStr(maxVisibleRow)) - 3)
        End If

        GetRealUsableWidth = ActiveWindow.UsableWidth - headingWidth
    Else
        GetRealUsableWidth = ActiveWindow.UsableWidth
    End If
End Function

Function ScrollCurrentTop(Optional ByVal g As String) As Boolean
    If gVim.Count > 0 Then
        Call MoveToSpecifiedRow(CStr(gVim.Count))
    End If
    ActiveWindow.ScrollRow = PointToRow(ActiveCell.Top - GetLengthWithZoomConsidered(gVim.Config.ScrollOffset), modeTop)
End Function

Function ScrollCurrentBottom(Optional ByVal g As String) As Boolean
    If gVim.Count > 0 Then
        Call MoveToSpecifiedRow(CStr(gVim.Count))
    End If

    Dim uh As Double
    uh = GetRealUsableHeight()

    ActiveWindow.ScrollRow = PointToRow(ActiveCell.Top + ActiveCell.Height - GetLengthWithZoomConsidered(uh - gVim.Config.ScrollOffset), modeBottom)
End Function

Function ScrollCurrentMiddle(Optional ByVal g As String) As Boolean
    If gVim.Count > 0 Then
        Call MoveToSpecifiedRow(CStr(gVim.Count))
    End If

    Dim uh As Double
    uh = GetRealUsableHeight()

    ActiveWindow.ScrollRow = PointToRow(ActiveCell.Top + ActiveCell.Height / 2 - GetLengthWithZoomConsidered(uh) / 2, modeMiddle)
End Function

Function ScrollCurrentLeft(Optional ByVal g As String) As Boolean
    If gVim.Count > 0 Then
        Call MoveToNthColumn
    End If

    ActiveWindow.ScrollColumn = ActiveCell.Column
End Function

Function ScrollCurrentRight(Optional ByVal g As String) As Boolean
    If gVim.Count > 0 Then
        Call MoveToNthColumn
    End If

    Dim uw As Double
    uw = GetRealUsableWidth()

    ActiveWindow.ScrollColumn = PointToColumn(ActiveCell.Left + ActiveCell.Width - GetLengthWithZoomConsidered(uw), modeRight)
End Function

Function ScrollCurrentCenter(Optional ByVal g As String) As Boolean
    If gVim.Count > 0 Then
        Call MoveToNthColumn
    End If

    Dim uw As Double
    uw = GetRealUsableWidth()

    ActiveWindow.ScrollColumn = PointToColumn(ActiveCell.Left + ActiveCell.Width / 2 - GetLengthWithZoomConsidered(uw) / 2, modeCenter)
End Function
