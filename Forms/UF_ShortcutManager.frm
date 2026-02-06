VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ShortcutManager
   Caption         =   "Shortcut Manager"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9600
   OleObjectBlob   =   "UF_ShortcutManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_ShortcutManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents txtSearch As MSForms.TextBox
Private WithEvents cboCategory As MSForms.ComboBox
Private WithEvents lstShortcuts As MSForms.ListBox
Private WithEvents btnEdit As MSForms.CommandButton
Private WithEvents btnDisable As MSForms.CommandButton
Private WithEvents btnReset As MSForms.CommandButton
Private WithEvents btnResetAll As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton

Private lblCategory As MSForms.Label
Private lblSearch As MSForms.Label
Private lblHeaderKey As MSForms.Label
Private lblHeaderAction As MSForms.Label
Private lblHeaderMacro As MSForms.Label
Private lblStatus As MSForms.Label
Private lblHint As MSForms.Label

Private mItems As Collection
Private mFiltered As Collection
Private mLoading As Boolean
Private mHasShown As Boolean

Private Const USE_FIXED_SIZE As Boolean = True
Private Const FIXED_WIDTH_PX As Long = 760
Private Const FIXED_HEIGHT_PX As Long = 1120
Private Const SEARCH_WIDTH_PX As Long = 320
Private Const SEARCH_MIN_WIDTH_PX As Long = 220
Private Const LIST_SCROLLBAR_PTS As Long = 12

#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
#Else
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const DEFAULT_DPI As Double = 96#

Public Property Get Items() As Collection
    Set Items = mItems
End Property

Public Sub ShowManager()
    If Me.Visible Then Exit Sub

    If lstShortcuts Is Nothing Then BuildControls

    EnsureMinimumSize
    mHasShown = True

    Me.StartUpPosition = 1
    Me.Show vbModeless

    EnableResize
    LayoutControls
    LoadItems
    On Error Resume Next
    lstShortcuts.SetFocus
    On Error GoTo 0
End Sub

Private Sub UserForm_Initialize()
    BuildControls
End Sub

Private Sub UserForm_Activate()
    EnableResize
    LayoutControls
End Sub

Private Sub UserForm_Resize()
    LayoutControls
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If lstShortcuts Is Nothing Then Exit Sub
    If X < lstShortcuts.Left Or X > (lstShortcuts.Left + lstShortcuts.Width) Then Exit Sub
    If Y < lstShortcuts.Top Or Y > (lstShortcuts.Top + lstShortcuts.Height) Then Exit Sub
    If Not Me.ActiveControl Is Nothing Then
        If Me.ActiveControl Is txtSearch Or Me.ActiveControl Is cboCategory Then Exit Sub
    End If
    On Error Resume Next
    lstShortcuts.SetFocus
    On Error GoTo 0
End Sub

Private Sub BuildControls()
    Me.Font.Name = "Segoe UI"
    Me.Font.Size = 9
    Me.BackColor = RGB(248, 248, 248)
    Set lblCategory = Me.Controls.Add("Forms.Label.1", "lblCategory")
    lblCategory.Caption = "Category"
    lblCategory.Font.Size = 9

    Set cboCategory = Me.Controls.Add("Forms.ComboBox.1", "cboCategory")
    cboCategory.Style = fmStyleDropDownList

    Set lblSearch = Me.Controls.Add("Forms.Label.1", "lblSearch")
    lblSearch.Caption = "Search"
    lblSearch.Font.Size = 9

    Set txtSearch = Me.Controls.Add("Forms.TextBox.1", "txtSearch")
    txtSearch.BackColor = vbWhite

    Set lblHint = Me.Controls.Add("Forms.Label.1", "lblHint")
    lblHint.Caption = "Tip: double-click to change a shortcut"
    lblHint.Font.Size = 8
    lblHint.ForeColor = RGB(100, 100, 100)

    Set lblHeaderKey = Me.Controls.Add("Forms.Label.1", "lblHeaderKey")
    lblHeaderKey.Caption = "Key"
    lblHeaderKey.Font.Bold = True
    lblHeaderKey.Font.Size = 9

    Set lblHeaderAction = Me.Controls.Add("Forms.Label.1", "lblHeaderAction")
    lblHeaderAction.Caption = "Action"
    lblHeaderAction.Font.Bold = True
    lblHeaderAction.Font.Size = 9

    Set lblHeaderMacro = Me.Controls.Add("Forms.Label.1", "lblHeaderMacro")
    lblHeaderMacro.Caption = "Macro"
    lblHeaderMacro.Font.Bold = True
    lblHeaderMacro.Font.Size = 9

    Set lstShortcuts = Me.Controls.Add("Forms.ListBox.1", "lstShortcuts")
    lstShortcuts.ColumnCount = 3
    lstShortcuts.MultiSelect = fmMultiSelectSingle
    lstShortcuts.IntegralHeight = False
    lstShortcuts.BackColor = vbWhite
    lstShortcuts.Font.Name = "Segoe UI"
    lstShortcuts.Font.Size = 9

    Set btnEdit = Me.Controls.Add("Forms.CommandButton.1", "btnEdit")
    btnEdit.Caption = "Edit Key"

    Set btnDisable = Me.Controls.Add("Forms.CommandButton.1", "btnDisable")
    btnDisable.Caption = "Disable"

    Set btnReset = Me.Controls.Add("Forms.CommandButton.1", "btnReset")
    btnReset.Caption = "Reset"

    Set btnResetAll = Me.Controls.Add("Forms.CommandButton.1", "btnResetAll")
    btnResetAll.Caption = "Reset All"

    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    btnClose.Caption = "Close"

    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus")
    lblStatus.Caption = ""

    LayoutControls
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleEscape KeyCode
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleEscape KeyCode
End Sub

Private Sub cboCategory_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleEscape KeyCode
End Sub

Private Sub lstShortcuts_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleEscape KeyCode
End Sub

Private Sub btnEdit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleEscape KeyCode
End Sub

Private Sub btnDisable_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleEscape KeyCode
End Sub

Private Sub btnReset_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleEscape KeyCode
End Sub

Private Sub btnResetAll_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleEscape KeyCode
End Sub

Private Sub btnClose_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    HandleEscape KeyCode
End Sub

Private Sub HandleEscape(ByRef KeyCode As MSForms.ReturnInteger)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        Unload Me
    End If
End Sub

Private Sub LayoutControls()
    Dim leftPad As Double
    Dim topRow As Double
    Dim colGap As Double
    Dim boxHeight As Double
    Dim innerWidth As Double
    Dim innerHeight As Double
    Dim headerTop As Double
    Dim listTop As Double
    Dim footerTop As Double
    Dim hintTop As Double
    Dim searchOnSecondRow As Boolean
    Dim searchWidthTwips As Double
    Dim minSearchWidthTwips As Double
    Dim usableListWidth As Double

    leftPad = 10
    topRow = 8
    colGap = 6
    boxHeight = 18
    innerWidth = GetInnerWidth()
    innerHeight = GetInnerHeight()

    lblCategory.Left = leftPad
    lblCategory.Top = topRow + 3
    lblCategory.Width = 60
    lblCategory.Height = 12

    cboCategory.Left = lblCategory.Left + lblCategory.Width + colGap
    cboCategory.Top = topRow
    cboCategory.Width = 130
    cboCategory.Height = boxHeight

    lblSearch.Left = cboCategory.Left + cboCategory.Width + 2 * colGap
    lblSearch.Top = topRow + 3
    lblSearch.Width = 45
    lblSearch.Height = 12

    txtSearch.Left = lblSearch.Left + lblSearch.Width + colGap
    txtSearch.Top = topRow
    searchWidthTwips = PixelsToTwipsX(SEARCH_WIDTH_PX)
    minSearchWidthTwips = PixelsToTwipsX(SEARCH_MIN_WIDTH_PX)
    searchOnSecondRow = (txtSearch.Left + minSearchWidthTwips > innerWidth - leftPad)
    If searchOnSecondRow Then
        lblSearch.Left = leftPad
        lblSearch.Top = topRow + boxHeight + 6
        txtSearch.Left = lblSearch.Left + lblSearch.Width + colGap
        txtSearch.Top = lblSearch.Top - 3
        txtSearch.Width = ClampSize(innerWidth - txtSearch.Left - leftPad, minSearchWidthTwips)
        hintTop = txtSearch.Top + boxHeight + 4
    Else
        txtSearch.Width = ClampSize(searchWidthTwips, minSearchWidthTwips)
        hintTop = topRow + 22
    End If
    txtSearch.Height = boxHeight

    lblHint.Left = leftPad
    lblHint.Top = hintTop
    lblHint.Width = ClampSize(innerWidth - 2 * leftPad, PixelsToTwipsX(180))
    lblHint.Height = 11

    headerTop = lblHint.Top + 14
    listTop = headerTop + 14
    footerTop = innerHeight - 32

    lstShortcuts.Left = leftPad
    lstShortcuts.Top = listTop
    lstShortcuts.Width = ClampSize(innerWidth - 2 * leftPad, PixelsToTwipsX(320))
    lstShortcuts.Height = ClampSize(footerTop - listTop - 6, PixelsToTwipsY(200))
    usableListWidth = lstShortcuts.Width - PointsToTwips(LIST_SCROLLBAR_PTS)
    If usableListWidth < 0 Then usableListWidth = lstShortcuts.Width
    UpdateColumnWidths usableListWidth
    UpdateHeaderLayout headerTop, usableListWidth

    Dim buttonWidth As Double
    Dim buttonGap As Double
    buttonWidth = 60
    buttonGap = 5

    btnClose.Top = footerTop
    btnClose.Left = innerWidth - leftPad - buttonWidth

    btnResetAll.Top = footerTop
    btnResetAll.Left = btnClose.Left - buttonGap - buttonWidth

    btnReset.Top = footerTop
    btnReset.Left = btnResetAll.Left - buttonGap - buttonWidth

    btnDisable.Top = footerTop
    btnDisable.Left = btnReset.Left - buttonGap - buttonWidth

    btnEdit.Top = footerTop
    btnEdit.Left = btnDisable.Left - buttonGap - buttonWidth

    lblStatus.Top = footerTop + 4
    lblStatus.Left = leftPad
    lblStatus.Width = ClampSize(btnEdit.Left - leftPad - buttonGap, PixelsToTwipsX(90))
    lblStatus.Height = 12
End Sub

Private Sub UpdateHeaderLayout(ByVal headerTop As Double, ByVal totalWidthTwips As Double)
    Dim colKeyPts As Double
    Dim colActionPts As Double
    Dim colMacroPts As Double
    Dim colKeyTwips As Double
    Dim colActionTwips As Double
    Dim colMacroTwips As Double

    ComputeColumnWidths totalWidthTwips, colKeyPts, colActionPts, colMacroPts
    colKeyTwips = PointsToTwips(colKeyPts)
    colActionTwips = PointsToTwips(colActionPts)
    colMacroTwips = PointsToTwips(colMacroPts)

    lblHeaderKey.Left = lstShortcuts.Left
    lblHeaderKey.Top = headerTop
    lblHeaderKey.Width = colKeyTwips
    lblHeaderKey.Height = 12

    lblHeaderAction.Left = lblHeaderKey.Left + colKeyTwips
    lblHeaderAction.Top = headerTop
    lblHeaderAction.Width = colActionTwips
    lblHeaderAction.Height = 12

    lblHeaderMacro.Left = lblHeaderAction.Left + colActionTwips
    lblHeaderMacro.Top = headerTop
    lblHeaderMacro.Width = colMacroTwips
    lblHeaderMacro.Height = 12
    lblHeaderKey.ZOrder 0
    lblHeaderAction.ZOrder 0
    lblHeaderMacro.ZOrder 0
End Sub

Private Sub UpdateColumnWidths(ByVal totalWidthTwips As Double)
    Dim colKeyPts As Double
    Dim colActionPts As Double
    Dim colMacroPts As Double

    ComputeColumnWidths totalWidthTwips, colKeyPts, colActionPts, colMacroPts
    lstShortcuts.ColumnWidths = CStr(colKeyPts) & " pt;" & CStr(colActionPts) & " pt;" & CStr(colMacroPts) & " pt"
End Sub

Private Sub ComputeColumnWidths(ByVal totalWidthTwips As Double, ByRef colKeyPts As Double, ByRef colActionPts As Double, ByRef colMacroPts As Double)
    Dim totalPts As Double
    totalPts = TwipsToPoints(totalWidthTwips)

    colKeyPts = 55
    colMacroPts = 120
    colActionPts = totalPts - (colKeyPts + colMacroPts)

    If colActionPts < 160 Then
        colActionPts = 160
        colMacroPts = totalPts - (colKeyPts + colActionPts)
        If colMacroPts < 100 Then colMacroPts = 100
    End If
End Sub

Private Sub LoadItems()
    mLoading = True
    Set mItems = ShortcutManager_BuildItems()
    PopulateCategories
    mLoading = False

    If cboCategory.ListCount > 0 Then cboCategory.ListIndex = 0
    txtSearch.Text = ""

    RefreshList
End Sub

Private Sub PopulateCategories()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim item As cls_ShortcutItem
    For Each item In mItems
        If Not dict.Exists(item.Category) Then
            dict.Add item.Category, True
        End If
    Next item

    cboCategory.Clear
    cboCategory.AddItem "All"

    Dim key As Variant
    For Each key In dict.Keys
        cboCategory.AddItem CStr(key)
    Next key
End Sub

Private Sub RefreshList()
    If mLoading Then Exit Sub

    Set mFiltered = New Collection
    lstShortcuts.Clear

    Dim categoryFilter As String
    categoryFilter = cboCategory.Value

    Dim searchText As String
    searchText = LCase$(Trim$(txtSearch.Text))

    Dim item As cls_ShortcutItem
    For Each item In mItems
        If categoryFilter <> "All" Then
            If StrComp(item.Category, categoryFilter, vbTextCompare) <> 0 Then GoTo ContinueRow
        End If

        If Len(searchText) > 0 Then
            Dim hay As String
            hay = LCase$(item.Action & " " & item.CurrentKey & " " & item.DefaultKey & " " & item.Category)
            If InStr(hay, searchText) = 0 Then GoTo ContinueRow
        End If

        mFiltered.Add item
        lstShortcuts.AddItem item.CurrentKey
        Dim actionLabel As String
        If IsCmdItem(item) Then
            actionLabel = CmdActionDescription(item.Action, item.DefaultKey)
        Else
            actionLabel = ActionDescription(item.Action)
        End If
        lstShortcuts.List(lstShortcuts.ListCount - 1, 1) = actionLabel
        lstShortcuts.List(lstShortcuts.ListCount - 1, 2) = item.Action

ContinueRow:
    Next item

    UpdateButtons
    UpdateStatus
End Sub

Private Function SelectedItem() As cls_ShortcutItem
    If lstShortcuts.ListIndex < 0 Then Exit Function
    If mFiltered Is Nothing Then Exit Function
    Set SelectedItem = mFiltered(lstShortcuts.ListIndex + 1)
End Function

Private Sub UpdateButtons()
    Dim hasSelection As Boolean
    hasSelection = (lstShortcuts.ListIndex >= 0)

    If hasSelection Then
        Dim item As cls_ShortcutItem
        Set item = SelectedItem
        If IsCmdItem(item) Then
            btnEdit.Enabled = False
            btnDisable.Enabled = False
            btnReset.Enabled = False
        Else
            btnEdit.Enabled = True
            btnDisable.Enabled = True
            btnReset.Enabled = True
        End If
    Else
        btnEdit.Enabled = False
        btnDisable.Enabled = False
        btnReset.Enabled = False
    End If
End Sub

Private Sub UpdateStatus()
    Dim item As cls_ShortcutItem
    Set item = SelectedItem
    If item Is Nothing Then
        lblStatus.Caption = CStr(lstShortcuts.ListCount) & " shortcuts"
    Else
        Dim curKey As String
        curKey = item.CurrentKey
        If Len(Trim$(curKey)) = 0 Then curKey = "(disabled)"
        lblStatus.Caption = "Default: " & item.DefaultKey & "  |  Current: " & curKey
    End If
End Sub

Private Sub txtSearch_Change()
    RefreshList
End Sub

Private Sub cboCategory_Change()
    RefreshList
End Sub

Private Sub lstShortcuts_Click()
    UpdateButtons
    UpdateStatus
End Sub

Private Sub lstShortcuts_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not Me.ActiveControl Is Nothing Then
        If Me.ActiveControl Is txtSearch Or Me.ActiveControl Is cboCategory Then Exit Sub
    End If
    On Error Resume Next
    lstShortcuts.SetFocus
    On Error GoTo 0
End Sub

Private Sub lstShortcuts_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstShortcuts.ListIndex < 0 Then Exit Sub
    btnEdit_Click
End Sub

Private Sub btnEdit_Click()
    Dim item As cls_ShortcutItem
    Set item = SelectedItem
    If item Is Nothing Then Exit Sub
    If IsCmdItem(item) Then Exit Sub

    Dim newKey As String
    newKey = InputBox("Enter new key (Vim style), or blank to disable:", "Shortcut Manager", item.CurrentKey)
    If StrPtr(newKey) = 0 Then Exit Sub
    newKey = Trim$(newKey)

    If Len(newKey) = 0 Then
        ShortcutManager_Disable item
        RefreshList
        Exit Sub
    End If

    Dim errMsg As String
    If Not ShortcutManager_ValidateKey(newKey, errMsg) Then
        MsgBox errMsg, vbExclamation, "Shortcut Manager"
        Exit Sub
    End If

    Dim conflict As cls_ShortcutItem
    Set conflict = ShortcutManager_FindByKey(mItems, newKey, item.Action)
    If Not conflict Is Nothing Then
        If MsgBox("Key is already assigned to: " & conflict.Action & vbCrLf & _
                  "Overwrite and disable that shortcut?", vbYesNo + vbExclamation, "Shortcut Manager") = vbNo Then
            Exit Sub
        End If
        ShortcutManager_Disable conflict
    End If

    ShortcutManager_ApplyChange item, newKey
    RefreshList
End Sub

Private Sub btnDisable_Click()
    Dim item As cls_ShortcutItem
    Set item = SelectedItem
    If item Is Nothing Then Exit Sub
    If IsCmdItem(item) Then Exit Sub

    ShortcutManager_Disable item
    RefreshList
End Sub

Private Sub btnReset_Click()
    Dim item As cls_ShortcutItem
    Set item = SelectedItem
    If item Is Nothing Then Exit Sub
    If IsCmdItem(item) Then Exit Sub

    ShortcutManager_ResetItem item
    RefreshList
End Sub

Private Sub btnResetAll_Click()
    If MsgBox("Reset all shortcuts to defaults?", vbYesNo + vbQuestion, "Shortcut Manager") = vbNo Then Exit Sub
    ShortcutManager_ResetAll mItems
    RefreshList
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Function ActionDescription(ByVal macroName As String) As String
    Dim desc As String
    desc = DescribeAction(macroName)
    If Len(desc) > 0 Then
        ActionDescription = desc
    Else
        ActionDescription = FriendlyAction(macroName)
    End If
End Function

Private Function CmdActionDescription(ByVal actionText As String, ByVal keyText As String) As String
    Dim baseName As String
    Dim argText As String
    SplitActionText actionText, baseName, argText

    Dim lower As String
    lower = LCase$(baseName)

    Select Case lower
    Case "shortcutmanager"
        CmdActionDescription = "Open shortcut manager"
    Case "showcommandinfo"
        CmdActionDescription = "Show command list"
    Case "namescrubber"
        CmdActionDescription = "Open name scrubber"
    Case "reloadvim"
        CmdActionDescription = "Reload add-in"
    Case "showversion"
        CmdActionDescription = "Show version"
    Case "toggledebugmode"
        CmdActionDescription = "Toggle debug mode"
    Case "toggleplainkeymappings"
        CmdActionDescription = "Toggle plain keys"
    Case "overrideshortcuts"
        CmdActionDescription = "Override shortcuts"
    Case "togglemacrosafemode"
        CmdActionDescription = "Toggle macro safe mode"
    Case "resetworkbookview"
        CmdActionDescription = "Reset workbook view"
    Case "resizeSelectiontowidthprompt"
        CmdActionDescription = "Resize selection width"
    Case "openactivebookdir"
        CmdActionDescription = "Open workbook folder"
    Case "yankactivebookpath"
        CmdActionDescription = "Copy workbook path"
    Case "removeduplicates"
        CmdActionDescription = "Remove duplicates"
    Case "reverseselectionorder"
        CmdActionDescription = "Reverse selection order"
    Case "printpreviewofactivesheet"
        CmdActionDescription = "Print preview sheet"
    Case "printpreviewandprint"
        CmdActionDescription = "Print sheet"
    Case "openworkbook"
        CmdActionDescription = "Open workbook"
    Case "reopenactiveworkbook"
        CmdActionDescription = "Reopen workbook"
    Case "saveworkbook"
        CmdActionDescription = "Save workbook"
    Case "closeasksaving"
        CmdActionDescription = "Close (ask save)"
    Case "closewithoutsaving"
        CmdActionDescription = "Close (no save)"
    Case "closewithsaving"
        CmdActionDescription = "Close + save"
    Case "saveasnewworkbook"
        CmdActionDescription = "Save as new"
    Case "activateworkbook"
        CmdActionDescription = "Switch workbook"
    Case "nextworkbook"
        CmdActionDescription = "Next workbook"
    Case "previousworkbook"
        CmdActionDescription = "Previous workbook"
    Case "sort"
        CmdActionDescription = "Sort selection"
    Case "pastecondensed"
        CmdActionDescription = "Paste condensed"
    Case "cmdinsertnumbers"
        CmdActionDescription = "Insert numbers"
    Case "econs_output_ppt_v2"
        CmdActionDescription = "Export to PPT"
    Case "formatoverviewgraph"
        CmdActionDescription = "Format overview graph"
    Case "launchresearchlink"
        If Len(argText) > 0 Then
            CmdActionDescription = "Open research: " & argText
        Else
            CmdActionDescription = "Open research"
        End If
    Case "searchhelp"
        CmdActionDescription = "Search help"
    Case "clearjumps"
        CmdActionDescription = "Clear jumps"
    Case Else
        CmdActionDescription = FriendlyAction(actionText)
    End Select
End Function

Private Sub SplitActionText(ByVal actionText As String, ByRef baseName As String, ByRef argText As String)
    Dim trimmed As String
    trimmed = Trim$(actionText)
    If Len(trimmed) = 0 Then Exit Sub

    Dim spacePos As Long
    spacePos = InStr(trimmed, " ")
    If spacePos > 0 Then
        baseName = Left$(trimmed, spacePos - 1)
        argText = Trim$(Mid$(trimmed, spacePos + 1))
        If Left$(argText, 1) = """" And Right$(argText, 1) = """" Then
            argText = Mid$(argText, 2, Len(argText) - 2)
        End If
    Else
        baseName = trimmed
        argText = ""
    End If
End Sub

Private Function DescribeAction(ByVal macroName As String) As String
    Dim nameLower As String
    nameLower = LCase$(macroName)

    If InStr(nameLower, "move") > 0 Then
        If InStr(nameLower, "left") > 0 Then DescribeAction = "Move left": Exit Function
        If InStr(nameLower, "right") > 0 Then DescribeAction = "Move right": Exit Function
        If InStr(nameLower, "up") > 0 Then DescribeAction = "Move up": Exit Function
        If InStr(nameLower, "down") > 0 Then DescribeAction = "Move down": Exit Function
        DescribeAction = "Move": Exit Function
    End If

    If InStr(nameLower, "select") > 0 Then
        If InStr(nameLower, "chart") > 0 Then DescribeAction = "Select chart": Exit Function
        If InStr(nameLower, "row") > 0 Then DescribeAction = "Select row": Exit Function
        If InStr(nameLower, "column") > 0 Then DescribeAction = "Select column": Exit Function
        DescribeAction = "Select": Exit Function
    End If

    If InStr(nameLower, "insert") > 0 Then
        If InStr(nameLower, "row") > 0 Then DescribeAction = "Insert rows": Exit Function
        If InStr(nameLower, "column") > 0 Then DescribeAction = "Insert columns": Exit Function
        If InStr(nameLower, "cell") > 0 Then DescribeAction = "Insert cells": Exit Function
        DescribeAction = "Insert": Exit Function
    End If

    If InStr(nameLower, "delete") > 0 Then
        If InStr(nameLower, "row") > 0 Then DescribeAction = "Delete rows": Exit Function
        If InStr(nameLower, "column") > 0 Then DescribeAction = "Delete columns": Exit Function
        If InStr(nameLower, "cell") > 0 Then DescribeAction = "Delete cells": Exit Function
        DescribeAction = "Delete": Exit Function
    End If

    If InStr(nameLower, "hide") > 0 Then
        If InStr(nameLower, "row") > 0 Then DescribeAction = "Hide rows": Exit Function
        If InStr(nameLower, "column") > 0 Then DescribeAction = "Hide columns": Exit Function
        DescribeAction = "Hide": Exit Function
    End If

    If InStr(nameLower, "unhide") > 0 Then
        If InStr(nameLower, "row") > 0 Then DescribeAction = "Unhide rows": Exit Function
        If InStr(nameLower, "column") > 0 Then DescribeAction = "Unhide columns": Exit Function
        DescribeAction = "Unhide": Exit Function
    End If

    If InStr(nameLower, "copy") > 0 Then
        DescribeAction = "Copy": Exit Function
    End If
    If InStr(nameLower, "paste") > 0 Then
        DescribeAction = "Paste": Exit Function
    End If

    If InStr(nameLower, "format") > 0 Then
        If InStr(nameLower, "chart") > 0 Then DescribeAction = "Format chart": Exit Function
        DescribeAction = "Format": Exit Function
    End If

    If InStr(nameLower, "border") > 0 Then
        DescribeAction = "Border cycle": Exit Function
    End If

    If InStr(nameLower, "color") > 0 Then
        DescribeAction = "Color cycle": Exit Function
    End If
End Function

Private Function FriendlyAction(ByVal macroName As String) As String
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim prev As String

    s = Replace(macroName, "_", " ")
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If i > 1 Then prev = Mid$(s, i - 1, 1) Else prev = ""

        If ch Like "[A-Z]" And prev Like "[a-z]" Then
            FriendlyAction = FriendlyAction & " " & ch
        Else
            FriendlyAction = FriendlyAction & ch
        End If
    Next i

    FriendlyAction = Trim$(FriendlyAction)
End Function

Private Function IsCmdItem(ByVal item As cls_ShortcutItem) As Boolean
    If item Is Nothing Then Exit Function
    IsCmdItem = (Left$(item.DefaultKey, 1) = ":")
End Function

Private Sub EnableResize()
    On Error Resume Next
#If VBA7 Then
    Dim hWnd As LongPtr
    hWnd = FindWindowA("ThunderDFrame", Me.Caption)
    If hWnd = 0 Then hWnd = FindWindowA("ThunderXFrame", Me.Caption)
    If hWnd <> 0 Then
        Dim style As LongPtr
        style = GetWindowLongPtr(hWnd, GWL_STYLE)
        style = style Or WS_THICKFRAME Or WS_MAXIMIZEBOX
        SetWindowLongPtr hWnd, GWL_STYLE, style
        SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
    End If
#Else
    Dim hWnd32 As Long
    hWnd32 = FindWindowA("ThunderDFrame", Me.Caption)
    If hWnd32 = 0 Then hWnd32 = FindWindowA("ThunderXFrame", Me.Caption)
    If hWnd32 <> 0 Then
        Dim style32 As Long
        style32 = GetWindowLong(hWnd32, GWL_STYLE)
        style32 = style32 Or WS_THICKFRAME Or WS_MAXIMIZEBOX
        SetWindowLong hWnd32, GWL_STYLE, style32
        SetWindowPos hWnd32, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
    End If
#End If
    On Error GoTo 0
End Sub

Private Sub EnsureMinimumSize()
    If Not USE_FIXED_SIZE Then Exit Sub
    Dim targetW As Double
    Dim targetH As Double
    targetW = PixelsToTwipsX(FIXED_WIDTH_PX)
    targetH = PixelsToTwipsY(FIXED_HEIGHT_PX)
    If Me.Width < targetW Then Me.Width = targetW
    If Me.Height < targetH Then Me.Height = targetH
End Sub

Private Sub ApplyFixedSize()
    If Not USE_FIXED_SIZE Then Exit Sub
    If FIXED_WIDTH_PX <= 0 Or FIXED_HEIGHT_PX <= 0 Then Exit Sub

    Me.Width = PixelsToTwipsX(FIXED_WIDTH_PX)
    Me.Height = PixelsToTwipsY(FIXED_HEIGHT_PX)
End Sub

Private Function GetInnerWidth() As Double
    Dim width As Double
    width = Me.InsideWidth
    If width <= 0 Then width = Me.Width
    GetInnerWidth = width
End Function

Private Function GetInnerHeight() As Double
    Dim height As Double
    height = Me.InsideHeight
    If height <= 0 Then height = Me.Height
    GetInnerHeight = height
End Function

Private Function ClampSize(ByVal value As Double, ByVal minValue As Double) As Double
    If value < minValue Then
        ClampSize = minValue
    Else
        ClampSize = value
    End If
End Function

Private Function TwipsToPoints(ByVal value As Double) As Double
    TwipsToPoints = value / 20#
End Function

Private Function PointsToTwips(ByVal value As Double) As Double
    PointsToTwips = value * 20#
End Function

Private Function PixelsToTwipsX(ByVal pixels As Long) As Double
    Dim pxPerPt As Double
    On Error Resume Next
    If Not Application.ActiveWindow Is Nothing Then
        pxPerPt = Application.ActiveWindow.PointsToScreenPixelsX(1)
    ElseIf Application.Windows.Count > 0 Then
        pxPerPt = Application.Windows(1).PointsToScreenPixelsX(1)
    End If
    If Err.Number <> 0 Or pxPerPt <= 0 Then
        Err.Clear
        pxPerPt = DEFAULT_DPI / 72#
    End If
    On Error GoTo 0
    PixelsToTwipsX = (pixels / pxPerPt) * 20#
End Function

Private Function PixelsToTwipsY(ByVal pixels As Long) As Double
    Dim pxPerPt As Double
    On Error Resume Next
    If Not Application.ActiveWindow Is Nothing Then
        pxPerPt = Application.ActiveWindow.PointsToScreenPixelsY(1)
    ElseIf Application.Windows.Count > 0 Then
        pxPerPt = Application.Windows(1).PointsToScreenPixelsY(1)
    End If
    If Err.Number <> 0 Or pxPerPt <= 0 Then
        Err.Clear
        pxPerPt = DEFAULT_DPI / 72#
    End If
    On Error GoTo 0
    PixelsToTwipsY = (pixels / pxPerPt) * 20#
End Function
