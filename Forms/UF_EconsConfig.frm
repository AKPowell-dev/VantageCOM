VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_EconsConfig
   Caption         =   "Econs - PowerPoint Output"
   ClientHeight    =   9660.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9120.001
   OleObjectBlob   =   "UF_EconsConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_EconsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Unified configuration dialog for the econs PowerPoint export.
' The dialog edits a cls_EconsConfig instance in place (the single source of
' truth) and reports whether the user confirmed via the Confirmed flag.
'
' Cases are chosen from a checkbox list populated with the case cell's
' data-validation dropdown options, so only valid cases can be selected. The
' Advanced button opens UF_EconsAdvanced for per-case extra outputs; the History
' button opens UF_EconsHistory to reload a previous run.
'
' Layout is computed at run time by Relayout from the form's current size, so
' the dialog can be resized (via C_FormResize) and all controls reflow/anchor
' accordingly. Cosmetic properties (fonts/colours) are baked into the .frx.

' --- Layout metrics (points) ---
Private Const MARGIN As Single = 12
Private Const GAP As Single = 6
Private Const HEADER_H As Single = 24
Private Const HEADER_TEXT_H As Single = 18
Private Const LABEL_H As Single = 16
Private Const SECTION_H As Single = 16
Private Const FIELD_H As Single = 18
Private Const ROW_H As Single = 24
Private Const REPLACE_H As Single = 22
Private Const BTN_W As Single = 84
Private Const BTN_H As Single = 26
Private Const ADV_W As Single = 96
Private Const ROWLABEL_W As Single = 22
Private Const REFRESH_W As Single = 80
Private Const HIST_W As Single = 84
Private Const RESET_W As Single = 70

Private mConfig As cls_EconsConfig
Private mInLayout As Boolean

Public Confirmed As Boolean

Public Property Get Config() As cls_EconsConfig
    Set Config = mConfig
End Property

Public Property Set Config(ByVal value As cls_EconsConfig)
    Set mConfig = value
End Property

Private Sub UserForm_Initialize()
    Confirmed = False
    lstCases.ListStyle = fmListStyleOption
    lstCases.MultiSelect = fmMultiSelectMulti
    ' The button clears the case selection; its .frx caption may differ.
    btnRefreshCases.Caption = "Clear Cases"
    ' The primary action runs the export.
    btnOK.Caption = "RUN"
End Sub

Private Sub UserForm_Activate()
    Call Relayout
    ' Make the window resizable and reflow on resize (Win32 subclass).
    On Error Resume Next
    Call EnableFormResize(Me, Me.Caption)
    On Error GoTo 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call DisableFormResize
    ' Treat the window close button / Alt+F4 as Cancel, but keep the instance
    ' alive (hidden) so the caller can still read Confirmed.
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Confirmed = False
        Me.Hide
    End If
End Sub

Private Sub UserForm_Terminate()
    Call DisableFormResize
End Sub

' Fill the controls from the current config. The caller assigns Config and then
' calls this before .Show so the dialog opens pre-populated with the preset.
Public Sub PopulateFromConfig()
    Dim i As Long

    If mConfig Is Nothing Then Set mConfig = New cls_EconsConfig

    cboNum.Clear
    For i = 1 To mConfig.MaxOutputs
        cboNum.AddItem CStr(i)
    Next i
    cboNum.value = CStr(mConfig.NumOutputs)

    chkReplace.value = mConfig.ReplaceExisting
    chkCaseFirst.value = mConfig.TitleCaseFirst

    For i = 1 To mConfig.MaxOutputs
        Me.Controls("txtName" & i).Text = mConfig.OutputName(i)
        Me.Controls("txtRange" & i).Text = mConfig.OutputRange(i)
    Next i

    txtCaseRange.Text = mConfig.CaseRangeName

    ' Load the dropdown options for the case cell and pre-check the cases that
    ' were used last time (the saved preset).
    Call LoadCaseOptions(mConfig.CasesArray())

    Call Relayout
End Sub

Private Sub cboNum_Change()
    Call Relayout
End Sub

' Re-read the dropdown options when the case-range name changes (on leaving the
' field), preserving the current checks.
Private Sub txtCaseRange_AfterUpdate()
    Call LoadCaseOptions(CheckedCasesArray())
End Sub

' Clear the case selection: uncheck every item in the list (the options
' themselves stay; only the checks are cleared).
Private Sub btnRefreshCases_Click()
    Dim i As Long
    For i = 0 To lstCases.ListCount - 1
        lstCases.Selected(i) = False
    Next i
End Sub

Private Sub btnAdvanced_Click()
    ' Keep the case range in sync so the advanced dialog reads the same dropdown.
    mConfig.CaseRangeName = Trim$(txtCaseRange.Text)

    UF_EconsAdvanced.InitAdvanced mConfig, AllCaseOptions()
    UF_EconsAdvanced.Show
    Unload UF_EconsAdvanced

    ' The advanced dialog installed its own resize hook (only one form can be
    ' hooked at a time), which removed ours. Re-establish it now that this is the
    ' active window again.
    On Error Resume Next
    Call EnableFormResize(Me, Me.Caption)
    On Error GoTo 0
End Sub

' Open the history picker; if the user loads an entry, deserialize it, prune the
' cases this model no longer offers, adopt it as the current config, and warn
' about anything that was dropped.
Private Sub btnHistory_Click()
    Dim ser As String, note As String
    Dim loaded As cls_EconsConfig

    UF_EconsHistory.InitHistory
    UF_EconsHistory.Show
    ser = ""
    If UF_EconsHistory.Confirmed Then ser = UF_EconsHistory.SelectedSerialized
    Unload UF_EconsHistory

    If ser <> "" Then
        Set loaded = New cls_EconsConfig
        If loaded.Deserialize(ser) Then
            note = EconsPruneToModel(loaded)
            Set mConfig = loaded
            Call PopulateFromConfig
            If note <> "" Then MsgBox note, vbInformation, "Econs - History"
        End If
    End If

    ' The history dialog installed its own resize hook (only one form can be
    ' hooked at a time), which removed ours. Re-establish it.
    On Error Resume Next
    Call EnableFormResize(Me, Me.Caption)
    On Error GoTo 0
End Sub

' Populate the case checkbox list from the case cell's validation dropdown,
' checking any item present in preselect. Falls back to preselect itself when
' the case cell has no list validation, so previously used cases stay available.
Private Sub LoadCaseOptions(ByVal preselect As Variant)
    Dim options As Variant
    Dim i As Long, j As Long

    options = EconsCaseOptions(Trim$(txtCaseRange.Text))
    If Not IsArray(options) Then options = Array()
    If UBound(options) < 0 Then options = preselect

    lstCases.Clear
    If Not IsArray(options) Then Exit Sub
    If UBound(options) < 0 Then Exit Sub

    For i = LBound(options) To UBound(options)
        lstCases.AddItem CStr(options(i))
    Next i

    If IsArray(preselect) Then
        If UBound(preselect) >= 0 Then
            For i = 0 To lstCases.ListCount - 1
                For j = LBound(preselect) To UBound(preselect)
                    If StrComp(lstCases.List(i), CStr(preselect(j)), vbTextCompare) = 0 Then
                        lstCases.Selected(i) = True
                        Exit For
                    End If
                Next j
            Next i
        End If
    End If
End Sub

' Every option currently in the case list (checked or not), for the advanced
' dialog to reuse. Falls back to the saved cases if the list is empty.
Private Function AllCaseOptions() As Variant
    Dim arr() As String
    Dim i As Long

    If lstCases.ListCount = 0 Then
        AllCaseOptions = mConfig.CasesArray()
        Exit Function
    End If

    ReDim arr(0 To lstCases.ListCount - 1)
    For i = 0 To lstCases.ListCount - 1
        arr(i) = lstCases.List(i)
    Next i
    AllCaseOptions = arr
End Function

' Currently checked cases as a 0-based string array (empty Array() if none).
Private Function CheckedCasesArray() As Variant
    Dim arr() As String
    Dim n As Long, i As Long

    n = 0
    For i = 0 To lstCases.ListCount - 1
        If lstCases.Selected(i) Then
            ReDim Preserve arr(0 To n)
            arr(n) = lstCases.List(i)
            n = n + 1
        End If
    Next i

    If n = 0 Then
        CheckedCasesArray = Array()
    Else
        CheckedCasesArray = arr
    End If
End Function

' Currently checked cases as a comma-separated string.
Private Function CheckedCasesCsv() As String
    Dim i As Long
    Dim result As String

    For i = 0 To lstCases.ListCount - 1
        If lstCases.Selected(i) Then
            If result <> "" Then result = result & ", "
            result = result & lstCases.List(i)
        End If
    Next i
    CheckedCasesCsv = result
End Function

' Position and size every control from the form's current dimensions. Output
' rows beyond the selected count are hidden. Called on activate, when the output
' count changes, and on every resize.
Public Sub Relayout()
    On Error Resume Next
    If mInLayout Then Exit Sub
    If mConfig Is Nothing Then Exit Sub
    mInLayout = True

    Dim W As Single, H As Single, y As Single
    Dim n As Long, i As Long
    Dim nameLeft As Single, totalW As Single, nameW As Single
    Dim rangeLeft As Single, rangeW As Single
    Dim refreshLeft As Single, btnTop As Single

    W = Me.InsideWidth
    H = Me.InsideHeight

    ' Header band, with the title text vertically centred over it and a Reset All
    ' button anchored to the top-right corner.
    lblHeaderBand.Left = 0: lblHeaderBand.Top = 0
    lblHeaderBand.Width = W: lblHeaderBand.Height = HEADER_H
    lblHeaderText.Left = MARGIN
    lblHeaderText.Top = (HEADER_H - HEADER_TEXT_H) / 2
    lblHeaderText.Width = W - 2 * MARGIN - RESET_W - GAP
    lblHeaderText.Height = HEADER_TEXT_H
    btnResetAll.Width = RESET_W: btnResetAll.Height = 18
    btnResetAll.Left = W - MARGIN - RESET_W
    btnResetAll.Top = (HEADER_H - 18) / 2

    y = HEADER_H + MARGIN

    ' Number of outputs.
    lblNum.Left = MARGIN: lblNum.Top = y + 2: lblNum.Height = LABEL_H: lblNum.Width = 120
    cboNum.Left = MARGIN + 128: cboNum.Top = y: cboNum.Width = 56: cboNum.Height = FIELD_H
    y = y + FIELD_H + GAP + 2

    ' Replace option, highlighted to stand out.
    chkReplace.Left = MARGIN: chkReplace.Top = y
    chkReplace.Width = W - 2 * MARGIN: chkReplace.Height = REPLACE_H
    y = y + REPLACE_H + 4

    ' Title-order option.
    chkCaseFirst.Left = MARGIN: chkCaseFirst.Top = y
    chkCaseFirst.Width = W - 2 * MARGIN: chkCaseFirst.Height = LABEL_H
    y = y + LABEL_H + GAP + 2

    ' Outputs section header.
    lblSecOutputs.Left = MARGIN: lblSecOutputs.Top = y
    lblSecOutputs.Width = W - 2 * MARGIN: lblSecOutputs.Height = SECTION_H
    y = y + SECTION_H + 2

    ' Column geometry: name column ~58% of the field area, range fills the rest.
    nameLeft = MARGIN + ROWLABEL_W
    totalW = W - 2 * MARGIN - ROWLABEL_W - GAP
    If totalW < 100 Then totalW = 100
    nameW = Int(totalW * 0.58)
    If nameW < 60 Then nameW = 60
    rangeLeft = nameLeft + nameW + GAP
    rangeW = (W - MARGIN) - rangeLeft
    If rangeW < 50 Then rangeW = 50

    lblHdrName.Left = nameLeft: lblHdrName.Top = y: lblHdrName.Width = nameW: lblHdrName.Height = LABEL_H
    lblHdrRange.Left = rangeLeft: lblHdrRange.Top = y: lblHdrRange.Width = rangeW: lblHdrRange.Height = LABEL_H
    y = y + LABEL_H + 2

    ' Output rows (only the selected count are shown).
    n = SelectedCount()
    For i = 1 To mConfig.MaxOutputs
        If i <= n Then
            Me.Controls("lblRow" & i).Left = MARGIN
            Me.Controls("lblRow" & i).Top = y + 2
            Me.Controls("lblRow" & i).Width = ROWLABEL_W
            Me.Controls("lblRow" & i).Height = LABEL_H
            Me.Controls("lblRow" & i).Visible = True

            Me.Controls("txtName" & i).Left = nameLeft
            Me.Controls("txtName" & i).Top = y
            Me.Controls("txtName" & i).Width = nameW
            Me.Controls("txtName" & i).Height = FIELD_H
            Me.Controls("txtName" & i).Visible = True

            Me.Controls("txtRange" & i).Left = rangeLeft
            Me.Controls("txtRange" & i).Top = y
            Me.Controls("txtRange" & i).Width = rangeW
            Me.Controls("txtRange" & i).Height = FIELD_H
            Me.Controls("txtRange" & i).Visible = True

            y = y + ROW_H
        Else
            Me.Controls("lblRow" & i).Visible = False
            Me.Controls("txtName" & i).Visible = False
            Me.Controls("txtRange" & i).Visible = False
        End If
    Next i

    y = y + GAP

    ' Cases section header.
    lblSecCases.Left = MARGIN: lblSecCases.Top = y
    lblSecCases.Width = W - 2 * MARGIN: lblSecCases.Height = SECTION_H
    y = y + SECTION_H + 2

    ' Case input cell row (Clear anchored to the right edge).
    refreshLeft = W - MARGIN - REFRESH_W
    lblCaseRange.Left = MARGIN: lblCaseRange.Top = y + 2: lblCaseRange.Width = 168: lblCaseRange.Height = LABEL_H
    txtCaseRange.Left = MARGIN + 174: txtCaseRange.Top = y: txtCaseRange.Height = FIELD_H
    txtCaseRange.Width = refreshLeft - GAP - (MARGIN + 174)
    If txtCaseRange.Width < 60 Then txtCaseRange.Width = 60
    btnRefreshCases.Left = refreshLeft: btnRefreshCases.Top = y - 1
    btnRefreshCases.Width = REFRESH_W: btnRefreshCases.Height = 20
    y = y + FIELD_H + GAP

    lblCases.Left = MARGIN: lblCases.Top = y: lblCases.Width = W - 2 * MARGIN: lblCases.Height = LABEL_H
    y = y + LABEL_H + 2

    ' Bottom row: Advanced + History (left), OK + Cancel (right).
    btnTop = H - MARGIN - BTN_H
    btnAdvanced.Width = ADV_W: btnAdvanced.Height = BTN_H
    btnAdvanced.Left = MARGIN: btnAdvanced.Top = btnTop
    btnHistory.Width = HIST_W: btnHistory.Height = BTN_H
    btnHistory.Left = btnAdvanced.Left + ADV_W + GAP: btnHistory.Top = btnTop
    btnCancel.Width = BTN_W: btnCancel.Height = BTN_H
    btnCancel.Left = W - MARGIN - BTN_W: btnCancel.Top = btnTop
    btnOK.Width = BTN_W: btnOK.Height = BTN_H
    btnOK.Left = btnCancel.Left - GAP - BTN_W: btnOK.Top = btnTop

    ' Cases list fills the remaining vertical space.
    lstCases.Left = MARGIN: lstCases.Top = y
    lstCases.Width = W - 2 * MARGIN
    lstCases.Height = (btnTop - GAP) - y
    If lstCases.Width < 60 Then lstCases.Width = 60
    If lstCases.Height < 24 Then lstCases.Height = 24

    mInLayout = False
End Sub

Private Function SelectedCount() As Long
    If IsNumeric(cboNum.value) Then
        SelectedCount = CLng(cboNum.value)
    Else
        SelectedCount = mConfig.NumOutputs
    End If
End Function

Private Sub btnOK_Click()
    Dim i As Long
    Dim errMsg As String

    ' Push control values into the config (single source of truth).
    mConfig.NumOutputs = SelectedCount()
    mConfig.ReplaceExisting = (chkReplace.value = True)
    mConfig.TitleCaseFirst = (chkCaseFirst.value = True)
    For i = 1 To mConfig.MaxOutputs
        mConfig.OutputName(i) = Trim$(Me.Controls("txtName" & i).Text)
        mConfig.OutputRange(i) = Trim$(Me.Controls("txtRange" & i).Text)
    Next i
    mConfig.CaseRangeName = Trim$(txtCaseRange.Text)
    mConfig.CasesText = CheckedCasesCsv()

    If Not mConfig.Validate(errMsg) Then
        MsgBox errMsg, vbExclamation, "Econs"
        Exit Sub
    End If

    Confirmed = True
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Confirmed = False
    Me.Hide
End Sub

' Blank all output names and ranges and deselect every case (main and advanced).
' History is stored separately and is left untouched.
Private Sub btnResetAll_Click()
    If MsgBox("Clear all output names and ranges and deselect every case " & _
              "(main and advanced)?" & vbCrLf & "(Your run history is kept.)", _
              vbQuestion + vbYesNo, "Econs - Reset All") <> vbYes Then Exit Sub

    mConfig.ClearAll
    Call PopulateFromConfig
End Sub
