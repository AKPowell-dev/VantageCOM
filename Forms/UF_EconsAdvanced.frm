VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_EconsAdvanced
   Caption         =   "Econs - Advanced"
   ClientHeight    =   8340.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "UF_EconsAdvanced.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_EconsAdvanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Advanced econs settings: up to two extra per-case outputs (the "3rd / 4th"
' outputs). Each has a name, a named range, and a checkbox list of the cases it
' applies to (the same case options as the main dialog). Edits the shared
' cls_EconsConfig in place; values are written back (and validated) only on OK.
'
' Like the main dialog, the layout is computed at run time by Relayout from the
' form's current size, so the window can be resized (via C_FormResize) and all
' controls reflow. Cosmetic properties (fonts/colours) are baked into the .frx.

Private Const ADV_COUNT As Long = 2

' --- Layout metrics (points), matched to UF_EconsConfig for a consistent look.
Private Const MARGIN As Single = 12
Private Const GAP As Single = 6
Private Const HEADER_H As Single = 24
Private Const HEADER_TEXT_H As Single = 18
Private Const LABEL_H As Single = 16
Private Const SECTION_H As Single = 16
Private Const FIELD_H As Single = 18
Private Const BTN_W As Single = 84
Private Const BTN_H As Single = 26
Private Const NAME_LBL_W As Single = 40
Private Const RANGE_LBL_W As Single = 46
Private Const MIN_LIST_H As Single = 48
Private Const ADVCLR_W As Single = 80

Private mConfig As cls_EconsConfig
Private mOptions As Variant
Private mInLayout As Boolean

' Called by the main dialog before .Show. caseOptions is the list of case names
' to offer (the main dialog's dropdown options).
Public Sub InitAdvanced(ByVal cfg As cls_EconsConfig, ByVal caseOptions As Variant)
    Dim i As Long

    Set mConfig = cfg
    If IsArray(caseOptions) Then
        mOptions = caseOptions
    Else
        mOptions = Array()
    End If

    For i = 1 To ADV_COUNT
        Me.Controls("txtAdvName" & i).Text = cfg.AdvName(i)
        Me.Controls("txtAdvRange" & i).Text = cfg.AdvRange(i)
        Call FillCaseList(Me.Controls("lstAdvCases" & i), cfg.AdvCasesArray(i))
    Next i
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    For i = 1 To ADV_COUNT
        With Me.Controls("lstAdvCases" & i)
            .ListStyle = fmListStyleOption
            .MultiSelect = fmMultiSelectMulti
        End With
        Me.Controls("btnAdvClear" & i).Caption = "Clear Cases"
    Next i
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
    ' Treat the window close button / Alt+F4 as Cancel (discard edits).
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub UserForm_Terminate()
    Call DisableFormResize
End Sub

' Populate a checkbox list with the shared case options, checking those in sel.
Private Sub FillCaseList(ByVal lst As Object, ByVal sel As Variant)
    Dim i As Long, j As Long

    lst.Clear
    If Not IsArray(mOptions) Then Exit Sub
    If UBound(mOptions) < 0 Then Exit Sub

    For i = LBound(mOptions) To UBound(mOptions)
        lst.AddItem CStr(mOptions(i))
    Next i

    If IsArray(sel) Then
        If UBound(sel) >= 0 Then
            For i = 0 To lst.ListCount - 1
                For j = LBound(sel) To UBound(sel)
                    If StrComp(lst.List(i), CStr(sel(j)), vbTextCompare) = 0 Then
                        lst.Selected(i) = True
                        Exit For
                    End If
                Next j
            Next i
        End If
    End If
End Sub

Private Function CheckedCsv(ByVal lst As Object) As String
    Dim i As Long, result As String

    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            If result <> "" Then result = result & ", "
            result = result & lst.List(i)
        End If
    Next i
    CheckedCsv = result
End Function

' Position and size every control from the form's current dimensions. The two
' extra-output blocks share the vertical space between the header and the button
' row, each with a stretchy case list. Called on activate and on every resize.
Public Sub Relayout()
    On Error Resume Next
    If mInLayout Then Exit Sub
    mInLayout = True

    Dim W As Single, H As Single, y As Single
    Dim idx As Long
    Dim btnTop As Single, listH As Single, blockFixed As Single
    Dim availW As Single, colW As Single
    Dim nameFieldLeft As Single, nameFieldW As Single
    Dim rangeLblLeft As Single, rangeFieldLeft As Single, rangeFieldW As Single

    W = Me.InsideWidth
    H = Me.InsideHeight

    ' Header band with the title text vertically centred over it.
    lblHeaderBand.Left = 0: lblHeaderBand.Top = 0
    lblHeaderBand.Width = W: lblHeaderBand.Height = HEADER_H
    lblHeaderText.Left = MARGIN
    lblHeaderText.Top = (HEADER_H - HEADER_TEXT_H) / 2
    lblHeaderText.Width = W - 2 * MARGIN
    lblHeaderText.Height = HEADER_TEXT_H

    btnTop = H - MARGIN - BTN_H

    ' Each block contributes a section header, a name/range row and an
    ' "apply to" label before its (stretchy) case list. Split the leftover
    ' vertical space evenly between the two lists.
    blockFixed = SECTION_H + 2 + FIELD_H + GAP + LABEL_H + 2
    listH = ((btnTop - GAP) - (HEADER_H + MARGIN) - ADV_COUNT * (blockFixed + GAP)) / ADV_COUNT
    If listH < MIN_LIST_H Then listH = MIN_LIST_H

    ' Two equal columns: Name on the left, Range on the right.
    availW = W - 2 * MARGIN
    colW = (availW - GAP) / 2
    If colW < 120 Then colW = 120
    nameFieldLeft = MARGIN + NAME_LBL_W
    nameFieldW = colW - NAME_LBL_W
    If nameFieldW < 60 Then nameFieldW = 60
    rangeLblLeft = MARGIN + colW + GAP
    rangeFieldLeft = rangeLblLeft + RANGE_LBL_W
    rangeFieldW = (W - MARGIN) - rangeFieldLeft
    If rangeFieldW < 60 Then rangeFieldW = 60

    y = HEADER_H + MARGIN
    For idx = 1 To ADV_COUNT
        Me.Controls("lblSec" & idx).Left = MARGIN
        Me.Controls("lblSec" & idx).Top = y
        Me.Controls("lblSec" & idx).Width = W - 2 * MARGIN
        Me.Controls("lblSec" & idx).Height = SECTION_H
        y = y + SECTION_H + 2

        Me.Controls("lblAdvName" & idx).Left = MARGIN
        Me.Controls("lblAdvName" & idx).Top = y + 2
        Me.Controls("lblAdvName" & idx).Width = NAME_LBL_W
        Me.Controls("lblAdvName" & idx).Height = LABEL_H
        Me.Controls("txtAdvName" & idx).Left = nameFieldLeft
        Me.Controls("txtAdvName" & idx).Top = y
        Me.Controls("txtAdvName" & idx).Width = nameFieldW
        Me.Controls("txtAdvName" & idx).Height = FIELD_H

        Me.Controls("lblAdvRange" & idx).Left = rangeLblLeft
        Me.Controls("lblAdvRange" & idx).Top = y + 2
        Me.Controls("lblAdvRange" & idx).Width = RANGE_LBL_W
        Me.Controls("lblAdvRange" & idx).Height = LABEL_H
        Me.Controls("txtAdvRange" & idx).Left = rangeFieldLeft
        Me.Controls("txtAdvRange" & idx).Top = y
        Me.Controls("txtAdvRange" & idx).Width = rangeFieldW
        Me.Controls("txtAdvRange" & idx).Height = FIELD_H
        y = y + FIELD_H + GAP

        Me.Controls("lblApply" & idx).Left = MARGIN
        Me.Controls("lblApply" & idx).Top = y + 2
        Me.Controls("lblApply" & idx).Width = W - 2 * MARGIN - ADVCLR_W - GAP
        Me.Controls("lblApply" & idx).Height = LABEL_H
        Me.Controls("btnAdvClear" & idx).Left = W - MARGIN - ADVCLR_W
        Me.Controls("btnAdvClear" & idx).Top = y
        Me.Controls("btnAdvClear" & idx).Width = ADVCLR_W
        Me.Controls("btnAdvClear" & idx).Height = 20
        y = y + 22

        Me.Controls("lstAdvCases" & idx).Left = MARGIN
        Me.Controls("lstAdvCases" & idx).Top = y
        Me.Controls("lstAdvCases" & idx).Width = W - 2 * MARGIN
        Me.Controls("lstAdvCases" & idx).Height = listH
        y = y + listH + GAP
    Next idx

    ' Button row: OK + Cancel anchored to the bottom-right.
    btnAdvCancel.Width = BTN_W: btnAdvCancel.Height = BTN_H
    btnAdvCancel.Left = W - MARGIN - BTN_W: btnAdvCancel.Top = btnTop
    btnAdvOK.Width = BTN_W: btnAdvOK.Height = BTN_H
    btnAdvOK.Left = btnAdvCancel.Left - GAP - BTN_W: btnAdvOK.Top = btnTop

    mInLayout = False
End Sub

Private Sub btnAdvOK_Click()
    Dim i As Long
    Dim errMsg As String

    ' Push control values into the shared config (single source of truth).
    For i = 1 To ADV_COUNT
        mConfig.AdvName(i) = Trim$(Me.Controls("txtAdvName" & i).Text)
        mConfig.AdvRange(i) = Trim$(Me.Controls("txtAdvRange" & i).Text)
        mConfig.AdvCasesText(i) = CheckedCsv(Me.Controls("lstAdvCases" & i))
    Next i

    ' Enforce the same completeness rule the main dialog applies on OK: each
    ' extra output must be fully specified or fully blank.
    If Not mConfig.ValidateAdvanced(errMsg) Then
        MsgBox errMsg, vbExclamation, "Econs - Advanced"
        Exit Sub
    End If

    Me.Hide
End Sub

Private Sub btnAdvCancel_Click()
    Me.Hide
End Sub

Private Sub btnAdvClear1_Click()
    Call ClearAdvCases(1)
End Sub

Private Sub btnAdvClear2_Click()
    Call ClearAdvCases(2)
End Sub

' Uncheck every case for an extra-output slot. Its name and range are kept.
Private Sub ClearAdvCases(ByVal idx As Long)
    Dim i As Long

    With Me.Controls("lstAdvCases" & idx)
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next i
    End With
End Sub
