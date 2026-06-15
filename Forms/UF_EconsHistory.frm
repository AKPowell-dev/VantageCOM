VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_EconsHistory
   Caption         =   "Econs - History"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9120.001
   OleObjectBlob   =   "UF_EconsHistory.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_EconsHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' History picker for the econs export: lists the last few successful runs (most
' recent first) in a two-column list - selected output names, then the cases.
' The caller reads Confirmed and SelectedSerialized after .Show, deserializes
' the chosen entry, and prunes it to the current model.
'
' Like the other econs dialogs, the layout is computed at run time by Relayout
' so the window can be resized (via C_FormResize).

' --- Layout metrics (points) ---
Private Const MARGIN As Single = 12
Private Const GAP As Single = 6
Private Const HEADER_H As Single = 24
Private Const HEADER_TEXT_H As Single = 18
Private Const LABEL_H As Single = 16
Private Const BTN_W As Single = 84
Private Const BTN_H As Single = 26

Private mSerials() As String
Private mDetails() As String
Private mCasesColW As Single
Private mTitlesColW As Single
Private mInLayout As Boolean

Public Confirmed As Boolean
Public SelectedSerialized As String

' Fill the list from the stored history. Call before .Show.
Public Sub InitHistory()
    Dim n As Long, i As Long, row As Long
    Dim cfg As cls_EconsConfig
    Dim ser As String, titles As String, cases As String
    Dim seen As Object, dispKey As String
    Dim col2txt As String, maxLen As Long, maxTitleLen As Long

    Confirmed = False
    SelectedSerialized = ""
    lstHistory.Clear

    n = EconsHistoryCount()
    ReDim mSerials(0 To IIf(n > 0, n - 1, 0))
    ReDim mDetails(0 To IIf(n > 0, n - 1, 0))
    txtDetails.Text = ""
    mCasesColW = 0
    mTitlesColW = 0

    If n = 0 Then
        lstHistory.AddItem "(no history yet)"
        Exit Sub
    End If

    Set cfg = New cls_EconsConfig
    Set seen = CreateObject("Scripting.Dictionary")
    row = 0
    For i = 1 To n
        ser = EconsHistorySerialized(i)
        If ser <> "" Then
            If cfg.Deserialize(ser) Then
                titles = cfg.TitleSummary()
                cases = cfg.CasesSummary()

                ' Skip rows that look identical to one already listed (the most
                ' recent wins) so duplicates don't take up space.
                dispKey = titles & Chr$(1) & cases
                If Not seen.Exists(dispKey) Then
                    seen(dispKey) = True

                    If titles = "" Then titles = "(no outputs)"
                    If cases = "" Then cases = "(no cases)"
                    If i = 1 Then cases = cases & "    (last run)"

                    lstHistory.AddItem titles
                    If Len(titles) > maxTitleLen Then maxTitleLen = Len(titles)
                    ' Leading bar divides the titles column from the cases column.
                    col2txt = "|  " & cases
                    lstHistory.List(row, 1) = col2txt
                    If Len(col2txt) > maxLen Then maxLen = Len(col2txt)
                    mSerials(row) = ser
                    mDetails(row) = cfg.DetailSummary()
                    row = row + 1
                End If
            End If
        End If
    Next i

    ' Widths (points) each column needs to show its longest entry in full, so
    ' neither titles nor cases are clipped; Relayout applies these and the list
    ' scrolls horizontally to the very end.
    mTitlesColW = maxTitleLen * 7 + 12
    mCasesColW = maxLen * 7 + 12

    ' Show the most recent run's full details by default.
    If lstHistory.ListCount > 0 Then lstHistory.ListIndex = 0
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
    ' Treat the window close button / Alt+F4 as Cancel.
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Confirmed = False
        Me.Hide
    End If
End Sub

Private Sub UserForm_Terminate()
    Call DisableFormResize
End Sub

Private Sub lstHistory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call btnHistLoad_Click
End Sub

' Show the selected run's full, wrapped detail in the read-only details box.
Private Sub lstHistory_Change()
    Dim r As Long

    r = lstHistory.ListIndex
    If r >= 0 And r <= UBound(mDetails) Then
        txtDetails.Text = mDetails(r)
    Else
        txtDetails.Text = ""
    End If
End Sub

Private Sub btnHistLoad_Click()
    Dim r As Long

    r = lstHistory.ListIndex
    If r < 0 Then
        MsgBox "Select a history entry first.", vbExclamation, "Econs - History"
        Exit Sub
    End If

    ' Guard the "(no history yet)" placeholder and any empty slot.
    If r > UBound(mSerials) Then Exit Sub
    If mSerials(r) = "" Then
        MsgBox "There is no saved run to load.", vbInformation, "Econs - History"
        Exit Sub
    End If

    SelectedSerialized = mSerials(r)
    Confirmed = True
    Me.Hide
End Sub

Private Sub btnHistCancel_Click()
    Confirmed = False
    Me.Hide
End Sub

' Position and size every control from the form's current dimensions. The list
' fills the space between the hint and the button row. Called on activate and on
' every resize.
Public Sub Relayout()
    On Error Resume Next
    If mInLayout Then Exit Sub
    mInLayout = True

    Dim W As Single, H As Single, y As Single, btnTop As Single
    Dim listW As Single, col1 As Single, col2 As Single, avail As Single, detailH As Single

    W = Me.InsideWidth
    H = Me.InsideHeight

    ' Header band with the title text vertically centred over it.
    lblHeaderBand.Left = 0: lblHeaderBand.Top = 0
    lblHeaderBand.Width = W: lblHeaderBand.Height = HEADER_H
    lblHeaderText.Left = MARGIN
    lblHeaderText.Top = (HEADER_H - HEADER_TEXT_H) / 2
    lblHeaderText.Width = W - 2 * MARGIN
    lblHeaderText.Height = HEADER_TEXT_H

    y = HEADER_H + MARGIN
    lblHistHint.Left = MARGIN: lblHistHint.Top = y
    lblHistHint.Width = W - 2 * MARGIN: lblHistHint.Height = LABEL_H
    y = y + LABEL_H + GAP

    btnTop = H - MARGIN - BTN_H

    ' The list and the (wrapped, scrollable) details box share the middle: the
    ' details box takes a fixed slice at the bottom so the selected run's full
    ' text is always visible; the list fills the rest.
    avail = (btnTop - GAP) - y
    detailH = avail * 0.32
    If detailH < 54 Then detailH = 54
    If detailH > avail - 40 Then detailH = avail - 40

    lstHistory.Left = MARGIN: lstHistory.Top = y
    lstHistory.Width = W - 2 * MARGIN
    lstHistory.Height = avail - detailH - GAP
    If lstHistory.Width < 80 Then lstHistory.Width = 80
    If lstHistory.Height < 40 Then lstHistory.Height = 40
    listW = lstHistory.Width
    ' Each column is sized to its longest entry so neither titles nor cases are
    ' clipped; the cases column then fills any leftover width. When the total
    ' exceeds the visible area the list scrolls horizontally to the very end.
    col1 = mTitlesColW
    If col1 < 80 Then col1 = 80
    col2 = mCasesColW
    If col2 < listW - col1 Then col2 = listW - col1
    lstHistory.ColumnCount = 2
    lstHistory.ColumnWidths = col1 & ";" & col2

    txtDetails.Left = MARGIN
    txtDetails.Top = lstHistory.Top + lstHistory.Height + GAP
    txtDetails.Width = W - 2 * MARGIN
    txtDetails.Height = detailH

    ' Button row: Load + Cancel anchored bottom-right.
    btnHistCancel.Width = BTN_W: btnHistCancel.Height = BTN_H
    btnHistCancel.Left = W - MARGIN - BTN_W: btnHistCancel.Top = btnTop
    btnHistLoad.Width = BTN_W: btnHistLoad.Height = BTN_H
    btnHistLoad.Left = btnHistCancel.Left - GAP - BTN_W: btnHistLoad.Top = btnTop

    mInLayout = False
End Sub
