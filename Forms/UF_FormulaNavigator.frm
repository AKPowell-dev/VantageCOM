VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_FormulaNavigator 
   Caption         =   "Formula"
   ClientHeight    =   405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1950
   OleObjectBlob   =   "UF_FormulaNavigator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_FormulaNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mStatusLabel As MSForms.Label
Private mProgressBack As MSForms.Label
Private mProgressFill As MSForms.Label
Private mProgressDot As MSForms.Label

Private Sub HandleNavigatorKey(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyH, vbKeyJ, vbKeyK, vbKeyL
                Call FormulaNavigatorCancel
                KeyCode = 0
                Exit Sub
        End Select
    End If

    If KeyCode = vbKeyEscape Then
        If (Shift And 2) = 2 Then
            Call FormulaNavigatorNext
        Else
            Call FormulaNavigatorCancel
        End If
        KeyCode = 0
        Exit Sub
    End If

    If KeyCode = OpeningSquareBracket_ And (Shift And 2) = 2 Then
        Call FormulaNavigatorNext
        KeyCode = 0
    End If
End Sub

Public Sub Launch(ByVal formulaText As String, ByVal highlightStart As Long, ByVal highlightLen As Long, _
                  Optional ByVal refIndex As Long = 0, Optional ByVal refCount As Long = 0, _
                  Optional ByVal currentToken As String = "", Optional ByVal startAddress As String = "")
    Call EnsureUiElements
    Call LayoutControls
    If Me.TextBox.Text <> formulaText Then
        Me.TextBox.Text = formulaText
    End If

    Call ApplyHighlight(highlightStart, highlightLen)
    Call UpdateCycleStatus(refIndex, refCount, currentToken, startAddress)

    If Not Me.Visible Then
        Me.Show vbModeless
    End If

    Call MoveTopRight
    On Error Resume Next
    Me.TextBox.SetFocus
    Call ApplyHighlight(highlightStart, highlightLen)
    On Error GoTo 0
End Sub

Public Sub UpdateDisplay(ByVal formulaText As String, ByVal highlightStart As Long, ByVal highlightLen As Long)
    Call Launch(formulaText, highlightStart, highlightLen)
End Sub

Public Sub HideNavigator()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    With Me
        .Caption = "Formula"
        .StartUpPosition = 0
    End With

    If Me.Height < 88 Then Me.Height = 88
    If Me.Width < 300 Then Me.Width = 300

    Call EnsureUiElements

    On Error Resume Next
    Me.Label_Prefix.Visible = False
    Me.Label_Prefix.Caption = ""
    Me.Label_Prefix.Width = 0
    On Error GoTo 0

    With Me.TextBox
        .Value = ""
        .MultiLine = True
        .WordWrap = False
        .ScrollBars = fmScrollBarsHorizontal
        .Locked = True
        .EnterKeyBehavior = False
        .TabKeyBehavior = False
        .HideSelection = False
    End With

    Call LayoutControls
End Sub

Private Sub UserForm_Activate()
    Call MoveTopRight
    On Error Resume Next
    Me.TextBox.SetFocus
    On Error GoTo 0
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call HandleNavigatorKey(KeyCode, Shift)
End Sub

Private Sub TextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call HandleNavigatorKey(KeyCode, Shift)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub LayoutControls()
    On Error Resume Next
    Call EnsureUiElements

    Dim padding As Single
    Dim infoGap As Single
    Dim statusHeight As Single
    Dim progressHeight As Single
    Dim infoHeight As Single

    padding = 4
    infoGap = 3
    statusHeight = 10
    progressHeight = 4
    infoHeight = statusHeight + progressHeight + infoGap * 2

    Me.TextBox.Left = padding
    Me.TextBox.Top = padding
    Me.TextBox.Width = Me.InsideWidth - padding * 2
    Me.TextBox.Height = Me.InsideHeight - infoHeight - padding

    mStatusLabel.Left = padding
    mStatusLabel.Top = Me.TextBox.Top + Me.TextBox.Height + infoGap
    mStatusLabel.Width = Me.TextBox.Width
    mStatusLabel.Height = statusHeight

    mProgressBack.Left = padding
    mProgressBack.Top = mStatusLabel.Top + statusHeight + infoGap
    mProgressBack.Width = Me.TextBox.Width
    mProgressBack.Height = progressHeight

    mProgressFill.Left = mProgressBack.Left
    mProgressFill.Top = mProgressBack.Top
    mProgressFill.Height = mProgressBack.Height

    mProgressDot.Top = mProgressBack.Top - 1
    mProgressDot.Height = mProgressBack.Height + 2
    On Error GoTo 0
End Sub

Private Sub EnsureUiElements()
    On Error Resume Next
    Set mStatusLabel = Me.Controls("Label_Status")
    Set mProgressBack = Me.Controls("Label_ProgressBack")
    Set mProgressFill = Me.Controls("Label_ProgressFill")
    Set mProgressDot = Me.Controls("Label_ProgressDot")
    On Error GoTo 0

    If mStatusLabel Is Nothing Then
        Set mStatusLabel = Me.Controls.Add("Forms.Label.1", "Label_Status", True)
        With mStatusLabel
            .Caption = ""
            .TextAlign = fmTextAlignLeft
            .Font.Name = "Consolas"
            .Font.Size = 8
            .BackStyle = fmBackStyleTransparent
        End With
    End If

    If mProgressBack Is Nothing Then
        Set mProgressBack = Me.Controls.Add("Forms.Label.1", "Label_ProgressBack", True)
        With mProgressBack
            .Caption = ""
            .BackColor = RGB(224, 224, 224)
            .BorderStyle = fmBorderStyleSingle
        End With
    End If

    If mProgressFill Is Nothing Then
        Set mProgressFill = Me.Controls.Add("Forms.Label.1", "Label_ProgressFill", True)
        With mProgressFill
            .Caption = ""
            .BackColor = RGB(0, 120, 215)
            .BorderStyle = fmBorderStyleNone
        End With
    End If

    If mProgressDot Is Nothing Then
        Set mProgressDot = Me.Controls.Add("Forms.Label.1", "Label_ProgressDot", True)
        With mProgressDot
            .Caption = ""
            .BackColor = RGB(0, 120, 215)
            .BorderStyle = fmBorderStyleSingle
            .Width = 6
        End With
    End If
End Sub

Private Sub UpdateCycleStatus(ByVal refIndex As Long, ByVal refCount As Long, _
                              ByVal currentToken As String, ByVal startAddress As String)
    Call EnsureUiElements

    Dim totalSteps As Long
    Dim stepIndex As Long
    Dim statusText As String

    If refCount < 0 Then refCount = 0
    totalSteps = refCount + 1

    If refIndex <= 0 Then
        stepIndex = 1
        statusText = "Base: " & FormatBaseAddress(startAddress)
    Else
        stepIndex = refIndex + 1
        statusText = "Ref " & CStr(refIndex) & "/" & CStr(refCount) & ": " & currentToken
    End If

    If Len(statusText) > 120 Then
        statusText = Left$(statusText, 117) & "..."
    End If
    mStatusLabel.Caption = statusText

    Call UpdateProgress(stepIndex, totalSteps)
End Sub

Private Function FormatBaseAddress(ByVal startAddress As String) As String
    Dim bangPos As Long
    Dim sheetPart As String
    Dim addr As String
    Dim bracketPos As Long

    bangPos = InStrRev(startAddress, "!")
    If bangPos = 0 Then
        FormatBaseAddress = startAddress
        Exit Function
    End If

    sheetPart = Left$(startAddress, bangPos - 1)
    addr = Mid$(startAddress, bangPos + 1)

    bracketPos = InStr(sheetPart, "]")
    If bracketPos > 0 Then
        sheetPart = Mid$(sheetPart, bracketPos + 1)
    End If

    If Len(sheetPart) >= 2 And Left$(sheetPart, 1) = "'" And Right$(sheetPart, 1) = "'" Then
        sheetPart = Mid$(sheetPart, 2, Len(sheetPart) - 2)
        sheetPart = Replace(sheetPart, "''", "'")
    End If

    addr = Replace(addr, "$", "")
    If Len(sheetPart) = 0 Then
        FormatBaseAddress = addr
    Else
        FormatBaseAddress = sheetPart & "!" & addr
    End If
End Function

Private Sub UpdateProgress(ByVal stepIndex As Long, ByVal totalSteps As Long)
    If mProgressBack Is Nothing Or mProgressFill Is Nothing Or mProgressDot Is Nothing Then Exit Sub

    If totalSteps <= 1 Then
        mProgressFill.Width = 0
        mProgressDot.Visible = False
        Exit Sub
    End If

    If stepIndex < 1 Then stepIndex = 1
    If stepIndex > totalSteps Then stepIndex = totalSteps

    Dim trackWidth As Single
    Dim ratio As Double
    trackWidth = mProgressBack.Width
    ratio = (stepIndex - 1) / (totalSteps - 1)

    mProgressFill.Width = trackWidth * ratio
    mProgressDot.Left = mProgressBack.Left + trackWidth * ratio - (mProgressDot.Width / 2)
    mProgressDot.Visible = True

    If mProgressDot.Left < mProgressBack.Left Then
        mProgressDot.Left = mProgressBack.Left
    End If
    If mProgressDot.Left > mProgressBack.Left + trackWidth - mProgressDot.Width Then
        mProgressDot.Left = mProgressBack.Left + trackWidth - mProgressDot.Width
    End If

    Call PulseProgress
End Sub

Private Sub PulseProgress()
    Dim oldColor As Long
    oldColor = mProgressDot.BackColor
    mProgressDot.BackColor = RGB(255, 153, 0)
    DoEvents
    mProgressDot.BackColor = oldColor
End Sub

Private Sub ApplyHighlight(ByVal highlightStart As Long, ByVal highlightLen As Long)
    Dim textLen As Long
    textLen = Len(Me.TextBox.Text)

    If highlightStart < 0 Then highlightStart = 0
    If highlightStart > textLen Then highlightStart = textLen
    If highlightLen < 0 Then highlightLen = 0
    If highlightStart + highlightLen > textLen Then
        highlightLen = textLen - highlightStart
    End If

    Me.TextBox.SelStart = highlightStart
    Me.TextBox.SelLength = highlightLen
End Sub

Private Sub MoveTopRight()
    Dim margin As Double
    Dim newLeft As Double

    margin = 12
    newLeft = Application.Left + Application.Width - Me.Width - margin
    If newLeft < Application.Left + margin Then newLeft = Application.Left + margin
    Me.Move newLeft, Application.Top + margin
End Sub
