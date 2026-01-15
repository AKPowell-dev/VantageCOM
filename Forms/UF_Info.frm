VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Info 
   Caption         =   "Vantage Info"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
   OleObjectBlob   =   "UF_Info.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    With Me
        .Caption = "Vantage Info"
        .StartUpPosition = 0
    End With

    With Me.TextBox
        .value = ""
        .Multiline = True
        .EnterKeyBehavior = True
        .WordWrap = False
        .TabKeyBehavior = False
        .Locked = True
        .ScrollBars = fmScrollBarsVertical
        .Top = 24
        .Left = 6
        .Width = Me.Width - 24
        .Height = Me.Height - 36
        .Font.Name = "Consolas"
        .Font.Size = 9
        .BackColor = vbWhite
    End With

    With Me.Label_Prefix
        .Caption = "X"
        .AutoSize = True
        .Font.Bold = True
        .Left = Me.Width - .Width - 12
        .Top = 6
    End With
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    Me.Move Application.Left + Application.Width - Me.Width - 10, Application.Top + 10
    On Error GoTo 0

    Me.TextBox.SetFocus
End Sub

Private Sub Label_Prefix_Click()
    Unload Me
End Sub

Public Sub ShowInfo(ByVal infoText As String, Optional ByVal formCaption As String = "Vantage Info")
    Me.Caption = formCaption
    Me.TextBox.Text = infoText
    Me.Show vbModeless
End Sub
