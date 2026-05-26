VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_ShortcutItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCategory As String
Private mAction As String
Private mDefaultKey As String
Private mCurrentKey As String

Public Property Get Category() As String
    Category = mCategory
End Property

Public Property Let Category(ByVal value As String)
    mCategory = value
End Property

Public Property Get Action() As String
    Action = mAction
End Property

Public Property Let Action(ByVal value As String)
    mAction = value
End Property

Public Property Get DefaultKey() As String
    DefaultKey = mDefaultKey
End Property

Public Property Let DefaultKey(ByVal value As String)
    mDefaultKey = value
End Property

Public Property Get CurrentKey() As String
    CurrentKey = mCurrentKey
End Property

Public Property Let CurrentKey(ByVal value As String)
    mCurrentKey = value
End Property

Public Property Get IsEnabled() As Boolean
    IsEnabled = (Len(Trim$(mCurrentKey)) > 0)
End Property
