Attribute VB_Name = "C_NetBridge"
Option Explicit
Option Private Module

Private Const COM_ADDIN_PROG_ID As String = "VantagePackageHolder.Addin"
Private mNetAddin As Object

Public Function NetAddin() As Object
    On Error GoTo Fail
    If Not mNetAddin Is Nothing Then
        ' Validate cached reference is still alive
        Dim dummy As String
        dummy = TypeName(mNetAddin)
        If dummy = "Nothing" Or dummy = "" Then Set mNetAddin = Nothing
    End If

    If mNetAddin Is Nothing Then
        Dim addinObj As COMAddIn
        For Each addinObj In Application.COMAddIns
            If StrComp(addinObj.ProgId, COM_ADDIN_PROG_ID, vbTextCompare) = 0 Then
                Set mNetAddin = addinObj.Object
                Exit For
            End If
        Next addinObj
    End If
    Set NetAddin = mNetAddin
    Exit Function
Fail:
    Set mNetAddin = Nothing
    Set NetAddin = Nothing
End Function
