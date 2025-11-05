Attribute VB_Name = "Z_ForDevelopper"
Option Explicit
Option Private Module

' CHECK LIST
'  - Nothing for now
Sub SaveAsAddin()
    Dim filePath As String
    Dim verStr As String
    Dim commitStr As String
    
    filePath = Replace(ThisWorkbook.FullName, ".xlsm", "")
    If Right(filePath, 5) <> ".xlam" Then
        filePath = filePath & ".xlam"
    End If
    
    If Dir(filePath) <> "" Then
        Call Kill(filePath)
    End If
    
    'ask version
    Do While True
        verStr = InputBox("version?")
        If verStr = "" Then
            Exit Sub
        End If
        
        commitStr = Trim(InputBox("commit hash?"))
        verStr = "v" & verStr
        If commitStr <> "" Then
            verStr = verStr & " (" & commitStr & ")"
        End If
        
        If MsgBox("Are you sure?" & vbLf & "Version: " & verStr, vbQuestion + vbYesNo) = vbYes Then
            Exit Do
        End If
    Loop
    
    'set comment
    ThisWorkbook.BuiltinDocumentProperties("Comments") = _
        "vim.xlam: " & verStr & vbLf & _
        "Vim experience in Excel" & vbLf & _
        "Source: https://github.com/sha5010/vim.xlam"
    
    'save as addin
    ThisWorkbook.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLAddIn
    
    'delete comment
    'ThisWorkbook.BuiltinDocumentProperties("Comments") = ""
        
    MsgBox "Saved." & vbLf & filePath, vbInformation
End Sub

Sub ExportAll()
    Dim i As Integer
    Dim basePath As String
    Dim moduleName As String
    Dim ext As String
    
    basePath = Left(ThisWorkbook.FullName, InStrRev(ThisWorkbook.FullName, "\"))
    basePath = basePath & "..\src\"

    With ThisWorkbook.VBProject.VBComponents
        For i = 1 To .Count
            moduleName = .item(i).Name
            
            If .item(i).Type = vbext_ct_ClassModule Then
                ext = ".cls"
            ElseIf .item(i).Type = vbext_ct_MSForm Then
                ext = ".frm"
            ElseIf .item(i).Type = vbext_ct_StdModule Then
                ext = ".bas"
            Else
                GoTo Continue
            End If
            
            If moduleName = "ThisWorkbook" Then
                .item(i).Export (basePath & "workbook\ThisWorkbook" & ext)
                Debug.Print basePath & "workbook\ThisWorkbook" & ext
                
            ElseIf InStr(moduleName, "UF_") = 1 Then
                moduleName = Replace(moduleName, "UF_", "")
                .item(i).Export (basePath & "userforms\" & moduleName & ext)
                Debug.Print basePath & "userforms\" & moduleName & ext
            
            ElseIf InStr(moduleName, "cls_") = 1 Then
                moduleName = Replace(moduleName, "cls_", "")
                .item(i).Export (basePath & "classes\" & moduleName & ext)
                Debug.Print basePath & "classes\" & moduleName & ext
                
            ElseIf InStr(moduleName, "A_") = 1 Then
                moduleName = Replace(moduleName, "A_", "")
                .item(i).Export (basePath & moduleName & ext)
                Debug.Print basePath & moduleName & ext
                
            ElseIf InStr(moduleName, "C_") = 1 Then
                moduleName = Replace(moduleName, "C_", "")
                .item(i).Export (basePath & "core\" & moduleName & ext)
                Debug.Print basePath & "core\" & moduleName & ext
                
            ElseIf InStr(moduleName, "F_") = 1 Then
                moduleName = Replace(moduleName, "F_", "")
                .item(i).Export (basePath & "functions\" & moduleName & ext)
                Debug.Print basePath & "functions\" & moduleName & ext
                
            End If
            
Continue:
        Next i
    End With
        
End Sub

Sub ReplaceAll()
    On Error Resume Next

    Dim fso As New FileSystemObject
    Dim modules() As String
    Dim module As Variant
    Dim ext As String
    Dim basePath As String
    Dim prefix As String
    Dim moduleName As String
    
    If MsgBox("All unsaved changes will be lost. Are you sure you want to replace modules?", vbExclamation + vbYesNo + vbDefaultButton2) <> vbYes Then
        Exit Sub
    End If
    
    basePath = Left(ThisWorkbook.FullName, InStrRev(ThisWorkbook.FullName, "\"))
    basePath = basePath & "..\src\"
    
    ReDim modules(0)
    Call SearchAllFile(basePath, modules)
    
    For Each module In modules
        'get extension as lower case
        ext = LCase(fso.GetExtensionName(module))
        
        'set prefix
        If InStr(module, "\classes\") > 0 Then
            prefix = "cls_"
        ElseIf InStr(module, "\core\") > 0 Then
            prefix = "C_"
        ElseIf InStr(module, "\functions\") > 0 Then
            prefix = "F_"
        ElseIf InStr(module, "\userforms\") > 0 Then
            prefix = "UF_"
        ElseIf InStr(module, "DefaultConfig") > 0 Then
            prefix = "A_"
        Else
            prefix = ""
        End If
        
        'import only .cls, .frm, .bas
        If (ext = "cls" Or ext = "frm" Or ext = "bas") And prefix <> "" Then
            With ThisWorkbook.VBProject.VBComponents
                moduleName = prefix & fso.GetBaseName(module)
                .item(moduleName).Name = moduleName & "_old"
                .Remove .item(moduleName & "_old")
                Debug.Print "Removed:", moduleName
            End With
        End If
    Next

    Call ImportAll(0)
End Sub

Sub ImportAll(ByVal Dummy As Long)
    Dim fso As New FileSystemObject
    Dim modules() As String
    Dim module As Variant
    Dim ext As String
    Dim basePath As String
    
    basePath = Left(ThisWorkbook.FullName, InStrRev(ThisWorkbook.FullName, "\"))
    basePath = basePath & "..\src\"
    
    ReDim modules(0)
    Call SearchAllFile(basePath, modules)

    For Each module In modules
        'get extension as lower case
        ext = LCase(fso.GetExtensionName(module))
        
        If (ext = "cls" Or ext = "frm" Or ext = "bas") And InStr(module, "ThisWorkbook.cls") = 0 Then
            Debug.Print "Imported:", module
            ThisWorkbook.VBProject.VBComponents.Import module
        End If
    Next
End Sub
 
' /**
'  * get file paths under specified directory
'  *
'  * @param dirPath  directory path
'  * @param ret      return file paths
'  */
Private Sub SearchAllFile(ByVal dirPath As String, ByRef ret() As String)
    Dim fso As New FileSystemObject
    Dim folder As folder
    Dim subFolder As folder
    Dim file As file
    Dim i As Long
    
    'exit if specified directory is not exists
    If Not fso.FolderExists(dirPath) Then
        Exit Sub
    End If
    
    Set folder = fso.GetFolder(dirPath)
    
    'recursive search
    For Each subFolder In folder.SubFolders
        Call SearchAllFile(subFolder.Path, ret)
    Next
    
    i = UBound(ret)
    
    'find file in the folder
    For Each file In folder.Files
        If (i <> 0 Or ret(i) <> "") Then
            i = i + 1
            ReDim Preserve ret(i)
        End If
        
        'store file path to return array
        ret(i) = file.Path
    Next
End Sub

Function GenerateHelps(Optional ByVal isJP As Boolean = False) As String
    Dim stream As Object
    Dim l As String
    Dim actionDict As Object
    Dim tableStarted As Boolean
    Dim lineCount As Long
    Dim k As Variant
    Dim sk As Variant
    
    If gVim Is Nothing Then
        Call StartVim
    End If
    
    ' Set file path
    Dim filePath As String
    
    If isJP Then
        filePath = ThisWorkbook.Path & "\..\README_ja.md"
    Else
        filePath = ThisWorkbook.Path & "\..\README.md"
    End If
    
    ' Dictionaryを作成
    Set actionDict = New Dictionary
    
    ' Create ADODB.Stream object
    Set stream = CreateObject("ADODB.Stream")
    
    ' Open the stream for reading the UTF-8 file
    With stream
        .Charset = "utf-8"          ' Set character encoding to UTF-8
        .Open
        .LoadFromFile filePath     ' Load the file
        
        ' Initialization
        tableStarted = False
        lineCount = 0
        
        ' Read the file by each Line
        Do Until .EOS
            l = .ReadText(-2) ' Read the Line (lines are read as text with Line breaks)
            lineCount = lineCount + 1
            
            ' Skip until "Expand all commands" appears
            If InStr(l, "<details><summary>") > 0 Then
                tableStarted = True
            End If
            
            ' Obtain data after table started
            If tableStarted Then
                If actionDict.Count > 0 And Trim(l) = "" Then
                    tableStarted = False
                ' | Type | Keystroke | Action | Description | Count |
                ElseIf InStr(l, "|") > 0 And InStr(l, "`") > 0 Then
                    ' Split by "|"
                    Dim cols() As String
                    cols = Split(l, " | ")
                    
                    ' Obtain columns
                    If UBound(cols) >= 4 Then
                        Dim key As String
                        Dim act As String
                        Dim desc As String
                        
                        key = Mid(Trim(cols(1)), 2)
                        key = Left(key, Len(key) - 1)
                        act = Mid(Trim(cols(2)), 2)
                        act = Left(act, Len(act) - 1)
                        desc = Trim(cols(3))
                        
                        ' Check if action exists
                        If Not actionDict.Exists(act) Then
                            actionDict.Add act, New Dictionary
                        End If
                        
                        ' Add description for keystroke
                        For Each k In Split(key, "`/`")
                            k = Replace(k, "\|", "|")
                            k = Replace(k, "[num]", "")
                            k = Replace(k, "[cell]", "")
                            k = Split(k, " ", 2)(0)
                            If InStr(k, ":") <> 1 Then
                                k = gVim.KeyMap.VimToVBA(k, KEY_SEPARATOR)
                            End If
                            
                            If Not actionDict(act).Exists(k) Then
                                actionDict(act).Add k, desc
                            End If
                        Next k
                    End If
                End If
            End If
        Loop
        
        ' Close the stream
        .Close
    End With
    
    Dim fileNumber As Integer
    Dim reading As Boolean
    Dim startLineFound As Boolean
    Dim ret As Dictionary
    Set ret = New Dictionary
    
    ' Set file path
    filePath = ThisWorkbook.Path & "\..\src\DefaultConfig.bas"
    
    ' Open the file for reading
    fileNumber = FreeFile
    On Error GoTo ErrorHandler
    Open filePath For Input As fileNumber
    
    ' Initialization
    reading = False
    startLineFound = False
    
    ' Read the file l by l
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, l
        
        ' Skip lines until we find "With gVim.KeyMap"
        If Not reading Then
            reading = (InStr(l, "With gVim.KeyMap") > 0)
        
        ' Stop reading at "End With"
        ElseIf InStr(l, "End With") > 0 Then
            Exit Do
        End If
        
        ' Process lines with ".Map" pattern
        If reading And InStr(l, ".Map") > 0 Then
            ' Extract $2 and $3 from the ".Map" Line
            l = Trim(Mid(l, InStr(l, ".Map") + 5)) ' Get the part after ".Map"
            l = Replace(l, """", "") ' Remove double quotes
            l = Replace(l, """""", """") ' Replace "" with a single quote (")
            
            ' Split the content by space
            Dim mapParts() As String
            mapParts = Split(l, " ", 3)
            
            If UBound(mapParts) = 2 Then
                Dim rhs As String
                Dim arg As String
                key = mapParts(1)
                rhs = Trim(Split(mapParts(2), "'")(0))
                act = Split(rhs, " ")(0)
                arg = ""
                If InStr(rhs, " ") > 0 Then
                    arg = " """"" & Join(Split(Split(rhs, " ", 2)(1), " "), """"",""""") & """"""
                End If
                                
                If InStr(key, "<cmd>") > 0 Or InStr(key, ":") = 1 Then
                    key = Replace(key, "<cmd>", ":")
                Else
                    key = gVim.KeyMap.VimToVBA(key, KEY_SEPARATOR)
                End If
                
                ' Customization
                If actionDict.Exists(act) Then
                    If actionDict(act).Exists(key) Then
                        If arg <> "" Then
                            k = "'" & act & arg & "'"
                        Else
                            k = act
                        End If
                        
                        If Not ret.Exists(k) Then
                            ret.Add k, ".Add """ & k & """, """ & actionDict(act)(key) & """"
                        End If
                    Else
                        Debug.Print act, key
                    End If
                Else
                    Debug.Print act
                End If
            End If
        End If
    Loop
    
    ' Close the file
    Close fileNumber

    'Set to clipboard
    Dim resultText As String
    For Each k In ret.Items()
        resultText = resultText & String(12, " ") & k & vbCrLf
    Next k
    resultText = resultText & String(12, " ") & "' Automatically generated from README and DefaultConfig"
    
    'With New DataObject
    '    .SetText resultText
    '    .PutInClipboard
    'End With
    
    GenerateHelps = resultText
    Exit Function

ErrorHandler:
    Close fileNumber
    MsgBox "Error: " & Err.Description
End Function

Sub ReplaceTest()
    Dim clsPath As String
    clsPath = ThisWorkbook.Path & "\..\src\classes\Help.cls"
    
    Dim tempPath As String
    tempPath = clsPath & ".tmp"
    
    Dim fIn As Integer, fOut As Integer
    Dim Line As String
    Dim insideJP As Boolean, insideEN As Boolean
    
    insideJP = False
    insideEN = False
    
    fIn = FreeFile
    Open clsPath For Input As #fIn
    
    fOut = FreeFile
    Open tempPath For Output As #fOut
    
    Do While Not EOF(fIn)
        Line Input #fIn, Line
        
        If Trim(Line) Like "With HELP_DICT_JP" Then
            ' JPセクション開始
            Print #fOut, Line
            Print #fOut, GenerateHelps(True)
            ' JPセクションの元のコードはスキップするのでEnd Withまで読み飛ばし
            Do While Not EOF(fIn)
                Line Input #fIn, Line
                If Trim(Line) = "End With" Then Exit Do
            Loop
        ElseIf Trim(Line) Like "With HELP_DICT_EN*" Then
            ' ENセクション開始
            Print #fOut, Line
            Print #fOut, GenerateHelps(False)
            ' ENセクションの元のコードをEnd Withまで読み飛ばす
            Do While Not EOF(fIn)
                Line Input #fIn, Line
                If Trim(Line) = "End With" Then Exit Do
            Loop
        End If
        
        Print #fOut, Line
    Loop
    
    Close #fIn
    Close #fOut
    
    If Dir(clsPath & ".bak") <> "" Then
        Kill clsPath & ".bak"
    End If
    Name clsPath As clsPath & ".bak"
    Name tempPath As clsPath
    
    ' 既存プロジェクトを上書き
    With ThisWorkbook.VBProject.VBComponents
        .item("cls_Help").Name = "cls_Help_old"
        .Remove .item("cls_Help_old")
        .Import clsPath
    End With
End Sub
