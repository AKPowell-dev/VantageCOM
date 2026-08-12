VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ColorPicker 
   Caption         =   "ColorPicker"
   ClientHeight    =   4425
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   5550
   OleObjectBlob   =   "UF_ColorPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const KEY_LIST1 As String = "wbertyuiop"    ' Theme color selection key (section 1)
Private Const KEY_LIST2 As String = "asdfghjkl;"    ' Default color selection key (section 2)
Private Const KEY_LIST3 As String = "zxcvm"        ' Custom color selection key (section 3)
Private Const KEY_LIST4 As String = "0123456789"      ' Expanded numeric block keys (section 4)
Private Const KEY_LIST4_COUNT As Long = 10            ' Number of keys/columns in the numeric block
Private Const KEY_DETAIL As String = "1234567890"   ' Detail keys (10 entries)
Private Const KEY_NULL As String = "n"              ' Auto/Null selection key
Private Const BORDER_COLOR As Long = vbBlack        ' Color box border color
Private Const DISABLED_COLOR As Long = &HDDDDDD     ' Disabled text/background color

Private Const IDX_LIST1 As Long = 2                 ' Theme color row number (base row)
Private Const IDX_DETAIL_TOP As Long = 5            ' Theme/detail block top row number
Private Const IDX_LIST2 As Long = 28                ' Default color row number
Private Const IDX_LIST3 As Long = 33                ' Custom color row number
Private Const IDX_RESULT As Long = 37               ' Result/status row number

Private Const CUSTOM_BLOCK_START_COL As Long = 12   ' First column index of new block
Private Const TOTAL_COLUMNS As Long = CUSTOM_BLOCK_START_COL + KEY_LIST4_COUNT ' overall grid width

Private Const TEXT_PREFIX As String = " "
Private Const PLACEHOLDER As String = " <SPACE> or #  ->  RGB Color"

Private BOX_GAP As Double

Private cColorTable As Collection
Private cColorObject As Collection
Private cLabelTable As Collection
Private cResultLabel As MSForms.Label
Private cTextLabel As MSForms.Label
Private cFocusLabel As MSForms.Label
Private cCmdBuf As String
Private cResultColor As cls_FontColor

' Generate label names
Private Function GetLabelName(ByVal x As Long, ByVal y As Long) As String
    GetLabelName = "Label_    "
    If x < 10 Then
        Mid$(GetLabelName, 7) = "0"
        Mid$(GetLabelName, 8) = CStr(x)
    Else
        Mid$(GetLabelName, 7) = CStr(x)
    End If

    If y < 10 Then
        Mid$(GetLabelName, 9) = "0"
        Mid$(GetLabelName, 10) = CStr(y)
    Else
        Mid$(GetLabelName, 9) = CStr(y)
    End If
End Function

' Parse coordinates from label name
Private Sub GetXYFromLabelName(ByVal labelName As String, ByRef x As Long, ByRef y As Long)
    If Not labelName Like "Label_[0-9][0-9][0-9][0-9]" Then
        x = -1
        y = -1
    Else
        x = CLng(Mid$(labelName, 7, 2))
        y = CLng(Mid$(labelName, 9, 2))
    End If
End Sub

' Create a label in the grid
Private Function PutLabel(ByVal x As Long, ByVal y As Long, _
                          Optional ByVal xSize As Long = 1, _
                          Optional ByVal ySize As Long = 1) As MSForms.Label
    Set PutLabel = Me.Controls.Add("Forms.Label.1", GetLabelName(x, y), True)
    With PutLabel
        .Left = BOX_GAP + x * (gVim.Config.ColorPickerSize + BOX_GAP)
        .Top = BOX_GAP + (gVim.Config.ColorPickerSize / 2) * y
        .Width = gVim.Config.ColorPickerSize * xSize + BOX_GAP * (xSize - 1)
        .Height = gVim.Config.ColorPickerSize * ySize
        .TextAlign = fmTextAlignCenter
        .Font.Name = "Consolas"
        .Font.Size = gVim.Config.ColorPickerSize / 4 * 3
    End With
End Function

' Create a colour tile
Private Function PutColor(ByVal x As Long, ByVal y As Long, ByRef associatedColor As cls_FontColor, _
                          Optional ByVal BorderColor As Long = vbBlack, _
                          Optional ByVal xSize As Long = 1, _
                          Optional ByVal ySize As Long = 1) As MSForms.Label
    Set PutColor = PutLabel(x, y, xSize, ySize)
    With PutColor
        .BackColor = associatedColor.Color
        .Tag = .BackColor
        If BorderColor <> xlNone Then
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = BorderColor
        End If
    End With
    cColorTable.Add PutColor, PutColor.Name
    cColorObject.Add associatedColor, PutColor.Name
End Function

' Create text label
Private Function PutText(ByVal x As Long, ByVal y As Long, ByVal Caption As String, _
                         Optional ByVal xSize As Long = 1, Optional ySize As Long = 1) As MSForms.Label
    Set PutText = PutLabel(x, y, xSize, ySize)
    PutText.Caption = Caption
    cLabelTable.Add PutText, PutText.Name
End Function

Private Sub UserForm_Activate()
    cCmdBuf = ""
    Set cResultColor = Nothing
    With ActiveWindow
        Me.Top = .Top + (.Height - Me.Height) / 2
        Me.Left = .Left + (.Width - Me.Width) / 2
    End With
End Sub

Private Sub UserForm_Initialize()
    Set cColorTable = New Collection
    Set cLabelTable = New Collection
    Set cColorObject = New Collection

    BOX_GAP = gVim.Config.ColorPickerSize / 4

    Dim i As Long, j As Long
    Dim Color As cls_FontColor
    Dim themeColorLuminance As Long

    Dim defaultColor As Variant
    defaultColor = Array(192, 255, 49407, 65535, 5296274, 5287936, 15773696, 12611584, 6299648, 10498160)

    Dim customColor As Variant
    customColor = Array(gVim.Config.CustomColor1, gVim.Config.CustomColor2, gVim.Config.CustomColor3, _
                        gVim.Config.CustomColor4, gVim.Config.CustomColor5)

    ' --- new section: customize these arrays freely ---------------------------------------
    Dim customBlockTopColors As Variant
    Dim customBlockDetailBaseColors As Variant
    Dim customBlockNavyDetailColors As Variant
    Dim customBlockCeruleanDetailColors As Variant
    Dim customBlockAquaDetailColors As Variant
    Dim customBlockEmeraldDetailColors As Variant
    Dim customBlockSeagreenDetailColors As Variant
    Dim customBlockYellowDetailColors As Variant
    Dim customBlockTerracottaDetailColors As Variant
    Dim customBlockRedDetailColors As Variant
    Dim customBlockSlateDetailColors As Variant
    Dim customBlockLuminance As Long

    ' 0 navy, 1 cerulean, 2 teal, 3 deep teal, 4 green, 5 gold, 6 terracotta, 7 red, 8 slate, 9 graphite
    customBlockTopColors = Array( _
        RGB(0, 32, 96), _
        RGB(38, 118, 178), _
        RGB(62, 158, 168), _
        RGB(52, 108, 108), _
        RGB(200, 221, 200), _
        RGB(222, 175, 99), _
        RGB(153, 70, 60), _
        RGB(190, 30, 30), _
        RGB(72, 88, 112), _
        RGB(58, 65, 75))

    ' Companion tones for the numeric detail grids (slightly lighter than tops).
    customBlockDetailBaseColors = Array( _
        RGB(15, 48, 112), _
        RGB(52, 132, 188), _
        RGB(75, 168, 178), _
        RGB(62, 118, 118), _
        RGB(107, 165, 107), _
        RGB(214, 162, 71), _
        RGB(185, 84, 73), _
        RGB(200, 45, 40), _
        RGB(88, 102, 125), _
        RGB(72, 78, 88))

    ' Navy gradient (position 0): near-white blue -> deep navy.
    customBlockNavyDetailColors = Array( _
        RGB(228, 232, 248), _
        RGB(195, 205, 238), _
        RGB(162, 175, 225), _
        RGB(128, 148, 208), _
        RGB(98, 120, 188), _
        RGB(70, 95, 168), _
        RGB(48, 72, 148), _
        RGB(30, 55, 128), _
        RGB(15, 42, 112), _
        RGB(5, 32, 98))

    ' Cerulean gradient (position 1): pale sky -> deep cerulean.
    customBlockCeruleanDetailColors = Array( _
        RGB(220, 238, 248), _
        RGB(185, 222, 240), _
        RGB(148, 202, 228), _
        RGB(112, 180, 215), _
        RGB(78, 158, 198), _
        RGB(50, 135, 180), _
        RGB(30, 115, 162), _
        RGB(15, 95, 142), _
        RGB(5, 75, 122), _
        RGB(0, 58, 102))

    ' Teal gradient (position 2): pale warm teal -> deep teal (minimal change from prior aqua).
    customBlockAquaDetailColors = Array( _
        RGB(218, 242, 244), _
        RGB(182, 228, 232), _
        RGB(148, 212, 218), _
        RGB(115, 195, 202), _
        RGB(82, 178, 185), _
        RGB(62, 158, 168), _
        RGB(42, 138, 148), _
        RGB(25, 115, 125), _
        RGB(12, 88, 98), _
        RGB(2, 62, 70))

    ' Deep teal gradient (position 3): pale blue-slate teal -> deep muted teal.
    customBlockEmeraldDetailColors = Array( _
        RGB(215, 232, 232), _
        RGB(178, 208, 208), _
        RGB(142, 182, 182), _
        RGB(108, 158, 158), _
        RGB(78, 135, 135), _
        RGB(52, 108, 108), _
        RGB(35, 88, 88), _
        RGB(22, 70, 70), _
        RGB(12, 52, 52), _
        RGB(5, 38, 38))

    ' Green gradient (position 4): #C8DDC8 -> #070C07.
    customBlockSeagreenDetailColors = Array( _
        RGB(187, 213, 187), _
        RGB(160, 197, 160), _
        RGB(134, 181, 134), _
        RGB(107, 165, 107), _
        RGB(87, 142, 87), _
        RGB(72, 117, 72), _
        RGB(56, 90, 56), _
        RGB(39, 65, 39), _
        RGB(23, 38, 23), _
        RGB(7, 12, 7))

    ' Gold gradient (position 5): #FCF8F1 -> #765318.
    customBlockYellowDetailColors = Array( _
        RGB(252, 248, 241), _
        RGB(247, 236, 217), _
        RGB(242, 224, 193), _
        RGB(237, 211, 169), _
        RGB(232, 200, 147), _
        RGB(226, 187, 122), _
        RGB(216, 162, 71), _
        RGB(198, 139, 43), _
        RGB(158, 112, 33), _
        RGB(118, 83, 24))

    ' Terracotta gradient (position 6): #F7EBEA -> #451F1B.
    customBlockTerracottaDetailColors = Array( _
        RGB(247, 235, 234), _
        RGB(234, 205, 202), _
        RGB(221, 175, 170), _
        RGB(209, 145, 137), _
        RGB(197, 115, 105), _
        RGB(184, 84, 73), _
        RGB(137, 63, 54), _
        RGB(115, 52, 45), _
        RGB(91, 41, 36), _
        RGB(69, 31, 27))

    ' Red gradient (position 7): pale rose -> very deep red (wider range, darker).
    customBlockRedDetailColors = Array( _
        RGB(248, 218, 218), _
        RGB(240, 178, 178), _
        RGB(230, 138, 138), _
        RGB(218, 98, 95), _
        RGB(200, 62, 58), _
        RGB(180, 38, 35), _
        RGB(152, 22, 20), _
        RGB(120, 12, 10), _
        RGB(85, 5, 5), _
        RGB(48, 2, 2))

    ' Slate gradient (position 8): pale blue-gray -> deep blue-slate.
    customBlockSlateDetailColors = Array( _
        RGB(225, 230, 238), _
        RGB(198, 205, 218), _
        RGB(168, 178, 198), _
        RGB(140, 152, 175), _
        RGB(115, 128, 155), _
        RGB(92, 105, 135), _
        RGB(72, 85, 115), _
        RGB(55, 68, 98), _
        RGB(40, 52, 78), _
        RGB(28, 38, 62))
    ' --------------------------------------------------------------------------------------

    Dim lncBlack As Variant
    lncBlack = Array(0, 60, 50, 40, 32, 24, 18, 13, 9, 6, 3)

    Dim lncDarkGray As Variant
    lncDarkGray = Array(0, 100, 88, 76, 64, 52, 40, 30, 20, 12, 5)

    Dim lncDefault As Variant
    lncDefault = Array(0, 90, 75, 60, 45, 30, 15, -10, -25, -40, -55)

    Dim lncLightGray As Variant
    lncLightGray = Array(0, -5, -15, -25, -35, -45, -55, -65, -75, -85, -95)

    Dim lncWhite As Variant
    lncWhite = Array(0, -5, -15, -25, -35, -45, -55, -65, -75, -85, -95)

    Dim lncComfortDark As Variant
    lncComfortDark = Array(0, 100, 88, 74, 58, 42, 26, 10, -6, -20, -34)

    Dim lncComfortDefault As Variant
    lncComfortDefault = Array(0, 88, 72, 56, 38, 22, 6, -10, -26, -42, -56)

    Dim lncComfortLight As Variant
    lncComfortLight = Array(0, -8, -18, -28, -38, -48, -58, -68, -78, -88, -96)

    ' Theme colours (unchanged)
    For i = 0 To 9
        Set Color = New cls_FontColor
        Color.Setup msoThemeColorIndex:=i + 1
        themeColorLuminance = Color.luminance

        PutText i + 1, IDX_LIST1 - 2, Mid$(KEY_LIST1, i + 1, 1)
        PutColor i + 1, IDX_LIST1, Color, BORDER_COLOR

        Dim baseColor As Long
        baseColor = Color.Color
        ' Override problematic "r" column to a smoother blue gradient
        If Mid$(KEY_LIST1, i + 1, 1) = "r" Then
            baseColor = RGB(0, 112, 192)
            Color.Setup colorCode:=baseColor
            themeColorLuminance = Color.luminance
        End If

        For j = 1 To Len(KEY_DETAIL)
            Set Color = New cls_FontColor
            Color.Setup colorCode:=baseColor

            Select Case themeColorLuminance
                Case 51 To 203:  Color.AddLuminance = lncDefault(j)
                Case 1 To 50:    Color.AddLuminance = lncDarkGray(j)
                Case 204 To 254: Color.AddLuminance = lncLightGray(j)
                Case 0:          Color.AddLuminance = lncBlack(j)
                Case 255:        Color.AddLuminance = lncWhite(j)
            End Select

            PutColor i + 1, IDX_DETAIL_TOP + (j - 1) * 2, Color
        Next j

        With PutLabel(i + 1, IDX_DETAIL_TOP, ySize:=Len(KEY_DETAIL))
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = BORDER_COLOR
            .BackStyle = fmBackStyleTransparent
        End With

        Set Color = New cls_FontColor
        Color.Setup colorCode:=defaultColor(i)
        PutText i + 1, IDX_LIST2 - 2, Mid$(KEY_LIST2, i + 1, 1)
        PutColor i + 1, IDX_LIST2, Color, BORDER_COLOR
    Next i

    ' Side number labels
    For j = 1 To Len(KEY_DETAIL)
        ChangeState PutText(0, IDX_DETAIL_TOP + (j - 1) * 2, Mid$(KEY_DETAIL, j, 1)), False
        ChangeState PutText(11, IDX_DETAIL_TOP + (j - 1) * 2, Mid$(KEY_DETAIL, j, 1)), False
    Next j

    ' Custom colours (existing five slots)
    For i = 1 To 5
        Set Color = New cls_FontColor
        Color.Setup colorCode:=customColor(i - 1)
        PutText i, IDX_LIST3 - 2, Mid$(KEY_LIST3, i, 1)
        PutColor i, IDX_LIST3, Color, BORDER_COLOR
    Next i

    ' Expanded numeric block (0-9)
    For i = 0 To KEY_LIST4_COUNT - 1
        Set Color = New cls_FontColor
        Color.Setup colorCode:=customBlockTopColors(i)
        PutText CUSTOM_BLOCK_START_COL + i, IDX_LIST1 - 2, Mid$(KEY_LIST4, i + 1, 1)
        PutColor CUSTOM_BLOCK_START_COL + i, IDX_LIST1, Color, BORDER_COLOR

        With PutLabel(CUSTOM_BLOCK_START_COL + i, IDX_DETAIL_TOP, ySize:=Len(KEY_DETAIL))
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = BORDER_COLOR
            .BackStyle = fmBackStyleTransparent
        End With

        For j = 1 To Len(KEY_DETAIL)
            Set Color = New cls_FontColor
            If i = 0 Then
                Color.Setup colorCode:=customBlockNavyDetailColors(j - 1)
            ElseIf i = 1 Then
                Color.Setup colorCode:=customBlockCeruleanDetailColors(j - 1)
            ElseIf i = 2 Then
                Color.Setup colorCode:=customBlockAquaDetailColors(j - 1)
            ElseIf i = 3 Then
                Color.Setup colorCode:=customBlockEmeraldDetailColors(j - 1)
            ElseIf i = 4 Then
                Color.Setup colorCode:=customBlockSeagreenDetailColors(j - 1)
            ElseIf i = 5 Then
                Color.Setup colorCode:=customBlockYellowDetailColors(j - 1)
            ElseIf i = 6 Then
                Color.Setup colorCode:=customBlockTerracottaDetailColors(j - 1)
            ElseIf i = 7 Then
                Color.Setup colorCode:=customBlockRedDetailColors(j - 1)
            ElseIf i = 8 Then
                Color.Setup colorCode:=customBlockSlateDetailColors(j - 1)
            Else
                Color.Setup colorCode:=customBlockDetailBaseColors(i)
                customBlockLuminance = Color.luminance

                Select Case customBlockLuminance
                    Case 51 To 203:  Color.AddLuminance = lncComfortDefault(j)
                    Case 1 To 50:    Color.AddLuminance = lncComfortDark(j)
                    Case 204 To 254: Color.AddLuminance = lncComfortLight(j)
                    Case 0:          Color.AddLuminance = lncComfortDark(j)
                    Case 255:        Color.AddLuminance = lncComfortLight(j)
                End Select
            End If

            PutColor CUSTOM_BLOCK_START_COL + i, IDX_DETAIL_TOP + (j - 1) * 2, Color
        Next j
    Next i

    ' Auto / Null
    Set Color = New cls_FontColor
    Color.Setup colorCode:=0
    Color.IsNull = True

    PutText 6, IDX_LIST3 - 2, KEY_NULL, xSize:=KEY_LIST4_COUNT
    With PutColor(6, IDX_LIST3, Color, xSize:=KEY_LIST4_COUNT)
        .Caption = "Auto, Null"
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = &HA0A0A0
    End With

    ' Result label
    Set cResultLabel = PutLabel(0, IDX_RESULT)
    With cResultLabel
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = BORDER_COLOR
        .BackStyle = fmBackStyleOpaque
        .BackColor = Me.BackColor
    End With

    ' Status text (now stretches across full width)
    Set cTextLabel = PutLabel(1, IDX_RESULT, xSize:=TOTAL_COLUMNS - 1)
    With cTextLabel
        .Caption = PLACEHOLDER
        .TextAlign = fmTextAlignLeft
    End With

    ' Focus rectangle
    Set cFocusLabel = PutLabel(0, 0)
    With cFocusLabel
        .BorderColor = &H1048EF
        .BorderStyle = fmBorderStyleSingle
        .BackStyle = fmBackStyleTransparent
        .Visible = False
    End With

    ' Resize form
    Dim marginWidth As Double:  marginWidth = Me.Width - Me.InsideWidth
    Dim marginHeight As Double: marginHeight = Me.Height - Me.InsideHeight
    Me.Width = marginWidth + gVim.Config.ColorPickerSize * TOTAL_COLUMNS + BOX_GAP * (TOTAL_COLUMNS + 1)
    Me.Height = marginHeight + gVim.Config.ColorPickerSize * (IDX_RESULT / 2 + 1) + BOX_GAP * 2

    Me.DrawBuffer = WorksheetFunction.Min(CLng(Me.InsideHeight * Me.InsideWidth / 9 * 16), 1048576)
End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim addChar As String

    If KeyAscii = Escape_ Then
        If cCmdBuf <> "" Then
            cCmdBuf = ""
        Else
            Quit False
            Exit Sub
        End If
    ElseIf KeyAscii = BackSpace_ Then
        If cCmdBuf <> "" Then
            cCmdBuf = Left$(cCmdBuf, Len(cCmdBuf) - 1)
        End If
    ElseIf KeyAscii = Enter_ Then
        If Not cResultColor Is Nothing Then Quit True
        Exit Sub
    ElseIf KeyAscii > 31 Then
        addChar = LCase$(Chr$(KeyAscii))

        If Len(cCmdBuf) = 0 Then
            If addChar = " " Or addChar = "#" _
               Or InStr(KEY_LIST1 & KEY_LIST2 & KEY_LIST3 & KEY_LIST4 & KEY_NULL, addChar) > 0 Then
                If addChar = " " Then
                    cCmdBuf = cCmdBuf & "#"
                Else
                    cCmdBuf = cCmdBuf & addChar
                End If
            End If
        ElseIf Left$(cCmdBuf, 1) = "#" _
            And InStr("0123456789abcdef", addChar) > 0 _
            And Len(cCmdBuf) < 7 Then
            cCmdBuf = cCmdBuf & addChar
        ElseIf Len(cCmdBuf) = 1 And InStr(KEY_DETAIL, addChar) > 0 Then
            cCmdBuf = cCmdBuf & addChar
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    CheckCommand
End Sub

Private Sub UpdateFocus(Optional ByVal x As Long = -1, Optional ByVal y As Long = -1)
    Dim labelObj As MSForms.Label, labelX As Long, labelY As Long, labelEnabled As Boolean

    For Each labelObj In cColorTable
        If x < 0 Then
            ChangeState labelObj, True
        Else
            GetXYFromLabelName labelObj.Name, labelX, labelY
            labelEnabled = (y < 0 And labelY < IDX_LIST2 And labelX = x) _
                           Or (labelX = x And labelY = y)
            ChangeState labelObj, labelEnabled
        End If
    Next labelObj

    For Each labelObj In cLabelTable
        GetXYFromLabelName labelObj.Name, labelX, labelY
        If labelX = 0 Or labelX = 11 Then
            If x < 0 Then
                ChangeState labelObj, False
            Else
                ChangeState labelObj, (y < 0 Or y = labelY)
            End If
        Else
            If x < 0 Then
                ChangeState labelObj, True
            ElseIf y < 0 Or y < IDX_LIST2 Then
                ChangeState labelObj, (x = labelX And labelY < IDX_LIST2 - 2)
            ElseIf y < IDX_LIST3 Then
                ChangeState labelObj, (x = labelX And labelY = IDX_LIST2 - 2)
            ElseIf y < IDX_RESULT Then
                ChangeState labelObj, (x = labelX And labelY = IDX_LIST3 - 2)
            End If
        End If
    Next labelObj
End Sub

Private Sub CheckCommand()
    Dim x As Long, y As Long, colorY As Long
    Dim firstKey As String, detailIndex As Long

    If cCmdBuf = "" Then
        UpdateFocus
        cTextLabel.Caption = PLACEHOLDER
        cFocusLabel.Visible = False
        Set cResultColor = Nothing

    ElseIf InStr(cCmdBuf, "#") = 1 Then
        If cCmdBuf = "#" Then UpdateFocus 0, 0

        Dim colorCode As String
        Dim colorValue As Long
        colorCode = Mid$(cCmdBuf, 2)
        colorValue = HexColorCodeToLong(colorCode)

        If colorValue < 0 Then
            Set cResultColor = Nothing
        Else
            Set cResultColor = New cls_FontColor
            cResultColor.Setup colorCode:=colorValue
        End If

        cTextLabel.Caption = TEXT_PREFIX & cCmdBuf
        cFocusLabel.Visible = False

    ElseIf LCase$(cCmdBuf) = KEY_NULL Then
        x = 6
        y = IDX_LIST3
        colorY = IDX_LIST3

        UpdateFocus x, y
        Set cResultColor = cColorObject(GetLabelName(x, colorY))
        With cColorTable(GetLabelName(x, colorY))
            cFocusLabel.Left = .Left
            cFocusLabel.Top = .Top
            cFocusLabel.Width = .Width
            cFocusLabel.Height = .Height
            cFocusLabel.Visible = True
        End With
        cTextLabel.Caption = TEXT_PREFIX & cCmdBuf

    ElseIf Len(cCmdBuf) > 0 Then
        firstKey = Left$(cCmdBuf, 1)
        x = InStr(KEY_LIST2 & KEY_LIST3, cCmdBuf)
        If x > 0 Then
            If x > 10 Then
                y = IDX_LIST3
            Else
                y = IDX_LIST2
            End If
            x = (x - 1) Mod 10 + 1
            colorY = y

        ElseIf InStr(KEY_LIST4, firstKey) > 0 Then
            Dim customIndex As Long
            customIndex = InStr(KEY_LIST4, firstKey)
            x = CUSTOM_BLOCK_START_COL + customIndex - 1

            If Len(cCmdBuf) = 1 Then
                y = -1
                colorY = IDX_LIST1
            Else
                detailIndex = InStr(KEY_DETAIL, Mid$(cCmdBuf, 2, 1))
                If detailIndex = 0 Then Exit Sub
                y = IDX_LIST1 + 1 + detailIndex * 2
                colorY = IDX_DETAIL_TOP + (detailIndex - 1) * 2
            End If

        Else
            x = InStr(KEY_LIST1, firstKey)
            If x = 0 Then Exit Sub

            If Len(cCmdBuf) = 1 Then
                y = -1
                colorY = IDX_LIST1
            Else
                detailIndex = InStr(KEY_DETAIL, Mid$(cCmdBuf, 2, 1))
                If detailIndex = 0 Then Exit Sub
                y = IDX_LIST1 + 1 + detailIndex * 2
                colorY = IDX_DETAIL_TOP + (detailIndex - 1) * 2
            End If
        End If

        UpdateFocus x, y

        Set cResultColor = cColorObject(GetLabelName(x, colorY))
        With cColorTable(GetLabelName(x, colorY))
            cFocusLabel.Left = .Left
            cFocusLabel.Top = .Top
            cFocusLabel.Width = .Width
            cFocusLabel.Height = .Height
            cFocusLabel.Visible = True
        End With

        If Not cResultColor.IsNull Then
            cTextLabel.Caption = Left$(TEXT_PREFIX & cCmdBuf & "     ", 6) & "#" & ColorCodeToHex(cResultColor.Color)
        Else
            cTextLabel.Caption = TEXT_PREFIX & cCmdBuf
        End If
    End If

    If cResultColor Is Nothing Then
        cResultLabel.BackColor = Me.BackColor
        cResultLabel.BorderStyle = fmBorderStyleNone
    ElseIf cResultColor.IsNull Then
        cResultLabel.BackColor = Me.BackColor
        cResultLabel.BorderStyle = fmBorderStyleNone
    Else
        cResultLabel.BackColor = cResultColor.Color
        cResultLabel.BorderStyle = fmBorderStyleSingle
    End If
End Sub

Private Sub Quit(ByVal returnResult As Boolean)
    If Not returnResult Then Set cResultColor = Nothing
    cTextLabel.Caption = PLACEHOLDER
    Me.Hide
End Sub

Private Sub ChangeState(ByRef targetLabel As MSForms.Label, ByVal Enabled As Boolean)
    With targetLabel
        If Enabled Then
            If .Tag <> "" Then
                If .BackColor <> .Tag Then .BackColor = .Tag
            End If
            If .ForeColor <> vbBlack Then .ForeColor = vbBlack
        Else
            If .Tag <> "" Then
                If .BackColor <> DISABLED_COLOR Then .BackColor = DISABLED_COLOR
            End If
            If .ForeColor <> DISABLED_COLOR Then .ForeColor = DISABLED_COLOR
        End If
    End With
End Sub

Public Function Launch() As cls_FontColor
    UF_Cmd.Hide
    Me.Show
    Set Launch = cResultColor
    Unload Me
End Function


