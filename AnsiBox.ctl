VERSION 5.00
Begin VB.UserControl AnsiBox 
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4050
   ScaleHeight     =   5400
   ScaleWidth      =   4050
   ToolboxBitmap   =   "AnsiBox.ctx":0000
   Begin VB.ComboBox cboFont 
      Height          =   315
      Left            =   3375
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3765
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3495
      ScaleHeight     =   255
      ScaleWidth      =   195
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2115
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picLarge 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2895
      ScaleHeight     =   465
      ScaleWidth      =   345
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2130
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picCharacterMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7020
      Left            =   15
      ScaleHeight     =   7020
      ScaleWidth      =   2865
      TabIndex        =   2
      Top             =   510
      Width           =   2865
   End
   Begin VB.TextBox txtChars 
      Height          =   285
      Left            =   915
      TabIndex        =   0
      Top             =   45
      Width           =   1695
   End
   Begin VB.PictureBox picShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   495
      Left            =   2940
      ScaleHeight     =   495
      ScaleWidth      =   375
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2175
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Font:"
      Height          =   195
      Index           =   0
      Left            =   2865
      TabIndex        =   7
      Top             =   3780
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Characters to Copy:"
      Height          =   390
      Index           =   1
      Left            =   30
      TabIndex        =   1
      Top             =   15
      Width           =   810
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "AnsiBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type POINTAPI 'Type for holding X & Y co-ordinates
    x As Long
    y As Long
End Type
Private Const strChars As String = _
    "!""#$%&'()*+,-./0123456789:;<=>?" & _
    "@ABCDEFGHIJKLMNÑOPQRSTUVWXYZ[\]^_" & _
    "`abcdefghijklmnñopqrstuvwxyz{|}~" & _
    "€‘’" & _
    " ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿" & _
    "ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞß" & _
    "àáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ" 'All the characters to show
Private Const intCharsPerRow As Integer = 12 'Amount of characters per row
'Button Messages (BM)
Private Const BM_SETSTYLE = &HF4
'Button Styles (BS)
Private Const BS_PUSHBUTTON = &H0&
Private Const BS_USERBUTTON = &H8&
Private intPixelBlockWidth As Integer, _
    intPixelBlockHeight As Integer  'The sizes of the block in pixels
Private Const intMagnification As Integer = 3 'The magnification of the large character
Private Const intShadowOffsetX As Integer = 2, _
    intShadowOffsetY As Integer = 3 'How much to move the shadow over by in pixels
Private intLastOn As Integer 'The last active character
Private sngBlockWidth As Single, sngBlockHeight As Single
Private bolHasFocus As Boolean 'Whether the picture box has focus
Private bolCursorVisible As Boolean 'Whether the cursor is visble or not
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long                      'Finds the cursor's co-ordinates
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long 'API for finding the hWnd of the window under the cursor

Private Const WM_COPY = &H301
Private Const WM_CUT = &H300
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Event CharacterSelected(ByVal Value As String)
Private m_IniFile As String
Public Sub Load()

    On Local Error Resume Next

    'Load the fonts
'    Call LoadFonts(cboFont)
    cboFont.Text = UserControl.Font.Name
    'Get the correct size (i.e. make the most of the size we have)of the blocks in pixels
    'intPixelBlockWidth = (picCharacterMap.ScaleWidth \ intCharsPerRow) \ TwipsX
    'intPixelBlockHeight = (picCharacterMap.ScaleHeight \ (Len(strChars) \ (intCharsPerRow - 1))) \ TwipsY
    intPixelBlockWidth = (picCharacterMap.ScaleWidth / intCharsPerRow) \ TwipsX
    intPixelBlockHeight = (picCharacterMap.ScaleHeight / (Len(strChars) / intCharsPerRow)) \ TwipsY
    
    'Size of the blocks in twips
    sngBlockWidth = TwipsX(intPixelBlockWidth)
    sngBlockHeight = TwipsY(intPixelBlockHeight)
    'Set the temp pic's size to the size of the block + the width of borders on one _
     side only, as the right/bottom will be covered by the next character
    With picTemp
        .Width = sngBlockWidth + TwipsX(picTemp.DrawWidth)
        .Height = sngBlockHeight + TwipsY(picTemp.DrawWidth)
        'Large/preview pic box size
        picLarge.Width = .Width * intMagnification
        picLarge.Height = .Height * intMagnification
        picLarge.FontSize = .FontSize * intMagnification
    End With
    intLastOn = 0
    'Draw the character map
    Call DrawMap

    'Cursor is visible
    bolCursorVisible = True
    
    'Make sure the large chars are at the front
    Call picShadow.ZOrder(vbBringToFront)
    Call picLarge.ZOrder(vbBringToFront)
    
End Sub

Private Sub cboFont_Click()

    On Local Error Resume Next
    With txtChars.Font
        .Name = cboFont.Text
        .Bold = False
        .Italic = False
        .Strikethrough = False
        .Underline = False
    End With
    With cboFont
        picCharacterMap.fontname = .Text
        picLarge.fontname = .Text
    End With
    Call DrawMap
    
End Sub
Private Sub DrawCharacter(ByVal character As String, _
    Optional ByVal Highlighted As Boolean = False, _
    Optional ByVal Focus As Boolean = False)
    On Local Error Resume Next
    With picTemp.Font
        .Bold = False
        .Italic = False
        .Strikethrough = False
        .Underline = False
    End With
    With picTemp
        'Remove old drawings
        .Cls
        'Back/Fore colour = Highlighted or not
        .BackColor = IIf(Highlighted, vbHighlight, vbWindowBackground)
        .ForeColor = IIf(Highlighted, vbHighlightText, vbWindowText)
        'Set the position of the char so that it's centered vertically and horizontally
        .CurrentX = (.ScaleWidth \ 2) - (.TextWidth(character) \ 2)
        .CurrentY = (.ScaleHeight \ 2) - (.TextHeight(character) \ 2)
        'Draw the character
        picTemp.Print character
        'Border
        picTemp.Line (0, 0)-(.ScaleWidth - TwipsX, .ScaleHeight - TwipsY), vbWindowFrame, B
        'Focus rect
        If Focus Then
            'Get the size of the pic box
            Dim rctTemp As RECT
            Call GetClientRect(.hwnd, rctTemp)
            'Move the rect values all in one so we don't end up with a focus rect over the border
            Call InflateRect(rctTemp, -.DrawWidth, -.DrawWidth * 2)
            'Draw the focus
            Call DrawFocusRect(.hdc, rctTemp)
        End If
        'Show changes
        If .AutoRedraw Then .Refresh
    End With
End Sub

Private Sub DrawMap()
    On Local Error Resume Next
    With picCharacterMap
        .Cls
        Dim intLoopCounter As Integer, intRowNumber As Integer, _
            intModulus As Integer
        'Make sure we have the right font
        picTemp.Font.Name = .Font.Name
        'Loop for all chars
        intRowNumber = -1
        For intLoopCounter = 1 To Len(strChars)
            'Get what's left over after dividing by the number of chars per row
            intModulus = (intLoopCounter - 1) Mod intCharsPerRow
            'If it's 0 then it's time to start a new line
            If intModulus = 0 Then intRowNumber = intRowNumber + 1
            'Draw the character to the temp pic box
            Call DrawCharacter(Mid(strChars, intLoopCounter, 1), _
                intLastOn = intLoopCounter, bolHasFocus And intLastOn = intLoopCounter)
            'Now copy it to the correct point in the character map (including the borders)
            Call BitBlt(.hdc, (intModulus * sngBlockWidth) / TwipsX, _
                (intRowNumber * sngBlockHeight) / TwipsY, _
                intPixelBlockWidth + (picTemp.DrawWidth * 2), _
                intPixelBlockHeight + (picTemp.DrawWidth * 2), _
                picTemp.hdc, 0, 0, vbSrcCopy)
        Next intLoopCounter
        '.ScaleLeft = 0
        '.ScaleTop = 0
        '.ScaleWidth = .Width
        '.ScaleHeight = .Height
        '.ScaleWidth = TwipsX(((intModulus * sngBlockWidth) / TwipsX) + intPixelBlockWidth + (picTemp.DrawWidth * 2))
        '.ScaleHeight = TwipsY((((intRowNumber * sngBlockHeight) / TwipsY) + intPixelBlockHeight + (picTemp.DrawWidth * 2)))
        'Show the changes
        If .AutoRedraw Then Call .Refresh
    End With
End Sub

Private Sub HighLightCharacter(ByVal Index As Integer)
    On Local Error Resume Next
    Dim strCharacter As String
    If Index > Len(strChars) Then Index = Len(strChars)
    strCharacter = Mid(strChars, Index, 1)
    With picLarge
        'Remove old drawings
        .Cls
        'Center character
        .CurrentX = (.ScaleWidth \ 2) - (.TextWidth(strCharacter) \ 2)
        .CurrentY = (.ScaleHeight \ 2) - (.TextHeight(strCharacter) \ 2)
        'Draw character
        picLarge.Print strCharacter
        'Show changes
        If .AutoRedraw Then .Refresh
    End With
    
    With picCharacterMap
        Dim sngX As Single, sngY As Single, sngTemp As Single
        sngY = Int(intLastOn / intCharsPerRow) * intPixelBlockHeight
        sngTemp = intLastOn Mod intCharsPerRow
        If sngTemp <> 0 Then
            sngX = (sngTemp - 1) * intPixelBlockWidth
        Else
            sngX = intPixelBlockWidth * (intCharsPerRow - 1)
            sngY = sngY - intPixelBlockHeight
        End If
        
        'Remove the last on character
        If intLastOn >= 0 Then Call DrawCharacter(Mid(strChars, intLastOn, 1))
        Call BitBlt(.hdc, sngX, sngY, intPixelBlockWidth, _
            intPixelBlockHeight, picTemp.hdc, 0, 0, vbSrcCopy)
        
        'Draw the new on character
        Call DrawCharacter(strCharacter, True, bolHasFocus)
        sngY = Int(Index / intCharsPerRow) * intPixelBlockHeight
        sngTemp = Index Mod intCharsPerRow
        If (sngTemp) <> 0 Then
            sngX = (sngTemp - 1) * intPixelBlockWidth
        Else
            sngX = intPixelBlockWidth * (intCharsPerRow - 1)
            sngY = sngY - intPixelBlockHeight
        End If
        Call BitBlt(.hdc, sngX, sngY, _
            intPixelBlockWidth, intPixelBlockHeight, picTemp.hdc, 0, 0, vbSrcCopy)
        If .AutoRedraw Then .Refresh
    End With
    intLastOn = Index
End Sub

Private Sub PositionLargeCharacter(ByVal Index As Integer)
    Dim intRow As Integer, intColumn As Integer
    If Index > Len(strChars) Then Index = Len(strChars)
    intRow = (Index \ (intCharsPerRow)) + 1
    intColumn = Index Mod intCharsPerRow
    If intColumn = 0 Then
        intColumn = intCharsPerRow
        intRow = intRow - 1
    End If
    picLarge.Move picCharacterMap.Left + ((sngBlockWidth * intColumn) - (sngBlockWidth \ 2)) - (picLarge.Width \ 2), _
        picCharacterMap.Top + ((sngBlockHeight * intRow) - (sngBlockHeight \ 2)) - (picLarge.Height \ 2)
    picShadow.Move picLarge.Left + TwipsX(intShadowOffsetX), _
        picLarge.Top + TwipsY(intShadowOffsetY)
    Call SetLargeCharacterVisible
End Sub

Private Sub SetLargeCharacterVisible(Optional ByVal Visible As Boolean = True)
    picLarge.Visible = Visible
    picShadow.Visible = Visible
End Sub

Private Sub picCharacterMap_DblClick()
    On Local Error Resume Next
    txtChars.SelText = Mid(strChars, intLastOn, 1)
End Sub


Private Sub picCharacterMap_GotFocus()
    On Local Error Resume Next
    If Not bolHasFocus Then
        bolHasFocus = True
        Call HighLightCharacter(intLastOn)
    End If
End Sub


Private Sub picCharacterMap_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intTemp As Integer
    Select Case KeyCode
        Case vbKeyLeft
            intTemp = intLastOn - IIf(Shift And vbCtrlMask, 2, 1)
            If intTemp > 0 And intLastOn <> intTemp Then
                Call HighLightCharacter(intTemp)
            Else
                Beep
            End If

        Case vbKeyRight
            intTemp = intLastOn + IIf(Shift And vbCtrlMask, 2, 1)
            If intTemp <= Len(strChars) And intLastOn <> intTemp Then
                Call HighLightCharacter(intTemp)
            Else
                Beep
            End If

        Case vbKeyUp
            If Shift And vbCtrlMask Then
                intTemp = intLastOn Mod intCharsPerRow
                If intTemp = 0 Then intTemp = intCharsPerRow
                If intLastOn <> intTemp And intTemp > 0 Then
                    Call HighLightCharacter(intTemp)
                Else
                    Beep
                End If
            Else
                If intLastOn > intCharsPerRow Then
                    Call HighLightCharacter(intLastOn - intCharsPerRow)
                Else
                    Beep
                End If
            End If
            
        Case vbKeyDown
            If Shift And vbCtrlMask Then
                intTemp = intLastOn Mod intCharsPerRow
                If intTemp = 0 Then intTemp = intCharsPerRow
                intTemp = Len(strChars) - (intCharsPerRow - intTemp)
                If intLastOn <> intTemp And intTemp <= Len(strChars) Then
                    Call HighLightCharacter(intTemp)
                Else
                    Beep
                End If
            Else
                If intLastOn < Len(strChars) - intCharsPerRow + 1 Then
                    Call HighLightCharacter(intLastOn + intCharsPerRow)
                Else
                    Beep
                End If
            End If
        
        Case vbKeyPageUp
            If intLastOn > (intCharsPerRow * 2) Then
                Call HighLightCharacter(intLastOn - (intCharsPerRow * 2))
            Else
                Beep
            End If
            
        Case vbKeyPageDown
            If intLastOn < Len(strChars) - (intCharsPerRow * 2) + 1 Then
                Call HighLightCharacter(intLastOn + (intCharsPerRow * 2))
            Else
                Beep
            End If

        Case vbKeyHome
            If Shift And vbCtrlMask Then
                If intLastOn <> 1 Then
                    Call HighLightCharacter(1)
                Else
                    Beep
                End If
            Else
                intTemp = (intLastOn Mod intCharsPerRow) - 1
                If intTemp = -1 Then intTemp = intCharsPerRow - 1
                If intTemp <> intLastOn And intTemp > 0 Then
                    Call HighLightCharacter(intLastOn - intTemp)
                Else
                    Beep
                End If
            End If
        Case vbKeyEnd
            If Shift And vbCtrlMask Then
                If intLastOn <> Len(strChars) Then
                    Call HighLightCharacter(Len(strChars))
                Else
                    Beep
                End If
            Else
                intTemp = intLastOn Mod intCharsPerRow
                If intLastOn + (intCharsPerRow - intTemp) <> intLastOn And intTemp <> 0 Then
                    Call HighLightCharacter(intLastOn + (intCharsPerRow - intTemp))
                Else
                    Beep
                End If
            End If
        Case Else
            Exit Sub
    End Select
    Call PositionLargeCharacter(intLastOn)
End Sub


Private Sub picCharacterMap_KeyPress(KeyAscii As Integer)
    Dim intCharacterPosition As Integer
    intCharacterPosition = InStr(1, strChars, Chr(KeyAscii))
    If intCharacterPosition > 0 Then
        Call HighLightCharacter(intCharacterPosition)
        Call PositionLargeCharacter(intCharacterPosition)
        On Local Error Resume Next
        txtChars.SelText = Mid(strChars, intLastOn, 1)
    Else
        Call SetLargeCharacterVisible(False)
        Beep
    End If
End Sub


Private Sub picCharacterMap_LostFocus()
    On Local Error Resume Next
    If bolHasFocus Then
        bolHasFocus = False
        Call HighLightCharacter(intLastOn)
        SetLargeCharacterVisible (False)
    End If
End Sub


Private Sub picCharacterMap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Local Error Resume Next
    Call picCharacterMap_MouseMove(Button, Shift, x, y)
    If Button And vbLeftButton Then
        Call SetLargeCharacterVisible
    End If
End Sub


Private Sub picCharacterMap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Local Error Resume Next
    If Button And vbLeftButton Then
        Dim sngTempX As Single, sngTempY As Single
        If x < 0 Then
            sngTempX = 0
        ElseIf x \ sngBlockWidth >= intCharsPerRow Then
            sngTempX = sngBlockWidth * (intCharsPerRow - 1)
        Else
            sngTempX = x
        End If
        sngTempY = (Len(strChars) \ intCharsPerRow)
        If y < 0 Then
            sngTempY = 0
        ElseIf y \ sngBlockHeight >= sngTempY Then
            sngTempY = (sngTempY - 1) * sngBlockHeight
            If Len(strChars) Mod intCharsPerRow Then sngTempY = sngTempY + sngBlockHeight
        Else
            sngTempY = y
        End If
        
        Dim intNewIndex As Integer
        intNewIndex = ((sngTempX \ TwipsX) \ intPixelBlockWidth) + 1 + _
            (((sngTempY \ TwipsY) \ intPixelBlockHeight) * intCharsPerRow)
        If intNewIndex <> intLastOn Then
            Call PositionLargeCharacter(intNewIndex)
            Call HighLightCharacter(intNewIndex)
        End If
        
        Dim rctCharacterMap As RECT
        Call GetWindowRect(picCharacterMap.hwnd, rctCharacterMap)
        If bolCursorVisible And (IsWindowHot(picCharacterMap.hwnd) Or _
            (IsWindowHot(picLarge.hwnd) And IsRECTHot(rctCharacterMap))) Then
            'Hide the cursor
            Call ShowCursor(0)
            bolCursorVisible = False
        ElseIf bolCursorVisible = False And (IsWindowHot(picCharacterMap.hwnd) = False And _
            (IsWindowHot(picLarge.hwnd) = False Or (IsWindowHot(picLarge.hwnd) And IsRECTHot(rctCharacterMap) = False))) Then
            'Show the cursor
            Call ShowCursor(1)
            bolCursorVisible = True
        End If
    End If
End Sub


Private Sub picCharacterMap_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Show the cursor if hidden
    If bolCursorVisible = False Then
        Call ShowCursor(1)
        bolCursorVisible = True
    End If
    If Button = vbLeftButton Then
        Call SetLargeCharacterVisible(False)
    End If
End Sub


Private Sub picLarge_Resize()
    picShadow.Height = picLarge.Height
    picShadow.Width = picLarge.Width
End Sub


Private Function TwipsX(Optional ByVal _
    Amount As Integer = 1) As Single
    On Local Error Resume Next
    'Return the amount of twips in the specified number of pixels
    TwipsX = Amount * Screen.TwipsPerPixelX
End Function

Private Function TwipsY(Optional ByVal _
    Amount As Integer = 1) As Single
    On Local Error Resume Next
    'Return the amount of twips in the specified number of pixels
    TwipsY = Amount * Screen.TwipsPerPixelY
End Function

Private Function IsWindowHot(ByVal hwnd As Long) As Boolean
    On Local Error Resume Next
    Dim CursorPosition As POINTAPI 'Variable for cursor's X & Y values

    'Get the Cursor position
    Call GetCursorPos(CursorPosition)
    IsWindowHot = WindowFromPoint(CursorPosition.x, CursorPosition.y) = hwnd 'Return     whether the object is hot
End Function

Private Function IsRECTHot(Area As RECT) As Boolean
    On Local Error Resume Next
    Dim CursorPosition As POINTAPI 'Variable for cursor's X & Y values

    'Get the Cursor position
    Call GetCursorPos(CursorPosition)
    IsRECTHot = CursorPosition.x >= Area.Left And _
        CursorPosition.x <= Area.Right And _
        CursorPosition.y >= Area.Top And _
        CursorPosition.y <= Area.Bottom
End Function

Private Sub UserControl_Resize()
    On Error Resume Next
    LockWindowUpdate hwnd
    'picCharacterMap.Move 0, 510, UserControl.Width - 810, UserControl.Height
    Call DrawMap
    LockWindowUpdate False
    Err = 0
End Sub



Public Property Get inifile() As String
    inifile = m_IniFile
End Property

Public Property Let inifile(ByVal pIniFile As String)
    m_IniFile = pIniFile
End Property
