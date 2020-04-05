VERSION 5.00
Object = "{5F37140E-C836-11D2-BEF8-525400DFB47A}#1.1#0"; "vbalTab6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl MapaHtml 
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4545
   ScaleHeight     =   3675
   ScaleWidth      =   4545
   ToolboxBitmap   =   "MapaHtml.ctx":0000
   Begin VB.TextBox txtChars 
      Height          =   285
      Left            =   1455
      TabIndex        =   8
      Top             =   15
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvwChar 
      Height          =   570
      Left            =   2880
      TabIndex        =   7
      Top             =   1860
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1005
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Character"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Entity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Preview"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox picCharacterMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1560
      Left            =   1095
      ScaleHeight     =   1560
      ScaleWidth      =   2865
      TabIndex        =   6
      Top             =   1425
      Width           =   2865
   End
   Begin VB.PictureBox picLarge 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2715
      ScaleHeight     =   465
      ScaleWidth      =   345
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3315
      ScaleHeight     =   255
      ScaleWidth      =   195
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ComboBox cboFont 
      Height          =   315
      Left            =   3195
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2610
      Visible         =   0   'False
      Width           =   2415
   End
   Begin vbalTabStrip6.TabControl tabMap 
      Height          =   2100
      Left            =   105
      TabIndex        =   0
      Top             =   450
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   3704
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   495
      Left            =   2760
      ScaleHeight     =   495
      ScaleWidth      =   375
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Characters to Copy:"
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1620
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Font:"
      Height          =   195
      Index           =   0
      Left            =   2685
      TabIndex        =   5
      Top             =   2625
      Visible         =   0   'False
      Width           =   390
   End
End
Attribute VB_Name = "MapaHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ultima As Integer

Private Type eHtml
    character As String
    entity As String
    preview As String
End Type
Private arr_html() As eHtml
Private Const limite_fila = 9

Public Event SpecialCharSelected(ByVal Value As String)

Private m_IniFile As String

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type POINTAPI 'Type for holding X & Y co-ordinates
    X As Long
    Y As Long
End Type
Private strChars As String 'All the characters to show
Private Const intCharsPerRow As Integer = 10 'Amount of characters per row
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
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long                      'Finds the cursor's co-ordinates
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long 'API for finding the hWnd of the window under the cursor

Private Const WM_COPY = &H301
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CUT = &H300
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Sub Load()

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

Public Sub LoadMap()

    Dim src As New cStringBuilder
    Dim sSections() As String
    Dim k As Integer
    Dim c As Integer
    Dim ele As String
    
    'cargar tabs
    With tabMap
        If .TabCount = 0 Then
            .AddTab "Grid View", , , "Grid", 1000
            .AddTab "Table View", 1, , "Table", 2000
            '.CoolTabs = etaDevStudio
            .Rebuild
        End If
    End With
            
    'cargar mapa de caracteres especiales
    get_info_section "mapahtml", sSections, m_IniFile
    
    strChars = ""
    
    ReDim arr_html(0)
    c = 1
    For k = 2 To UBound(sSections)
        ele = sSections(k)
        src.Append Util.Explode(ele, 2, ";")
        ReDim Preserve arr_html(c)
        arr_html(c).character = Util.Explode(ele, 3, ";")
        arr_html(c).entity = Util.Explode(ele, 1, ";")
        arr_html(c).preview = Util.Explode(ele, 2, ";")
        
        lvwChar.ListItems.Add , "k" & c, arr_html(c).character
        lvwChar.ListItems(c).SubItems(1) = arr_html(c).entity
        lvwChar.ListItems(c).SubItems(2) = arr_html(c).preview
        c = c + 1
    Next k
    
    strChars = src.ToString '& Space$(50)
            
    UserControl_Resize
    
    Call Load
    
    tabMap_TabClick 1
    
End Sub

Public Sub LoadMap2()

'    Dim k As Integer
'    Dim j As Integer
'    Dim t As Integer
'    Dim c As Integer
'    Dim l As Integer
'    Dim p As Integer
'    Dim num
'    Dim Linea As String
'
'    Dim sSections() As String
'    Dim Atributos() As String
'
'    get_info_section "mapahtml", sSections, m_IniFile
'
'    lblDescrip.Caption = ""
'    lblcodigo.Caption = ""
'
'    pic(0).Move 0, 0
'
'    ReDim arr_html(0)
'    c = 0
'    For k = 2 To UBound(sSections)
'        Linea = sSections(k)
'
'        If Len(Linea) > 0 Then
'            ReDim Preserve arr_html(c)
'            arr_html(c).traduccion = Util.Explode(Linea, 1, ";")
'            arr_html(c).caracter = Util.Explode(Linea, 2, ";")
'            arr_html(c).ayuda = Util.Explode(Linea, 3, ";")
'            c = c + 1
'        End If
'    Next k
'
'    t = pic(0).Top
'    l = pic(0).Left
'    c = 0
'
'    For j = 0 To UBound(arr_html)
'
'        If j > 0 Then
'            Load pic(j)
'        End If
'
'        If c = 0 Then
'            pic(j).Left = pic(0).Left
'        Else
'            pic(j).Left = pic(c - 1).Left + pic(j - 1).Height
'        End If
'
'        pic(j).Top = t
'        pic(j).Width = pic(0).Width
'        pic(j).Visible = True
'        pic(j).CurrentX = 60
'        pic(j).CurrentY = 3
'        pic(j).Font.Size = 10
'
'        pic(j).Print arr_html(j).caracter
'        pic(j).tag = c & "#" & arr_html(j).caracter & "#" & arr_html(j).traduccion & "#" & arr_html(j).ayuda
'
'        If c > limite_fila Then
'            c = 0
'            t = t + pic(0).Height
'        Else
'            c = c + 1
'        End If
'    Next j


End Sub


Private Sub tabMap_TabClick(ByVal lTab As Long)
    
    If lTab = 1 Then
        lvwChar.Visible = False
        picCharacterMap.Visible = True
        'picCharacterMap.ZOrder 0
    Else
        picCharacterMap.Visible = False
        lvwChar.Visible = True
    End If
    
End Sub


Private Sub UserControl_Resize()
    On Error Resume Next
    LockWindowUpdate hWnd
    tabMap.Move 0, 345, UserControl.Width, UserControl.Height - 345
    picCharacterMap.Move 50, 700, UserControl.Width - 100, UserControl.Height - 800
    lvwChar.Move picCharacterMap.Left, picCharacterMap.Top, picCharacterMap.Width, picCharacterMap.Height
    LockWindowUpdate False
    Err = 0
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
            Call GetClientRect(.hWnd, rctTemp)
            'Move the rect values all in one so we don't end up with a focus rect over the border
            Call InflateRect(rctTemp, -.DrawWidth, -.DrawWidth * 2)
            'Draw the focus
            Call DrawFocusRect(.hDC, rctTemp)
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
            Call BitBlt(.hDC, (intModulus * sngBlockWidth) / TwipsX, _
                (intRowNumber * sngBlockHeight) / TwipsY, _
                intPixelBlockWidth + (picTemp.DrawWidth * 2), _
                intPixelBlockHeight + (picTemp.DrawWidth * 2), _
                picTemp.hDC, 0, 0, vbSrcCopy)
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
        Call BitBlt(.hDC, sngX, sngY, intPixelBlockWidth, _
            intPixelBlockHeight, picTemp.hDC, 0, 0, vbSrcCopy)
        
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
        Call BitBlt(.hDC, sngX, sngY, _
            intPixelBlockWidth, intPixelBlockHeight, picTemp.hDC, 0, 0, vbSrcCopy)
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
    'txtChars.SelText = Mid(strChars, intLastOn, 1)
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
        'txtChars.SelText = Mid(strChars, intLastOn, 1)
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


Private Sub picCharacterMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    Call picCharacterMap_MouseMove(Button, Shift, X, Y)
    If Button And vbLeftButton Then
        Call SetLargeCharacterVisible
    End If
End Sub


Private Sub picCharacterMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    If Button And vbLeftButton Then
        Dim sngTempX As Single, sngTempY As Single
        If X < 0 Then
            sngTempX = 0
        ElseIf X \ sngBlockWidth >= intCharsPerRow Then
            sngTempX = sngBlockWidth * (intCharsPerRow - 1)
        Else
            sngTempX = X
        End If
        sngTempY = (Len(strChars) \ intCharsPerRow)
        If Y < 0 Then
            sngTempY = 0
        ElseIf Y \ sngBlockHeight >= sngTempY Then
            sngTempY = (sngTempY - 1) * sngBlockHeight
            If Len(strChars) Mod intCharsPerRow Then sngTempY = sngTempY + sngBlockHeight
        Else
            sngTempY = Y
        End If
        
        Dim intNewIndex As Integer
        intNewIndex = ((sngTempX \ TwipsX) \ intPixelBlockWidth) + 1 + _
            (((sngTempY \ TwipsY) \ intPixelBlockHeight) * intCharsPerRow)
        If intNewIndex <> intLastOn Then
            Call PositionLargeCharacter(intNewIndex)
            Call HighLightCharacter(intNewIndex)
        End If
        
        Dim rctCharacterMap As RECT
        Call GetWindowRect(picCharacterMap.hWnd, rctCharacterMap)
        If bolCursorVisible And (IsWindowHot(picCharacterMap.hWnd) Or _
            (IsWindowHot(picLarge.hWnd) And IsRECTHot(rctCharacterMap))) Then
            'Hide the cursor
            Call ShowCursor(0)
            bolCursorVisible = False
        ElseIf bolCursorVisible = False And (IsWindowHot(picCharacterMap.hWnd) = False And _
            (IsWindowHot(picLarge.hWnd) = False Or (IsWindowHot(picLarge.hWnd) And IsRECTHot(rctCharacterMap) = False))) Then
            'Show the cursor
            Call ShowCursor(1)
            bolCursorVisible = True
        End If
    End If
End Sub


Private Sub picCharacterMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Function IsWindowHot(ByVal hWnd As Long) As Boolean
    On Local Error Resume Next
    Dim CursorPosition As POINTAPI 'Variable for cursor's X & Y values

    'Get the Cursor position
    Call GetCursorPos(CursorPosition)
    IsWindowHot = WindowFromPoint(CursorPosition.X, CursorPosition.Y) = hWnd 'Return     whether the object is hot
End Function

Private Function IsRECTHot(Area As RECT) As Boolean
    On Local Error Resume Next
    Dim CursorPosition As POINTAPI 'Variable for cursor's X & Y values

    'Get the Cursor position
    Call GetCursorPos(CursorPosition)
    IsRECTHot = CursorPosition.X >= Area.Left And _
        CursorPosition.X <= Area.Right And _
        CursorPosition.Y >= Area.Top And _
        CursorPosition.Y <= Area.Bottom
End Function



Public Property Get IniFile() As String
    IniFile = m_IniFile
End Property

Public Property Let IniFile(ByVal pIniFile As String)
    m_IniFile = pIniFile
End Property
