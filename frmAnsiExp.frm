VERSION 5.00
Object = "{FCFAF346-DE8A-4FB6-8612-5000548EFDC7}#2.0#0"; "vbsListView6.ocx"
Object = "{D890B066-6CE9-4233-9AC2-5E66E7917BF3}#2.0#0"; "vbsTab6.ocx"
Begin VB.Form frmAnsiExp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ansi Browser"
   ClientHeight    =   7785
   ClientLeft      =   1095
   ClientTop       =   2370
   ClientWidth     =   10050
   Icon            =   "frmAnsiExp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   75
      ScaleHeight     =   330
      ScaleWidth      =   9870
      TabIndex        =   9
      Top             =   7335
      Width           =   9900
      Begin VB.Label lbltag 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   60
         Width           =   45
      End
   End
   Begin VB.PictureBox picLarge 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2790
      ScaleHeight     =   465
      ScaleWidth      =   345
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Double clic select character"
      Top             =   1125
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   495
      Left            =   2835
      ScaleHeight     =   495
      ScaleWidth      =   375
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1185
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3390
      ScaleHeight     =   255
      ScaleWidth      =   195
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1125
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picCharacterMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6525
      Left            =   75
      ScaleHeight     =   6525
      ScaleWidth      =   9900
      TabIndex        =   1
      ToolTipText     =   "Double clic select character"
      Top             =   720
      Width           =   9900
   End
   Begin VB.TextBox txtChars 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   870
      TabIndex        =   0
      Top             =   75
      Width           =   9105
   End
   Begin vbalTabStrip6.TabControl tabMap 
      Height          =   2100
      Left            =   15
      TabIndex        =   4
      Top             =   420
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
   Begin vbalListViewLib6.vbalListViewCtl lvwChar 
      Height          =   1980
      Left            =   285
      TabIndex        =   8
      Top             =   795
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   3493
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   1
      MultiSelect     =   -1  'True
      LabelEdit       =   0   'False
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
      AutoArrange     =   0   'False
      Appearance      =   0
      FlatScrollBar   =   -1  'True
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Font:"
      Height          =   195
      Index           =   0
      Left            =   2760
      TabIndex        =   7
      Top             =   2790
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected :"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   105
      Width           =   855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAnsiExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private ultima As Integer

Private Type eHtml
    character As String
    entity As String
    preview As String
End Type
Private arr_html() As eHtml

Public Event SpecialCharSelected(ByVal Value As String)

Private m_IniFile As String

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
Private strChars As String 'All the characters to show
Private Const intCharsPerRow As Integer = 10 'Amount of characters per row
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
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long                      'Finds the cursor's co-ordinates
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long 'API for finding the hWnd of the window under the cursor

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Function buscar_entity(ByVal str As String) As String

    Dim k As Integer
    
    For k = 1 To UBound(arr_html)
        DoEvents
        If str = arr_html(k).character Then
            buscar_entity = arr_html(k).entity
            Exit Function
        End If
    Next k
        
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    On Local Error Resume Next

    Dim src As New cStringBuilder
    Dim sSections() As String
    Dim k As Integer
    Dim C As Integer
    Dim ele As String
    
    util.CenterForm Me
    util.Hourglass hwnd, True
        
    'cargar tabs
    With tabMap
        .AddTab "Grid View", , , "Grid", 1000
        .AddTab "Table View", 1, , "Table", 2000
        .Rebuild
    End With
                
    With lvwChar
        .Columns.Add , "k1", "Character", , 3000
        .Columns.Add , "k2", "Alt Code", , 1440
    End With
    
    'cargar mapa de caracteres especiales
    get_info_section "ansi", sSections, m_IniFile
    
    strChars = vbNullString
    
    ReDim arr_html(0)
    C = 1
    For k = 2 To UBound(sSections)
        ele = sSections(k)
        src.Append util.Explode(ele, 1, "#")
        ReDim Preserve arr_html(C)
        arr_html(C).character = util.Explode(ele, 1, "#")
        arr_html(C).entity = util.Explode(ele, 2, "#")
                
        lvwChar.ListItems.Add , "k" & C, arr_html(C).character
        lvwChar.ListItems(C).SubItems(1).Caption = arr_html(C).entity
        C = C + 1
    Next k
    
    strChars = src.ToString '& Space$(50)
    
    'Load the fonts
'    Call LoadFonts(cboFont)
    'cboFont.Text = Me.Font.Name
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
    
    Form_Resize
    
    'Draw the character map
    Call DrawMap

    'Cursor is visible
    bolCursorVisible = True
    
    'Make sure the large chars are at the front
    Call picShadow.ZOrder(vbBringToFront)
    Call picLarge.ZOrder(vbBringToFront)
    
    util.Hourglass hwnd, False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub lvwChar_ItemClick(ITem As vbalListViewLib6.cListItem)
    txtChars.Text = ITem.Text
End Sub

Private Sub tabMap_TabClick(ByVal lTab As Long)
    
    If lTab = 1 Then
        lvwChar.Visible = False
        picCharacterMap.Visible = True
        picCharacterMap.ZOrder 0
    Else
        picCharacterMap.Visible = False
        lvwChar.ZOrder 0
        lvwChar.Visible = True
    End If
    
    'Make sure the large chars are at the front
    Call picShadow.ZOrder(vbBringToFront)
    Call picLarge.ZOrder(vbBringToFront)
    
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
        'Debug.Print "caracter : " & strCharacter
        lbltag.Caption = buscar_entity(strCharacter)
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



Public Property Get inifile() As String
    inifile = m_IniFile
End Property

Public Property Let inifile(ByVal pIniFile As String)
    m_IniFile = pIniFile
End Property

Private Sub Form_Resize()

    On Error Resume Next
    If WindowState <> vbMinimized Then
        LockWindowUpdate hwnd
        tabMap.Move 0, 380, Width - 100, Height - 1250
        'picCharacterMap.Move 50, 700, Width - 100, Height - 800
        lvwChar.Move picCharacterMap.Left, picCharacterMap.Top, picCharacterMap.Width, picCharacterMap.Height
        LockWindowUpdate False
        Err = 0
    End If
    
End Sub


