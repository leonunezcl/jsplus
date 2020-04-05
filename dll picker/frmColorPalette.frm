VERSION 5.00
Begin VB.Form frmColorPalette 
   BorderStyle     =   0  'None
   ClientHeight    =   3885
   ClientLeft      =   4155
   ClientTop       =   1485
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmColorPalette.frx":0000
   ScaleHeight     =   259
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   251
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmColorPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Private WithEvents m_cMenu As cPopupMenu
Attribute m_cMenu.VB_VarHelpID = -1

Private bIsMouseOver As Boolean
Private bOverMore As Boolean
Private m_lStartX As Long
Private m_lStartY As Long

Private Enum enumPalettes
  cc2 = 0
  cc8 = 1
  cc16 = 2
  ccVB = 3
  cc256 = 4
  cc256Gray = 5
  ccWin = 6
  ccNamed = 7
  ccSafe = 8
  ccCustom = 9
End Enum

Private PtAPI As PointAPI
Private Declare Function ChooseColor Lib "COMDLG32.DLL" Alias "ChooseColorA" (pChoosecolor As udtCHOOSECOLOR) As Long

Private Type udtCHOOSECOLOR
  lStructSize As Long
  hWndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
Private Const CC_FULLOPEN = &H2
Private Const CC_ANYCOLOR = &H100
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private m_lBtnCount As Long
Private m_lColorCount As Long
Private m_lSelectedPallete As enumPalettes
Private LastSavedCustClr As Long
Public SelectedColor As OLE_COLOR
Public PathPaleta As String
Public CodeColor As String
Public IsCanceled As Boolean
Public IsMouseOver As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyEscape) Then
    Me.Hide
  End If
End Sub

Private Sub Form_Load()
  Dim R As RECT
    
  Me.ScaleMode = vbPixels
  Me.Font.Name = "Arial"
    
  Call SetCapture(Me.hwnd)
    
  IsCanceled = True
  IsMouseOver = False
  
  m_lStartX = 3
  m_lStartY = 28
  pRows = 1
  pCols = 1
  
  'm_sLastPal = App.Path & "\Palettes\16 Colors Palette.pal"
  'If Right$(PathPaleta, 1) <> "\" Then
  '  PathPaleta = PathPaleta & "\"
  'End If
   
  m_sLastPal = PathPaleta
     
  'MsgBox m_sLastPal
  
  m_lSelectedPallete = cc16
  m_lColorCount = LoadPalette(m_sLastPal, m_oClrNames)
  
  Me.Width = (m_lStartX * 2 + pCols * (m_lBoxSize + m_lSpace)) * Screen.TwipsPerPixelX
  Me.Height = (m_lStartY + (pRows + 2) * (m_lBoxSize + m_lSpace) + 5 * m_lStartX) * Screen.TwipsPerPixelY
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Long, j As Long
  Dim IsMouseOnBut As Boolean
  Dim m_iClrIndex As Long
  Dim R As RECT
  
  If Not (Button = 1) Then Exit Sub
  
  IsMouseOnBut = (x >= m_lStartX + (m_lBoxSize + m_lSpace) * 17 And y >= m_lStartX) And (x <= ScaleWidth - m_lStartX And y <= m_lStartY - 2 * m_lStartX)
  If IsMouseOnBut Then
    Call SetRect(R, m_lStartX + (m_lBoxSize + m_lSpace) * 17, m_lStartX, ScaleWidth - m_lStartX, m_lStartY - 2 * m_lStartX)
    Call DrawRect(hdc, R, vbButtonText, vbWhite, , True)
  End If
  
  'default button
  IsMouseOnBut = (x >= m_lStartX And y >= m_lStartY + 3 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace)) And (x <= pCols / 2 * (m_lBoxSize + m_lSpace) And y <= m_lStartY + 4 * m_lStartX + (pRows + 2) * (m_lBoxSize + m_lSpace))
  If IsMouseOnBut Then
    Call SetRect(R, m_lStartX, m_lStartY + 3 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace), pCols / 2 * (m_lBoxSize + m_lSpace), m_lStartY + 4 * m_lStartX + (pRows + 2) * (m_lBoxSize + m_lSpace))
    Call DrawRect(hdc, R, vbBlack, vbWhite, , True)
  End If
  
  'ColorDlg button
  IsMouseOnBut = (x >= m_lStartX + pCols / 2 * (m_lBoxSize + m_lSpace) And y >= m_lStartY + 3 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace)) And (x <= ScaleWidth - 1.5 * m_lStartX And y <= m_lStartY + 4 * m_lStartX + (pRows + 2) * (m_lBoxSize + m_lSpace))
  If IsMouseOnBut Then
    Call SetRect(R, m_lStartX + pCols / 2 * (m_lBoxSize + m_lSpace), m_lStartY + 3 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace), ScaleWidth - 1.5 * m_lStartX, m_lStartY + 4 * m_lStartX + (pRows + 2) * (m_lBoxSize + m_lSpace))
    Call DrawRect(hdc, R, vbBlack, vbWhite, , True)
  End If
  
  'draw 18 custom colors
  For j = 0 To 17
    IsMouseOnBut = (x >= m_lStartX + j * (m_lBoxSize + m_lSpace) And y >= m_lStartY + pRows * (m_lBoxSize + m_lSpace) And (x <= m_lStartX + j * (m_lBoxSize + m_lSpace) + m_lBoxSize And y <= m_lStartY + pRows * (m_lBoxSize + m_lSpace) + m_lBoxSize))
    If IsMouseOnBut Then
      Call SetRect(R, m_lStartX + j * (m_lBoxSize + m_lSpace), m_lStartY + pRows * (m_lBoxSize + m_lSpace) + m_lStartX, m_lStartX + j * (m_lBoxSize + m_lSpace) + m_lBoxSize, m_lStartY + pRows * (m_lBoxSize + m_lSpace) + m_lBoxSize + m_lStartX)
      Call DrawRect(hdc, R, vbBlack, vbWhite, VBClr(vbWhite))
      Exit For
    End If
  Next
  
  m_iClrIndex = 1
  For i = 0 To pRows - 1
    For j = 0 To pCols - 1
      If m_iClrIndex > m_lColorCount Then Exit Sub
      IsMouseOnBut = (x >= m_lStartX + j * (m_lBoxSize + m_lSpace) And y >= m_lStartY + i * (m_lBoxSize + m_lSpace)) And (x <= m_lStartX + j * (m_lBoxSize + m_lSpace) + 14 And y <= m_lStartY + i * (m_lBoxSize + m_lSpace) + 14)
      If IsMouseOnBut Then
        Call DrawButEdge(i, j, 2)
        Exit For
      End If
      m_iClrIndex = m_iClrIndex + 1
    Next j
    If IsMouseOnBut Then Exit For
  Next i
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Static iLastRow As Long, iLastCol As Long
  Static iLastCust As Long
  
  Dim i As Long, j As Long
  Dim IsMouseOnBut As Boolean
  Dim R As RECT
  Dim m_iClrIndex As Long
    
  IsMouseOver = ((x >= 0) And (y >= 0) And (x <= ScaleWidth) And (y <= ScaleHeight))
  If Not IsMouseOver Then Exit Sub
  
  'menu button
  IsMouseOnBut = (x >= m_lStartX + (m_lBoxSize + m_lSpace) * 17 And y >= m_lStartX) And (x <= ScaleWidth - m_lStartX And y <= m_lStartY - 2 * m_lStartX)
  If IsMouseOnBut Then
    Call SetRect(R, m_lStartX + (m_lBoxSize + m_lSpace) * 17, m_lStartX, ScaleWidth - m_lStartX, m_lStartY - 2 * m_lStartX)
    Call DrawRect(hdc, R, vbWhite, vbButtonText, , True)
  Else
    Call SetRect(R, m_lStartX + (m_lBoxSize + m_lSpace) * 17, m_lStartX, ScaleWidth - m_lStartX, m_lStartY - 2 * m_lStartX)
    Call DrawRect(hdc, R, vbButtonFace, vbButtonFace, , True)
  End If
  
  'draw 18 custom colors
  Call SetRect(R, m_lStartX + iLastCust * (m_lBoxSize + m_lSpace), m_lStartY + pRows * (m_lBoxSize + m_lSpace) + m_lStartX, m_lStartX + iLastCust * (m_lBoxSize + m_lSpace) + m_lBoxSize, m_lStartY + pRows * (m_lBoxSize + m_lSpace) + m_lBoxSize + m_lStartX)
  Call DrawRect(hdc, R, vbButtonText, vbButtonText, VBClr(vbWhite))
  For j = 0 To 17
    IsMouseOnBut = (x >= m_lStartX + j * (m_lBoxSize + m_lSpace) And y >= m_lStartY + pRows * (m_lBoxSize + m_lSpace) And (x <= m_lStartX + j * (m_lBoxSize + m_lSpace) + m_lBoxSize And y <= m_lStartY + pRows * (m_lBoxSize + m_lSpace) + m_lBoxSize))
    If IsMouseOnBut Then
      Call SetRect(R, m_lStartX + j * (m_lBoxSize + m_lSpace), m_lStartY + pRows * (m_lBoxSize + m_lSpace) + m_lStartX, m_lStartX + j * (m_lBoxSize + m_lSpace) + m_lBoxSize, m_lStartY + pRows * (m_lBoxSize + m_lSpace) + m_lBoxSize + m_lStartX)
      Call DrawRect(hdc, R, , , VBClr(vbWhite))
      iLastCust = j
      Exit For
    End If
  Next
  
  'default button
  IsMouseOnBut = (x >= m_lStartX And y >= m_lStartY + 3 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace)) And (x <= pCols / 2 * (m_lBoxSize + m_lSpace) And y <= m_lStartY + 4 * m_lStartX + (pRows + 2) * (m_lBoxSize + m_lSpace))
  If IsMouseOnBut Then
    Call SetRect(R, m_lStartX, m_lStartY + 3 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace), pCols / 2 * (m_lBoxSize + m_lSpace), m_lStartY + 4 * m_lStartX + (pRows + 2) * (m_lBoxSize + m_lSpace))
    Call DrawRect(hdc, R, vbWhite, vbButtonText, , True)
  
    'draw color sample
    Call SetRect(R, m_lStartX, m_lStartX, m_lStartX + (m_lBoxSize + m_lSpace) * 6, m_lStartY - 2 * m_lStartX)
    Call DrawRect(hdc, R, vbButtonText, vbWhite, VBClr(m_oColors(0)))
  
    'draw color name
    Call SetRect(R, m_lStartX + (m_lBoxSize + m_lSpace) * 6 + 4, m_lStartX, m_lStartX + (m_lBoxSize + m_lSpace) * 12, m_lStartY - 2 * m_lStartX)
    Call DrawRect(hdc, R, vbButtonFace, vbButtonFace, vbButtonFace)
    Call DrawText(hdc, "Default", Len("Default"), R, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
  
  Else
    Call SetRect(R, m_lStartX, m_lStartY + 3 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace), pCols / 2 * (m_lBoxSize + m_lSpace), m_lStartY + 4 * m_lStartX + (pRows + 2) * (m_lBoxSize + m_lSpace))
    Call DrawRect(hdc, R, vbButtonFace, vbButtonFace, , True)
  End If
  
  'ColorDlg button
  IsMouseOnBut = (x >= m_lStartX + pCols / 2 * (m_lBoxSize + m_lSpace) And y >= m_lStartY + 3 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace)) And (x <= ScaleWidth - 1.5 * m_lStartX And y <= m_lStartY + 4 * m_lStartX + (pRows + 2) * (m_lBoxSize + m_lSpace))
  If IsMouseOnBut Then
    Call SetRect(R, m_lStartX + pCols / 2 * (m_lBoxSize + m_lSpace), m_lStartY + 3 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace), ScaleWidth - 1.5 * m_lStartX, m_lStartY + 4 * m_lStartX + (pRows + 2) * (m_lBoxSize + m_lSpace))
    Call DrawRect(hdc, R, vbWhite, vbButtonText, , True)
  Else
    Call SetRect(R, m_lStartX + pCols / 2 * (m_lBoxSize + m_lSpace), m_lStartY + 3 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace), ScaleWidth - 1.5 * m_lStartX, m_lStartY + 4 * m_lStartX + (pRows + 2) * (m_lBoxSize + m_lSpace))
    Call DrawRect(hdc, R, vbButtonFace, vbButtonFace, , True)
  End If
  
  iLastCol = IIf(iLastCol > m_lColorCount, m_lColorCount - 1, iLastCol)
  iLastRow = IIf(iLastRow * pCols + iLastCol > m_lColorCount, 0, iLastRow)
  Call SetRect(R, m_lStartX + iLastCol * (m_lBoxSize + m_lSpace), m_lStartY + iLastRow * (m_lBoxSize + m_lSpace), m_lStartX + iLastCol * (m_lBoxSize + m_lSpace) + m_lBoxSize, m_lStartY + iLastRow * (m_lBoxSize + m_lSpace) + m_lBoxSize)
  Call DrawRect(hdc, R, vbButtonText, vbButtonText, VBClr(m_oColors(iLastRow * pCols + iLastCol)))
  m_iClrIndex = 1
  For i = 0 To pRows - 1
    For j = 0 To pCols - 1
      If m_iClrIndex > m_lColorCount Then Exit Sub
      IsMouseOnBut = (x >= m_lStartX + j * (m_lBoxSize + m_lSpace) And y >= m_lStartY + i * (m_lBoxSize + m_lSpace)) And (x <= m_lStartX + j * (m_lBoxSize + m_lSpace) + 14 And y <= m_lStartY + i * (m_lBoxSize + m_lSpace) + 14)
      If IsMouseOnBut Then
        'draw button were on
        Call SetRect(R, m_lStartX + j * (m_lBoxSize + m_lSpace), m_lStartY + i * (m_lBoxSize + m_lSpace), m_lStartX + j * (m_lBoxSize + m_lSpace) + m_lBoxSize, m_lStartY + i * (m_lBoxSize + m_lSpace) + m_lBoxSize)
        Call DrawRect(hdc, R, vbScrollBars, vbButtonShadow, VBClr(m_oColors(i * 18 + j)))
       
        'draw color sample
        Call SetRect(R, m_lStartX, m_lStartX, m_lStartX + (m_lBoxSize + m_lSpace) * 6, m_lStartY - 2 * m_lStartX)
        Call DrawRect(hdc, R, vbButtonText, vbWhite, VBClr(m_oColors(i * 18 + j)))
  
        'draw color name
        Call SetRect(R, m_lStartX + (m_lBoxSize + m_lSpace) * 6 + 4, m_lStartX, m_lStartX + (m_lBoxSize + m_lSpace) * 12, m_lStartY - 2 * m_lStartX)
        Call DrawRect(hdc, R, vbButtonFace, vbButtonFace, vbButtonFace)
        CodeColor = CStr(RGB2Hex(m_oColors(i * 18 + j)))
        Call DrawText(hdc, CStr(RGB2Hex(m_oColors(i * 18 + j))), Len(CStr(RGB2Hex(m_oColors(i * 18 + j)))), R, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
        
        iLastRow = i
        iLastCol = j
        Exit For
      End If
      m_iClrIndex = m_iClrIndex + 1
    Next j
    If IsMouseOnBut Then Exit For
  Next i
  
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Long, j As Long
  Dim IsMouseOnBut As Boolean
  Dim m_iClrIndex As Long
  
  If Not (Button = 1) Then Exit Sub
  IsMouseOver = x >= 0 And y >= 0 And x <= ScaleWidth And y <= ScaleHeight
    
  If IsMouseOver Then
     Call SetCapture(hwnd)
  Else
     Call ReleaseCapture
     Call Form_KeyDown(vbKeyEscape, 0)
     Exit Sub
  End If
  
  'default button
  IsMouseOnBut = (x >= m_lStartX And y >= m_lStartY + 3 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace)) And (x <= pCols / 2 * (m_lBoxSize + m_lSpace) And y <= m_lStartY + 4 * m_lStartX + (pRows + 2) * (m_lBoxSize + m_lSpace))
  If IsMouseOnBut Then
    SelectedColor = m_oColors(0)
    IsCanceled = False
    Call Form_KeyDown(vbKeyEscape, 0)
    Exit Sub
  End If
  
  'ColorDlg button
  IsMouseOnBut = (x >= m_lStartX + pCols / 2 * (m_lBoxSize + m_lSpace) And y >= m_lStartY + 3 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace)) And (x <= ScaleWidth - 1.5 * m_lStartX And y <= m_lStartY + 4 * m_lStartX + (pRows + 2) * (m_lBoxSize + m_lSpace))
  If IsMouseOnBut Then
    SelectedColor = ShowColor
    CodeColor = CStr(RGB2Hex(SelectedColor))
    Dim R As RECT
    Call DrawText(hdc, CodeColor, Len(CodeColor), R, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
    IsCanceled = False
    Call Form_KeyDown(vbKeyEscape, 0)
    Exit Sub
  End If
  
  IsMouseOnBut = (x >= m_lStartX + (m_lBoxSize + m_lSpace) * 17 And y >= m_lStartY - 27) And (x <= ScaleWidth - 2 And y <= m_lStartY - 4)
'  If IsMouseOnBut Then
'    Dim iMain As Long, iP As Long
'    Set m_cMenu = New cPopupMenu
'    With m_cMenu
'      .hWndOwner = hwnd
'      .MenuStyle = Office10
'      iMain = .AddItem("Menu", , , , , , , "PALMENU")
'      .AddItem "&2 Colors", , , iMain, , (m_lSelectedPallete = cc2), , "mnuPalMenu_2"
'      .AddItem "&8 Colors", , , iMain, , (m_lSelectedPallete = cc8), , "mnuPalMenu_8"
'      .AddItem "&16 Colors", , , iMain, , (m_lSelectedPallete = cc16), , "mnuPalMenu_16"
'      .AddItem "&Default colors", , , iMain, , (m_lSelectedPallete = ccVB), , "mnuPalMenu_VB"
'      .AddItem "2&56 Colors", , , iMain, , (m_lSelectedPallete = cc256), , "mnuPalMenu_256"
'      .AddItem "256 &Grays", , , iMain, , (m_lSelectedPallete = cc256Gray), , "mnuPalMenu_256Gray"
'      .AddItem "&Safe Colors", , , iMain, , (m_lSelectedPallete = ccSafe), , "mnuPalMenu_Safe"
'      .AddItem "&Netscape Colors", , , iMain, , (m_lSelectedPallete = ccNamed), , "mnuPalMenu_Named"
'      .AddItem "&Windows Colors", , , iMain, , (m_lSelectedPallete = ccWin), , "mnuPalMenu_Win"
'      '.AddItem "-", , , iMain
'      '.AddItem "Cargar &Paleta de Colores...", , , iMain, , , , "mnuPalMenu_Load"
'      .ShowPopupMenuObj Me, .IndexForKey("PALMENU") + 1, ScaleWidth - 2, m_lStartY - 27
'    End With
'    Set m_cMenu = Nothing
'  End If
  
  m_iClrIndex = 1
  For i = 0 To pRows - 1
    For j = 0 To pCols - 1
      If m_iClrIndex > m_lColorCount Then Exit Sub
      IsMouseOnBut = (x >= m_lStartX + j * (m_lBoxSize + m_lSpace) And y >= m_lStartY + i * (m_lBoxSize + m_lSpace)) And (x <= m_lStartX + j * (m_lBoxSize + m_lSpace) + 14 And y <= m_lStartY + i * (m_lBoxSize + m_lSpace) + 14)
      If IsMouseOnBut Then
        Call DrawButEdge(i, j, 0)
        SelectedColor = m_oColors(i * 18 + j)
        IsCanceled = False
        Call Form_KeyDown(vbKeyEscape, 0)
        Exit For
      End If
      m_iClrIndex = m_iClrIndex + 1
    Next j
    If IsMouseOnBut Then Exit For
  Next i
End Sub

Private Sub DrawButEdge(i As Long, j As Long, EdgeStyle As Integer)
  Dim hPen As Long, hPenOld As Long
  Dim tJunk As PointAPI, R As RECT
  
  Call SetRect(R, m_lStartX + j * (m_lBoxSize + m_lSpace), m_lStartY + i * (m_lBoxSize + m_lSpace), m_lStartX + j * (m_lBoxSize + m_lSpace) + m_lBoxSize, m_lStartY + i * (m_lBoxSize + m_lSpace) + m_lBoxSize)
  
  Select Case EdgeStyle
    Case 0
      hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbButtonShadow And &H1F&))
      hPenOld = SelectObject(hdc, hPen)
      MoveToEx hdc, R.Left, R.Top, tJunk
      LineTo hdc, R.Right, R.Top
      LineTo hdc, R.Right, R.Bottom
      LineTo hdc, R.Left, R.Bottom
      LineTo hdc, R.Left, R.Top
      Call DeleteObject(hPen)
      Call DeleteObject(hPenOld)
    Case 1
      hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbScrollBars And &H1F&))
      hPenOld = SelectObject(hdc, hPen)
      MoveToEx hdc, R.Left, R.Top, tJunk
      LineTo hdc, R.Right, R.Top
      MoveToEx hdc, R.Left, R.Top, tJunk
      LineTo hdc, R.Left, R.Bottom
      Call DeleteObject(hPen)
      Call DeleteObject(hPenOld)
      
      hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbButtonShadow And &H1F&))
      hPenOld = SelectObject(hdc, hPen)
      MoveToEx hdc, R.Right, R.Top, tJunk
      LineTo hdc, R.Right, R.Bottom
      LineTo hdc, R.Left, R.Bottom
      Call DeleteObject(hPen)
      Call DeleteObject(hPenOld)
    Case 2
      hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbButtonShadow And &H1F&))
      hPenOld = SelectObject(hdc, hPen)
      MoveToEx hdc, R.Left, R.Top, tJunk
      LineTo hdc, R.Right, R.Top
      MoveToEx hdc, R.Left, R.Top, tJunk
      LineTo hdc, R.Left, R.Bottom
      Call DeleteObject(hPen)
      Call DeleteObject(hPenOld)
      
      hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbScrollBars And &H1F&))
      hPenOld = SelectObject(hdc, hPen)
      MoveToEx hdc, R.Right, R.Top, tJunk
      LineTo hdc, R.Right, R.Bottom
      LineTo hdc, R.Left, R.Bottom
      Call DeleteObject(hPen)
      Call DeleteObject(hPenOld)
  End Select
End Sub

Private Sub DrawPalette(hdc As Long, Height As Long, Width As Long)
  Dim R As RECT
  Dim lClrIdx As Long
  Dim x As Long, y As Long
  
  Call SetRect(R, 0, 0, Width, Height)
  Call DrawRect(hdc, R, vbButtonFace, vbButtonFace, vbButtonFace)
  
  Dim OldFont As String, OldSize As Integer
  OldFont = Me.Font.Name
  OldSize = Me.Font.Size
  Me.Font.Name = "Marlett"
  Me.Font.Size = 10
  Call SetRect(R, m_lStartX + (m_lBoxSize + m_lSpace) * 17, m_lStartX, ScaleWidth - m_lStartX, m_lStartY - 2 * m_lStartX)
  Call DrawRect(hdc, R, vbButtonFace, vbButtonFace, , True)
  Call SetRect(R, m_lStartX + (m_lBoxSize + m_lSpace) * 17 + 1, m_lStartX, ScaleWidth - m_lStartX + 1, m_lStartY - 2 * m_lStartX)
  Call DrawText(hdc, "4", Len("6"), R, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
  Me.Font.Name = OldFont
  Me.Font.Size = OldSize
  
  'draw color sample
  Call SetRect(R, m_lStartX, m_lStartX, m_lStartX + (m_lBoxSize + m_lSpace) * 6, m_lStartY - 2 * m_lStartX)
  Call DrawRect(hdc, R, vbButtonText, vbWhite, VBClr(m_oColors(0)))
  
  'draw color name
  Call SetRect(R, m_lStartX + (m_lBoxSize + m_lSpace) * 6 + 4, m_lStartX, m_lStartX + (m_lBoxSize + m_lSpace) * 12, m_lStartY - 2 * m_lStartX)
  Call DrawRect(hdc, R, vbButtonFace, vbButtonFace, vbButtonFace)
  CodeColor = CStr(RGB2Hex(m_oColors(0)))
  Call DrawText(hdc, CStr(RGB2Hex(m_oColors(0))), Len(CStr(RGB2Hex(m_oColors(0)))), R, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
  
  'separator
  Call SetRect(R, 0, m_lStartY - m_lStartX, ScaleWidth, m_lStartY - m_lStartX)
  Call DrawRect(hdc, R, vbButtonText, vbButtonText, , True)
  
  For y = 0 To pRows - 1
    ' Get out of this loop if there's no more colors
    If lClrIdx > (m_lColorCount - 1) Then Exit For
    For x = 0 To pCols - 1
      ' Get out of this loop if there's no more colors
      If lClrIdx > (m_lColorCount - 1) Then Exit For
      Call SetRect(R, m_lStartX + x * (m_lBoxSize + m_lSpace), m_lStartY + y * (m_lBoxSize + m_lSpace), m_lStartX + x * (m_lBoxSize + m_lSpace) + m_lBoxSize, m_lStartY + y * (m_lBoxSize + m_lSpace) + m_lBoxSize)
      Call DrawRect(hdc, R, vbButtonText, vbButtonText, VBClr(m_oColors(lClrIdx)))
      lClrIdx = lClrIdx + 1
    Next
  Next
   
  'draw 18 custom colors
  For x = 0 To 17
    Call SetRect(R, m_lStartX + x * (m_lBoxSize + m_lSpace), m_lStartY + pRows * (m_lBoxSize + m_lSpace) + m_lStartX, m_lStartX + x * (m_lBoxSize + m_lSpace) + m_lBoxSize, m_lStartY + pRows * (m_lBoxSize + m_lSpace) + m_lBoxSize + m_lStartX)
    Call DrawRect(hdc, R, vbButtonText, vbButtonText, VBClr(vbWhite))
  Next
  
  'separator
  Call SetRect(R, 0, m_lStartY + 2 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace), ScaleWidth, m_lStartY + 2 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace))
  Call DrawRect(hdc, R, vbButtonText, vbButtonText, , True)

  'draw default color button
  Call SetRect(R, m_lStartX, m_lStartY + 3 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace), pCols / 2 * (m_lBoxSize + m_lSpace), m_lStartY + 4 * m_lStartX + (pRows + 2) * (m_lBoxSize + m_lSpace))
  Call DrawRect(hdc, R, vbButtonFace, vbButtonFace, vbButtonFace)
  Call DrawText(hdc, "Default", Len("Default"), R, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
  Call SetRect(R, m_lStartX * 2, m_lStartY + 4 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace), m_lStartX + (m_lBoxSize + m_lSpace), m_lStartY + 3 * m_lStartX + (pRows + 2) * (m_lBoxSize + m_lSpace))
  Call DrawRect(hdc, R, vbButtonText, vbButtonText, VBClr(m_oColors(0)))
    
  'draw ColorDialog button
  Call SetRect(R, m_lStartX + pCols / 2 * (m_lBoxSize + m_lSpace), m_lStartY + 3 * m_lStartX + (pRows + 1) * (m_lBoxSize + m_lSpace), ScaleWidth - 1.5 * m_lStartX, m_lStartY + 4 * m_lStartX + (pRows + 2) * (m_lBoxSize + m_lSpace))
  Call DrawRect(hdc, R, vbButtonFace, vbButtonFace, vbButtonFace)
  Call DrawText(hdc, "Windows Colors", Len("Windows Colors"), R, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
    
   'draw control border
  Call SetRect(R, 0, 0, ScaleWidth - 1, ScaleHeight - 1)
  Call DrawRect(hdc, R, vbWhite, vbButtonText, , True)
End Sub

Private Function ShowColor() As Long
    Dim ClrInf As udtCHOOSECOLOR
    Static CustomColors(64) As Byte
    Dim i As Integer
    
    For i = LBound(CustomColors) To UBound(CustomColors)
      CustomColors(i) = 0
    Next i
    
    With ClrInf
      .lStructSize = Len(ClrInf)              'Size of the structure
      .hWndOwner = Me.hwnd                    'Handle of owner window
      .hInstance = App.hInstance              'Instance of application
      .lpCustColors = StrConv(CustomColors, vbUnicode)       'Array of 16 byte values
      .flags = CC_FULLOPEN                    'Flags to open in full mode
    End With
    
    If Not ChooseColor(ClrInf) = 0 Then
        ShowColor = ClrInf.rgbResult
    Else
        ShowColor = -1
    End If
End Function

Private Function LoadCustClr(index As Long) As Long
  Dim clr As Long
  If index > 18 Or index < 1 Then
    clr = 1
  Else
    clr = m_oCustClrs(index)
  End If
  LoadCustClr = clr
End Function

Private Sub SaveCustClr(ClrVal As Long)
  If (LastSavedCustClr = 0) Then
    ReDim Preserve m_oCustClrs(1) As Long
  Else
    If (UBound(m_oCustClrs) < 18) Then
      ReDim Preserve m_oCustClrs(UBound(m_oCustClrs) + 1) As Long
    End If
  End If

  LastSavedCustClr = LastSavedCustClr + 1
  If (LastSavedCustClr > 18) Then LastSavedCustClr = 1

  m_oCustClrs(LastSavedCustClr) = ClrVal
End Sub

Private Sub Form_Paint()
  Call DrawPalette(Me.hdc, Height, Width)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmColorPalette = Nothing
  Erase m_oColors
End Sub

'Public Sub TipTimer(hwnd As Long, uMsg As Long, idEvent As Long, dwTime As Long)
'    Select Case idEvent
'        Case 1
'            Call ShowTip(True)
'
'            Call KillTimer(Me.hwnd, CLng(TipTmr1))
'            IsTmr1Active = False
'        Case 2
'            Call ShowTip(False)
'    End Select
'End Sub

'Private Sub ShowTip(State As Boolean)
'  If State Then
'    Dim Rct As RECT
'    Dim pTop As POINTAPI
'    Dim TipTxt As String
        
    'Store the tip text in a variable
'    TipTxt = m_oColors(MouseButId).Tip
'    If TipTxt = "" Then Exit Sub
        
    'Clear Tip Form
'    frmTip.Cls
        
    'Draw Tip text and position the Tip Form
'    Call GetCursorPos(pTop)
'    Call SetRect(Rct, 0, 0, frmTip.ScaleWidth, frmTip.ScaleHeight)
'    Call DrawText(frmTip.hdc, TipTxt, CLng(Len(TipTxt)), Rct, DT_CALCRECT)
'    Call SetRect(Rct, 0, 0, Rct.Right + 8, Rct.Bottom + 6)
'    Call DrawText(frmTip.hdc, TipTxt, CLng(Len(TipTxt)), Rct, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER)
'    Call DrawEdge(frmTip.hdc, Rct, BDR_RAISEDINNER, BF_RECT)
'    frmTip.Move (pTop.x + 2) * Screen.TwipsPerPixelX, (pTop.y + 20) * Screen.TwipsPerPixelY, _
'                    Rct.Right * Screen.TwipsPerPixelX, Rct.Bottom * Screen.TwipsPerPixelY
'    frmTip.ZOrder
'    frmTip.Refresh
'    Call ShowWindow(frmTip.hwnd, SW_SHOWNOACTIVATE)
        
    'Set Timer 2 for the duration of tip
'    Call SetTimer(Me.hwnd, CLng(TipTmr2), 4000, AddressOf Timer)
'    IsTmr2Active = True
'  Else
'    On Error Resume Next
    
    'Hide Tip Form
'    Call ShowWindow(frmTip.hwnd, SW_HIDE)
        
    'Kill Timer 2 if it is active
'    If IsTmr2Active Then
'      Call KillTimer(Me.hwnd, CLng(TipTmr2))
'      IsTmr2Active = False
'    End If
        
    'Kill Timer 1 if it is active
'    If IsTmr1Active Then
'      Call KillTimer(Me.hwnd, CLng(TipTmr1))
'      IsTmr1Active = False
'     End If
'  End If
'End Sub

Public Function RGB2Hex(lCdjor As Long) As String
    Dim j As Long
    Dim iRed, iGreen, iBlue As Integer
    Dim vHexR, vHexG, vHexB As Variant
    
    'Break out the R, G, B values from the common dialog color
    j = lCdjor
    iRed = j Mod &H100
    j = j \ &H100
    iGreen = j Mod &H100
    j = j \ &H100
    iBlue = j Mod &H100
    
    'Determine Red Hex
    vHexR = Hex(iRed)

    If Len(vHexR) < 2 Then
      vHexR = "0" & vHexR
    End If

    'Determine Green Hex
    vHexG = Hex(iGreen)
    If Len(vHexG) < 2 Then
      vHexG = "0" & iGreen
    End If

    'Determine Blue Hex
    vHexB = Hex(iBlue)
    If Len(vHexB) < 2 Then
      vHexB = "0" & vHexB
    End If
    'Add it up, return the function value
    RGB2Hex = "#" & vHexR & vHexG & vHexB
End Function

Private Sub m_cMenu_Click(ItemNumber As Long)
  
  On Error GoTo huboerror
    Dim sKey As String
  sKey = m_cMenu.ItemKey(ItemNumber)
  Select Case sKey
    Case "mnuPalMenu_2"
      m_lColorCount = LoadPalette(PathPaleta & "2.pal", m_oClrNames)
      m_lSelectedPallete = cc2
    Case "mnuPalMenu_8"
      m_lColorCount = LoadPalette(PathPaleta & "8.pal", m_oClrNames)
      m_lSelectedPallete = cc8
    Case "mnuPalMenu_16"
      m_lSelectedPallete = cc16
      m_lColorCount = LoadPalette(PathPaleta & "16.pal", m_oClrNames)
    Case "mnuPalMenu_256"
      m_lSelectedPallete = cc256
      m_lColorCount = LoadPalette(PathPaleta & "256c.pal", m_oClrNames)
    Case "mnuPalMenu_256Gray"
      m_lSelectedPallete = cc256Gray
      m_lColorCount = LoadPalette(PathPaleta & "256g.pal", m_oClrNames)
    Case "mnuPalMenu_VB"
      m_lSelectedPallete = ccVB
      m_lColorCount = LoadPalette(PathPaleta & "default.pal", m_oClrNames)
    Case "mnuPalMenu_Safe"
      m_lSelectedPallete = ccSafe
      m_lColorCount = LoadPalette(PathPaleta & "browser.pal", m_oClrNames)
    Case "mnuPalMenu_Named"
      m_lSelectedPallete = ccNamed
      m_lColorCount = LoadPalette(PathPaleta & "named.pal", m_oClrNames)
    Case "mnuPalMenu_Win"
      m_lSelectedPallete = ccWin
      m_lColorCount = LoadPalette(PathPaleta & "windows.pal", m_oClrNames)
    'Case "mnuPalMenu_Load"
    '  Dim sFileName As String
    '  Dim cc As New cCommonDialog
    '  If cc.VBGetOpenFileName(sFileName, , , , , , "Palettes|*.pal;*.hpl", 1, PathPaleta, "Seleccionar paleta de colores ...", "pal", Me.hwnd) Then
    '    m_sLastPal = sFileName
    '    m_lColorCount = LoadPalette(m_sLastPal, m_oClrNames)
    '    m_lSelectedPallete = ccCustom
    '  End If
    '  Set cc = Nothing
  End Select
  Me.Width = (m_lStartX * 2 + 18 * (m_lBoxSize + m_lSpace)) * Screen.TwipsPerPixelX
  Me.Height = (m_lStartY + (pRows + 2) * (m_lBoxSize + m_lSpace) + 5 * m_lStartX) * Screen.TwipsPerPixelY
  Call Form_Paint
  Call SetCapture(hwnd)
  
  Exit Sub
  
huboerror:
  MsgBox "error en carga de paletas : " & Err & " " & Error$, vbCritical
  Err = 0
  
End Sub
