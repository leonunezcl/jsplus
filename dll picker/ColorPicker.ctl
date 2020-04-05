VERSION 5.00
Begin VB.UserControl ColorPicker 
   AutoRedraw      =   -1  'True
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   35
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   137
   ToolboxBitmap   =   "ColorPicker.ctx":0000
End
Attribute VB_Name = "ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Private RClr As RECT
Private RBut As RECT

Private IsInFocus As Boolean
Private IsButDown As Boolean

'Default Property Values:
Private Const m_def_Color = &HFFFFFF
Private Const m_def_BoxSize = 14
Private Const m_def_Spacing = 0

'Property Variables:
Private m_Color                 As OLE_COLOR
Private m_BoxSize               As Integer
Private m_Spacing               As Integer
Private m_Code                  As String
Private m_PathPaleta            As String

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Resize()
Event ColorSelected(m_Color As OLE_COLOR, m_Code As String)
Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

Private Sub UserControl_GotFocus()
    IsInFocus = True
    Call RedrawControl(m_Color)
End Sub

Private Sub UserControl_Initialize()
  ScaleMode = vbPixels
  Call UserControl_InitProperties
End Sub

Private Sub UserControl_LostFocus()
  IsInFocus = False
  Call RedrawControl(m_Color)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x * Screen.TwipsPerPixelX, y * Screen.TwipsPerPixelY)
    
  If Button = 1 Then
    If (x >= ScaleLeft And x <= ScaleWidth) And (y >= ScaleTop And y <= ScaleHeight) Then
      IsButDown = True
      Call RedrawControl(m_Color)
    End If
  End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x * Screen.TwipsPerPixelX, y * Screen.TwipsPerPixelY)
    
  If IsButDown Then
    If Not ((x >= ScaleLeft And x <= ScaleWidth) And (y >= ScaleTop And y <= ScaleHeight)) Then
      IsButDown = False
      Call RedrawControl(m_Color)
    End If
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x * Screen.TwipsPerPixelX, y * Screen.TwipsPerPixelY)
    
  If Button = 1 Then
    If IsButDown Then
      IsButDown = False
      Call RedrawControl(m_Color)
    End If
        
    If ((x >= ScaleLeft And x <= ScaleWidth) And (y >= ScaleTop And y <= ScaleHeight)) Then
      Call ShowPalette
    End If
  End If
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
    If Height < 285 Then Height = 285
    
    Call RedrawControl(m_Color)
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub RedrawControl(lColor As Long)
  Dim rct As RECT
  Dim Brsh As Long, Color As Long
  Dim tJunk As PointAPI
  Dim hPen As Long
  Dim hPenOld As Long
    
  Dim x1 As Long, y1 As Long
  Dim x2 As Long, y2 As Long
    
  x1 = ScaleLeft
  y1 = ScaleTop
  x2 = ScaleWidth
  y2 = ScaleHeight
    
  Cls
    
  'Draw background
  If Not IsButDown Then
    hPen = CreatePen(PS_SOLID, 1, vbWhite) ' GetSysColor(vbWhite And &H1F&))
    hPenOld = SelectObject(hdc, hPen)
    
    Call MoveToEx(hdc, x1, y1, tJunk)
    Call LineTo(hdc, x2 - 1, y1)
    Call MoveToEx(hdc, x1, y1, tJunk)
    Call LineTo(hdc, x1, y2 - 1)
    Call DeleteObject(hPen)
    Call DeleteObject(hPenOld)
    
    hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbButtonText And &H1F&))
    hPenOld = SelectObject(hdc, hPen)
    
    Call MoveToEx(hdc, x2 - 1, y1, tJunk)
    Call LineTo(hdc, x2 - 1, y2 - 1)
    Call LineTo(hdc, x1, y2 - 1)
    
    Call DeleteObject(hPen)
    Call DeleteObject(hPenOld)
  End If
  
  'Draw button
  Dim CurFontName As String
  CurFontName = Font.Name
  Font.Name = "Marlett"
  Call OleTranslateColor(vbButtonFace, ByVal 0&, Color)
  Brsh = CreateSolidBrush(Color)
  If IsButDown Then
    Call SetRect(RBut, x2 - 10, y2 - 10, x2 - 2, y2 - 2)
    Call SetRect(rct, RBut.Left + 2, RBut.Top, RBut.Right, RBut.Bottom)
    Call DrawText(hdc, "6", 1&, rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
  Else
    Call SetRect(RBut, x2 - 10, y2 - 10, x2 - 2, y2 - 2)
    Call SetRect(rct, RBut.Left + 1, RBut.Top, RBut.Right, RBut.Bottom - 1)
    Call DrawText(hdc, "6", 1&, rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
  End If
  Font.Name = CurFontName
  Call DeleteObject(Brsh)
  Call DeleteObject(Color)
  
  'Draw Color
  Call OleTranslateColor(lColor, ByVal 0&, Color)
  Brsh = CreateSolidBrush(Color)
  Call SetRect(RClr, 2, 2, x2 - 3, y2 - 10)
  Call FillRect(hdc, RClr, Brsh)
  Call SetRect(RClr, 2, 2, x2 - 10, y2 - 3)
  Call FillRect(hdc, RClr, Brsh)
  Call DeleteObject(Brsh)
  Call DeleteObject(Color)
  
  hPen = CreatePen(PS_SOLID, 1, GetSysColor(vbButtonText And &H1F&))
  hPenOld = SelectObject(hdc, hPen)
  
  Call MoveToEx(hdc, 2, 2, tJunk)
  Call LineTo(hdc, x2 - 3, 2)
  Call MoveToEx(hdc, 2, 2, tJunk)
  Call LineTo(hdc, 2, y2 - 3)
  
  Call DeleteObject(hPen)
  Call DeleteObject(hPenOld)
  
  hPen = CreatePen(PS_SOLID, 1, vbWhite) 'GetSysColor(vbScrollBars And &H1F&))
  hPenOld = SelectObject(hdc, hPen)
    
  Call MoveToEx(hdc, x2 - 3, 2, tJunk)
  Call LineTo(hdc, x2 - 3, y2 - 10)
  Call LineTo(hdc, x2 - 10, y2 - 10)
  Call LineTo(hdc, x2 - 10, y2 - 3)
  Call LineTo(hdc, 2, y2 - 3)
    
  Call DeleteObject(hPen)
  Call DeleteObject(hPenOld)
End Sub

Public Sub ShowPalette()
  Dim ClrCtrlPos As RECT
    
  If m_PathPaleta = "" Then
    MsgBox "Falta definir el path de las paletas graficas"
    Exit Sub
  End If
  
  Call GetWindowRect(hwnd, ClrCtrlPos)
    
  m_lDefault = m_Color
  frmColorPalette.PathPaleta = m_PathPaleta
  Load frmColorPalette
  With frmColorPalette
    .Left = ClrCtrlPos.Left * Screen.TwipsPerPixelX
    .Top = (ClrCtrlPos.Bottom) * Screen.TwipsPerPixelY
    If (.Top + .Height) > Screen.Height Then
      .Top = ClrCtrlPos.Top * Screen.TwipsPerPixelY - .Height
    End If
        
    .Show vbModal
        
    If Not .IsCanceled Then
        m_Color = .SelectedColor
        m_Code = .CodeColor
        RaiseEvent ColorSelected(m_Color, m_Code)
    End If
        
    Call RedrawControl(m_Color)
  End With
  Unload frmColorPalette
End Sub

Private Sub UserControl_InitProperties()
  m_BoxSize = m_def_BoxSize
  m_lBoxSize = m_BoxSize
  m_Spacing = m_def_Spacing
  m_lSpace = m_Spacing
  m_Color = m_def_Color
  m_Code = ""
  Height = 315
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_BoxSize = PropBag.ReadProperty("BoxSize", m_def_BoxSize)
  m_Spacing = PropBag.ReadProperty("BoxSize", m_def_Spacing)
  m_Color = PropBag.ReadProperty("Color", m_def_Color)
  m_Code = PropBag.ReadProperty("Color", "")
  Call RedrawControl(m_Color)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BoxSize", m_BoxSize, m_def_BoxSize)
  Call PropBag.WriteProperty("Spacing", m_Spacing, m_def_Spacing)
  Call PropBag.WriteProperty("Color", m_Color, m_def_Color)
  Call PropBag.WriteProperty("Code", m_Code, "")
End Sub

Public Property Get Color() As OLE_COLOR
Attribute Color.VB_Description = "Returns/Sets the selected color"
Attribute Color.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Color.VB_UserMemId = 0
  Color = m_Color
End Property

Public Property Let Color(ByVal New_Color As OLE_COLOR)
  m_Color = New_Color
  PropertyChanged "Value"
    
  'Call RedrawControl(m_defColor)
End Property

Public Property Get BoxSize() As Integer
  BoxSize = m_BoxSize
End Property

Public Property Let BoxSize(ByVal New_BoxSize As Integer)
  m_BoxSize = New_BoxSize
  m_lBoxSize = m_BoxSize
  PropertyChanged "BoxSize"
End Property

Public Property Get Spacing() As Integer
  Spacing = m_Spacing
End Property

Public Property Let Spacing(ByVal New_Spacing As Integer)
  m_Spacing = New_Spacing
  m_lSpace = m_Spacing
  PropertyChanged "Spacing"
End Property


Public Property Get code() As String
    code = m_Code
End Property

Private Property Let code(ByVal pCode As String)
    m_Code = pCode
End Property

Public Property Get PathPaleta() As String
    PathPaleta = m_PathPaleta
End Property

Public Property Let PathPaleta(ByVal pPathPaleta As String)
    m_PathPaleta = pPathPaleta
End Property
