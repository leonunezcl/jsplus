VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type PointAPI
   X As Long
   Y As Long
End Type

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move
Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Private Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Private Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Private Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Private Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Public Property Get X() As Long
Dim tp As PointAPI
   GetCursorPos tp
   X = tp.X
End Property
Public Property Get Y() As Long
Dim tp As PointAPI
   GetCursorPos tp
   Y = tp.Y
End Property
Public Property Let X(ByVal X As Long)
   MoveTo X, Y ' y from property get
End Property
Public Property Let Y(ByVal Y As Long)
   MoveTo X, Y ' x from property get
End Property

Public Sub MoveTo(ByVal X As Long, ByVal Y As Long)
Dim xl As Double
Dim yl As Double
Dim xMax As Long
Dim yMax As Long
   
   ' mouse_event ABSOLUTE coords run from 0 to 65535:
   xMax = Screen.Width \ Screen.TwipsPerPixelX
   yMax = Screen.Height \ Screen.TwipsPerPixelY
   xl = X * 65535 / xMax
   yl = Y * 65535 / yMax
   ' Move the mouse:
   mouse_event MOUSEEVENTF_MOVE Or MOUSEEVENTF_ABSOLUTE, xl, yl, 0, 0
   
End Sub

Public Sub Click(Optional ByVal eButton As MouseButtonConstants = vbLeftButton)
Dim lFlagDown As Long
Dim lFlagUp As Long
   Select Case eButton
   Case vbRightButton
      lFlagDown = MOUSEEVENTF_RIGHTDOWN
      lFlagUp = MOUSEEVENTF_RIGHTUP
   Case vbMiddleButton
      lFlagDown = MOUSEEVENTF_MIDDLEDOWN
      lFlagUp = MOUSEEVENTF_MIDDLEUP
   Case Else
      lFlagDown = MOUSEEVENTF_LEFTDOWN
      lFlagUp = MOUSEEVENTF_LEFTUP
   End Select
   ' A click = down then up
   mouse_event lFlagDown Or MOUSEEVENTF_ABSOLUTE, 0, 0, 0, 0
   mouse_event lFlagUp Or MOUSEEVENTF_ABSOLUTE, 0, 0, 0, 0
End Sub

