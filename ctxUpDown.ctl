VERSION 5.00
Begin VB.UserControl ctxUpDown 
   CanGetFocus     =   0   'False
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   192
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   ToolboxBitmap   =   "ctxUpDown.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Timer Timer1 
      Left            =   3060
      Top             =   1740
   End
End
Attribute VB_Name = "ctxUpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
'   You are free to use this source as long as this copyright message
'     appears on your program's "About" dialog:
'
'   Outlook Bar Project
'   Copyright (c) 2002 Vlad Vissoultchev (wqweto@myrealbox.com)
'
'=========================================================================
Option Explicit
Private Const MODULE_NAME As String = "ctxUpDown"

'=========================================================================
' Events
'=========================================================================

Event Change(ByVal lValue As Long)

'=========================================================================
' API
'=========================================================================

'--- for mouse_event
Private Const MOUSEEVENTF_LEFTDOWN      As Long = &H2
'--- for DrawFrameControl
Private Const DFCS_FLAT                 As Long = &H4000
Private Const DFCS_PUSHED               As Long = &H200
Private Const DFCS_SCROLLDOWN           As Long = &H1
Private Const DFCS_SCROLLUP             As Long = &H0
Private Const DFC_SCROLL                As Long = 3

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_bPressed          As Boolean
Private m_lButton           As Long

'=========================================================================
' Control events
'=========================================================================

Private Sub Timer1_Timer()
    Timer1.Interval = 50
    If m_lButton = 1 Then
        If m_bPressed Then
            RaiseEvent Change(1)
        End If
    ElseIf m_lButton = 2 Then
        If m_bPressed Then
            RaiseEvent Change(-1)
        End If
    Else
        Timer1.Enabled = False
    End If
End Sub

Private Sub UserControl_DblClick()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Y >= 0 And Y < ScaleHeight \ 2 And x >= 0 And x < ScaleWidth Then
        m_lButton = 1
    ElseIf Y >= ScaleHeight \ 2 And Y < ScaleHeight Then
        m_lButton = 2
    End If
    m_bPressed = True
    Timer1.Interval = 500
    Timer1.Enabled = False
    Timer1.Enabled = True
    Refresh
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim bState      As Boolean
    If m_lButton = 1 Then
        bState = Y >= 0 And Y < ScaleHeight \ 2 And x >= 0 And x < ScaleWidth
    End If
    If m_lButton = 2 Then
        bState = Y >= ScaleHeight \ 2 And Y < ScaleHeight And x >= 0 And x < ScaleWidth
    End If
    If m_bPressed <> bState Then
        m_bPressed = bState
        Refresh
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Y >= 0 And Y < ScaleHeight \ 2 And x >= 0 And x < ScaleWidth Then
        If m_lButton = 1 Then
            RaiseEvent Change(1)
        End If
    End If
    If Y >= ScaleHeight \ 2 And Y < ScaleHeight And x >= 0 And x < ScaleWidth Then
        If m_lButton = 2 Then
            RaiseEvent Change(-1)
        End If
    End If
    m_lButton = 0
    m_bPressed = False
    Refresh
End Sub

Private Sub UserControl_Paint()
    Dim rc As RECT
    rc.Right = ScaleWidth
    rc.Bottom = ScaleHeight \ 2
    DrawFrameControl hDC, rc, DFC_SCROLL, DFCS_SCROLLUP Or IIf(m_bPressed And m_lButton = 1, DFCS_FLAT Or DFCS_PUSHED, 0)
    rc.Top = rc.Bottom
    rc.Bottom = rc.Top + ScaleHeight \ 2
    DrawFrameControl hDC, rc, DFC_SCROLL, DFCS_SCROLLDOWN Or IIf(m_bPressed And m_lButton = 2, DFCS_FLAT Or DFCS_PUSHED, 0)
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub
