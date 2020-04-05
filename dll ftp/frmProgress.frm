VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1485
   ClientLeft      =   1320
   ClientTop       =   1695
   ClientWidth     =   3360
   ControlBox      =   0   'False
   ForeColor       =   &H00404040&
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   99
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   224
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picFiles 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   360
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.PictureBox picFile 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   360
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblTimeLeft 
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo estimado : 0s @ 0Kbps"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3195
   End
   Begin VB.Image imgClose 
      Height          =   105
      Left            =   3255
      Picture         =   "frmProgress.frx":000C
      Top             =   15
      Width           =   105
   End
   Begin VB.Image imgDownload 
      Height          =   240
      Left            =   60
      Picture         =   "frmProgress.frx":00F6
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgUpload 
      Height          =   240
      Left            =   60
      Picture         =   "frmProgress.frx":2898
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblFilesProg 
      Alignment       =   1  'Right Justify
      Caption         =   "0%"
      Height          =   195
      Left            =   2925
      TabIndex        =   5
      Top             =   885
      Width           =   390
   End
   Begin VB.Label lblFileProg 
      Alignment       =   1  'Right Justify
      Caption         =   "0%"
      Height          =   195
      Left            =   2925
      TabIndex        =   4
      Top             =   510
      Width           =   390
   End
   Begin VB.Image Image2 
      Height          =   150
      Left            =   120
      Picture         =   "frmProgress.frx":503A
      Top             =   525
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   195
      Left            =   90
      Picture         =   "frmProgress.frx":5194
      Top             =   855
      Width           =   225
   End
   Begin VB.Label lblFileCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(0/0)"
      Height          =   195
      Left            =   2400
      TabIndex        =   1
      Top             =   150
      Width           =   945
   End
   Begin VB.Label lblCurrentFile 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1845
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'WinFTP, created by the KPD-Team 2000
'This file can be downloaded from http://www.allapi.net/
'For questions or comments, contact us at KPDTeam@Allapi.net

' You are free to use this code within your own applications,
' but you are expressly forbidden from selling or otherwise
' distributing this source code without prior written consent.
' This includes both posting free demo projects made from this
' code as well as reproducing the code in text or html format.

Const DC_ACTIVE = &H1
Const DC_ICON = &H4
Const DC_TEXT = &H8
Const BDR_RAISEDINNER = &H4
Const BDR_RAISEDOUTER = &H1
Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Const BF_BOTTOM = &H8
Const BF_LEFT = &H1
Const BF_RIGHT = &H4
Const BF_TOP = &H2
Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function DrawEdge Lib "User32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetRect Lib "User32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Sub UpdateProgress()
    If Me.Visible = False Then Exit Sub
    Dim R As RECT, T As Long
    lblCurrentFile.Caption = ActiveFile
    If foItems = 0 Then
        lblFileCount.Caption = "(0/0)"
    Else
        lblFileCount.Caption = "(" + CStr(ActiveIndex) + "/" + CStr(foItems) + ")"
    End If
    picFile.Line (0, 0)-(picFile.ScaleWidth, picFile.ScaleHeight), &H8000000F, BF
    picFiles.Line (0, 0)-(picFiles.ScaleWidth, picFiles.ScaleHeight), &H8000000F, BF
    If ActiveFileBytesTotal <> 0 Then
        lblFileProg.Caption = CStr(Int(ActiveFileBytesSent / ActiveFileBytesTotal * 100)) + "%"
        SetRect R, 0, 0, ActiveFileBytesSent / ActiveFileBytesTotal * picFile.ScaleWidth, picFile.ScaleHeight
        DrawEdge picFile.hdc, R, EDGE_RAISED, BF_RECT
    Else
        lblFileProg.Caption = ""
    End If
    If TotalFileSize <> 0 Then
        lblFilesProg.Caption = CStr(Int(SentBytes / TotalFileSize * 100)) + "%"
        SetRect R, 0, 0, SentBytes / TotalFileSize * picFiles.ScaleWidth, picFiles.ScaleHeight
        DrawEdge picFiles.hdc, R, EDGE_RAISED, BF_RECT
    Else
        lblFilesProg.Caption = ""
    End If
    If SentBytes <> 0 Then
        T = GetTickCount - StartT
        If T <> 0 Then
            OldSpeed = (OldSpeed + ((SentBytes / 1000) / (T / 1000))) / 2
            lblTimeLeft.Caption = "Estimated time left: " + CStr(Int(((TotalFileSize - SentBytes) / 1000) / OldSpeed)) + "s @ " + Format(OldSpeed, "#.##") + "Kbps"
            
        End If
    End If
    If ActiveProcedure = FOP_UPLOAD Then
        imgUpload.Visible = True
        imgDownload.Visible = False
    ElseIf ActiveProcedure = FOP_DOWNLOAD Then
        imgUpload.Visible = False
        imgDownload.Visible = True
    Else
        imgUpload.Visible = False
        imgDownload.Visible = False
    End If
    picFile.Refresh
    picFiles.Refresh
    DoEvents
End Sub
Private Sub Form_Activate()
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long
    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub
Private Sub imgClose_Click()
    Me.Visible = False
End Sub

