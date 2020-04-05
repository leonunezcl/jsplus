VERSION 5.00
Begin VB.Form frmUpload 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1650
   ClientLeft      =   4155
   ClientTop       =   3540
   ClientWidth     =   4260
   ControlBox      =   0   'False
   Icon            =   "frmUpload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblfile 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Please wait while file is uploaded....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Image imgup 
      Height          =   540
      Left            =   2775
      Picture         =   "frmUpload.frx":1042
      Top             =   240
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Image imgdw 
      Height          =   540
      Left            =   1290
      Picture         =   "frmUpload.frx":191B
      Top             =   195
      Width           =   1470
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait while file is uploaded....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   0
      Top             =   1320
      Width           =   4170
   End
End
Attribute VB_Name = "frmUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tipo_accion As Integer
Public file As String
Private Sub Form_Activate()
    windowontop hwnd
    Refresh
    DoEvents
End Sub

Private Sub Form_Load()
  
    On Error Resume Next
    
    util.CenterForm Me
       
    If tipo_accion = 0 Then
        Me.Caption = "Download File"
        lblmsg.Caption = "Please wait while file is downloaded ..."
    Else
        Me.Caption = "Upload File"
        lblmsg.Caption = "Please wait while file is uploaded ..."
        imgdw.Visible = False
        imgup.Move imgdw.Left, imgdw.Top
        imgup.Visible = True
        imgup.ZOrder 0
    End If
    
    Err = 0
End Sub

Private Sub Form_Terminate()
'    Set frmUpload = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmUpload = Nothing
End Sub
