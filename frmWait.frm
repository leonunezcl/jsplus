VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opening Document ...."
   ClientHeight    =   1260
   ClientLeft      =   4530
   ClientTop       =   1860
   ClientWidth     =   4515
   ControlBox      =   0   'False
   Icon            =   "frmWait.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picWait 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   30
      ScaleHeight     =   1125
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   45
      Width           =   4485
      Begin VB.Label lblaccion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Executing ..."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   825
         Width           =   4320
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Executing ..."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   45
         TabIndex        =   2
         Top             =   375
         Width           =   885
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while JavaScript Plus! load selected file."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   1
         Top             =   75
         Width           =   3675
      End
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    windowontop hwnd
    Refresh
    DoEvents
End Sub

Private Sub Form_Load()

    util.CenterForm Me
    
End Sub


Private Sub Form_Terminate()
    Set frmWait = Nothing
End Sub


