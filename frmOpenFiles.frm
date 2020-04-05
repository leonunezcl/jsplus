VERSION 5.00
Object = "{866F095F-113F-4DC1-B803-F4CF4AFC96EE}#1.0#0"; "vbspgbbar.ocx"
Begin VB.Form frmOpenFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opening Files .... Please Wait ..."
   ClientHeight    =   1440
   ClientLeft      =   930
   ClientTop       =   5505
   ClientWidth     =   4755
   ControlBox      =   0   'False
   Icon            =   "frmOpenFiles.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin vbsprgbar.ucProgressBar pgb 
      Height          =   255
      Left            =   45
      TabIndex        =   1
      Top             =   1095
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   12937777
   End
   Begin VB.Label lblfile 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   945
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   4530
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmOpenFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cancelo As Boolean

Private Sub cmd_Click()
    Cancelo = True
End Sub


Private Sub Form_Activate()

    windowontop hwnd
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOpenFiles = Nothing
End Sub


