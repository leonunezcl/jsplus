VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   5025
   ClientTop       =   8625
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   7200
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   840
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Img As cVBALImageList

Private Sub Command1_Click()
    Picture1.Picture = LoadResPicture(200, vbResIcon)
End Sub

Private Sub Form_Load()

    Set m_Img = New cVBALImageList
    
    With m_Img
        .IconSizeX = 16: .IconSizeY = 16: .ColourDepth = ILC_COLOR24
        .Create
        .AddFromResourceID 241, App.hInstance, IMAGE_ICON, "k1"
        .AddFromResourceID 239, App.hInstance, IMAGE_ICON, "k2"
        .AddFromResourceID 228, App.hInstance, IMAGE_ICON, "k3"
        .AddFromResourceID 114, App.hInstance, IMAGE_ICON, "k4"
        .AddFromResourceID 282, App.hInstance, IMAGE_ICON, "k5"
        .AddFromResourceID 175, App.hInstance, IMAGE_ICON, "k6"
        .AddFromResourceID 108, App.hInstance, IMAGE_ICON, "k7"
        .AddFromResourceID 109, App.hInstance, IMAGE_ICON, "k8"
        .AddFromResourceID 101, App.hInstance, IMAGE_ICON, "k9"
        .AddFromResourceID 102, App.hInstance, IMAGE_ICON, "k10"
        .AddFromResourceID 103, App.hInstance, IMAGE_ICON, "k11"
        .AddFromResourceID 104, App.hInstance, IMAGE_ICON, "k12"
        .AddFromResourceID 107, App.hInstance, IMAGE_ICON, "k13"
        .AddFromResourceID 205, App.hInstance, IMAGE_ICON, "k14"
        .AddFromResourceID 206, App.hInstance, IMAGE_ICON, "k15"
        .AddFromResourceID 207, App.hInstance, IMAGE_ICON, "k16"
        .AddFromResourceID 208, App.hInstance, IMAGE_ICON, "k17"
    End With
    
End Sub


