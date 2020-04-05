VERSION 5.00
Object = "*\AClrPckr.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   4395
   ClientTop       =   4965
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin ClrPckr.ColorPicker ColorPicker1 
      Height          =   285
      Left            =   1035
      TabIndex        =   0
      Top             =   360
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   503
      BoxSize         =   20
      Spacing         =   20
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ColorPicker1_ColorSelected(m_Color As stdole.OLE_COLOR, m_Code As String)
    Picture1.BackColor = m_Color
    Label1.Caption = m_Code
End Sub


Private Sub Form_Load()
  Me.ColorPicker1.BoxSize = 13
  Me.ColorPicker1.Spacing = 0
End Sub
