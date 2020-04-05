VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   2265
   ClientTop       =   1605
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   7200
   Begin jsplus.ColPicker ColPicker1 
      Height          =   7650
      Left            =   1245
      TabIndex        =   0
      Top             =   0
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   13494
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ColPicker1.Load
End Sub


