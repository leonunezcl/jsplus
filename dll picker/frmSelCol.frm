VERSION 5.00
Begin VB.Form frmSelCol 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Color"
   ClientHeight    =   1110
   ClientLeft      =   4275
   ClientTop       =   3960
   ClientWidth     =   3270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   Begin JSPadColorPicker.ColorPicker colpick 
      Height          =   315
      Left            =   135
      TabIndex        =   2
      Top             =   135
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Exit"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmSelCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PathPaleta As String
Private showfirst As Boolean
Private Sub cmd_Click(index As Integer)
    
    If index = 0 Then
        gCodeColor = colpick.code
        gSelectColor = colpick.Color
    Else
        gCodeColor = ""
        gSelectColor = 0
    End If
    Unload Me
    
End Sub

Private Sub Form_Activate()

    If Not showfirst Then
        showfirst = True
        colpick.ShowPalette
    End If
    
End Sub

Private Sub Form_Load()
    colpick.PathPaleta = PathPaleta
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmSelCol = Nothing
End Sub


