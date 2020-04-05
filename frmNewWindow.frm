VERSION 5.00
Begin VB.Form frmNewWindow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Window"
   ClientHeight    =   1695
   ClientLeft      =   6180
   ClientTop       =   4575
   ClientWidth     =   3945
   Icon            =   "frmNewWindow.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select ..."
      Height          =   960
      Left            =   45
      TabIndex        =   0
      Top             =   105
      Width           =   3840
      Begin VB.OptionButton opt 
         Caption         =   "Window Popup"
         Height          =   240
         Index           =   3
         Left            =   2100
         TabIndex        =   4
         Top             =   510
         Width           =   1515
      End
      Begin VB.OptionButton opt 
         Caption         =   "Window Prompt"
         Height          =   255
         Index           =   2
         Left            =   2100
         TabIndex        =   3
         Top             =   270
         Width           =   1515
      End
      Begin VB.OptionButton opt 
         Caption         =   "Window Confirm"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   510
         Width           =   1515
      End
      Begin VB.OptionButton opt 
         Caption         =   "Window Alert"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Value           =   -1  'True
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmNewWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If opt(0).Value Then
            If frmMain.ActiveForm.Name = "frmEdit" Then
                frmMain.CreateAlert
            End If
        ElseIf opt(1).Value Then
            If frmMain.ActiveForm.Name = "frmEdit" Then
                frmMain.CreateConfirm
            End If
        ElseIf opt(2).Value Then
            If frmMain.ActiveForm.Name = "frmEdit" Then
                frmMain.CreatePrompt
            End If
        Else
            If frmMain.ActiveForm.Name = "frmEdit" Then
                frmPopup.Show vbModal
            End If
        End If
    End If
    Unload Me
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    util.CenterForm Me
    
    Debug.Print "load"
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmNewWindow = Nothing
End Sub


