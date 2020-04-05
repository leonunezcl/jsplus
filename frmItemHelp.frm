VERSION 5.00
Begin VB.Form frmItemHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quick Help"
   ClientHeight    =   2970
   ClientLeft      =   4035
   ClientTop       =   2640
   ClientWidth     =   5700
   Icon            =   "frmItemHelp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtHelp 
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   1515
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmItemHelp.frx":014A
      Top             =   855
      Width           =   5490
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Member"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   5
      Top             =   360
      Width           =   780
   End
   Begin VB.Label lblMember 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Member"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   105
      Width           =   780
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Member"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   105
      Width           =   780
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   615
      Width           =   435
   End
End
Attribute VB_Name = "frmItemHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public member As String
Public mtype As String
Public mhelp As String
Private Sub cmd_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    util.CenterForm Me
    
    lblMember.Caption = member
    If mtype = "1" Then
        lblType.Caption = "Property"
    ElseIf mtype = "2" Then
        lblType.Caption = "Method"
    ElseIf mtype = "3" Then
        lblType.Caption = "Event"
    ElseIf mtype = "4" Then
        lblType.Caption = "Constant"
    End If
    txtHelp.Text = mhelp
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmItemHelp = Nothing
End Sub


