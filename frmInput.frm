VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Input"
   ClientHeight    =   1755
   ClientLeft      =   3885
   ClientTop       =   5040
   ClientWidth     =   5715
   ControlBox      =   0   'False
   Icon            =   "frmInput.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      TabIndex        =   3
      Top             =   1275
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3450
      TabIndex        =   2
      Top             =   1275
      Width           =   990
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   4710
   End
   Begin VB.Image picIcon 
      Height          =   720
      Left            =   120
      Picture         =   "frmInput.frx":000C
      Top             =   180
      Width           =   720
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   225
      Width           =   4635
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
  Result = vbNullString
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Result = txtInput.Text
  Unload Me
End Sub

Private Sub Form_Load()
  
    util.CenterForm Me
    'SetLayered hwnd, True
    
    Debug.Print "load"
    
    'DrawXPCtl Me
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmInput = Nothing
End Sub

