VERSION 5.00
Object = "{C03D5026-EB27-402A-BD60-0E05020E600B}#2.0#0"; "vbsframex.ocx"
Begin VB.Form frmQuickTip 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "QuickTip"
   ClientHeight    =   3015
   ClientLeft      =   5580
   ClientTop       =   6195
   ClientWidth     =   4860
   Icon            =   "frmQuickTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin vbsFrames.jcFrames framex 
      Height          =   3015
      Left            =   0
      Top             =   0
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   5318
      Caption         =   "jcFrames1"
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.CommandButton cmd 
         Caption         =   "Hide"
         Height          =   255
         Left            =   4080
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lbltype 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Property"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblhelp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   2235
         Left            =   60
         TabIndex        =   0
         Top             =   705
         Width           =   4665
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmQuickTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Height As Long
Private m_Width As Long
Private Sub cmd_Click()

    If cmd.Caption = "Hide" Then
        cmd.Caption = "Expand"
        Me.Height = 780
        Me.Width = m_Width
    Else
        Me.Height = m_Height
        Me.Width = m_Width
        cmd.Caption = "Hide"
    End If
    
    Me.Refresh
    
End Sub

Private Sub Form_Activate()
    windowontop hwnd
End Sub

Private Sub Form_Load()
    
    framex.Caption = vbNullString
    lbltype.Caption = vbNullString
    lblhelp.Caption = vbNullString
    
    m_Height = Me.Height
    m_Width = Me.Width
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    glbquickon = False
    Set frmQuickTip = Nothing
End Sub


