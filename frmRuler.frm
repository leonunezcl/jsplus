VERSION 5.00
Object = "{246E535D-09D2-4109-80DA-2FF183F4D185}#2.1#0"; "colorpick.ocx"
Begin VB.Form frmRuler 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Horizontal Line"
   ClientHeight    =   2040
   ClientLeft      =   4035
   ClientTop       =   3165
   ClientWidth     =   3345
   Icon            =   "frmRuler.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   1455
      Index           =   0
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   3255
      Begin VB.TextBox txtColorText 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   630
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   1
         Top             =   675
         Width           =   900
      End
      Begin VB.ComboBox cboAlign 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   630
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   2415
      End
      Begin VB.CheckBox chk 
         Caption         =   "No shading"
         Height          =   255
         Left            =   630
         TabIndex        =   2
         Top             =   1035
         Width           =   1185
      End
      Begin ColorPick.ClrPicker ClrPicker1 
         Height          =   300
         Left            =   1530
         TabIndex        =   6
         Top             =   660
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Align"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   330
         Width           =   345
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   705
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmRuler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub insertar_regla()

    Dim src As New cStringBuilder
    
    src.Append "<hr"
    
    If cboAlign.ListIndex <> -1 Then
        src.Append " align=" & Chr$(34) & cboAlign.Text & Chr$(34)
    End If
    
    If txtColorText.Text <> "" Then
        src.Append " color=" & Chr$(34) & txtColorText.Text & Chr$(34)
    End If
    
    If chk.Value Then
        src.Append " noshade=" & Chr$(34) & "noshade" & Chr$(34)
    End If
    
    src.Append ">"
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(src.ToString)
    End If
    
    Set src = Nothing
    
End Sub




Private Sub ClrPicker1_ColorSelected(m_Color As stdole.OLE_COLOR, m_Code As String)
    txtColorText.Text = m_Code
    txtColorText.Tag = m_Color
End Sub


Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call insertar_regla
        Unload Me
    ElseIf Index = 1 Then
        Unload Me
    ElseIf Index = 2 Then
        'picker
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    util.CenterForm Me
        
    cboAlign.AddItem "left"
    cboAlign.AddItem "center"
    cboAlign.AddItem "right"
    
    ClrPicker1.PathPaleta = App.Path & "\pal\256c.pal"
    Debug.Print "load"
       
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmRuler = Nothing
End Sub


