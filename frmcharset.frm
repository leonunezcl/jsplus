VERSION 5.00
Begin VB.Form frmCharSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Character Set"
   ClientHeight    =   5370
   ClientLeft      =   3180
   ClientTop       =   3090
   ClientWidth     =   5040
   Icon            =   "frmcharset.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Available Character Set"
      Height          =   5250
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   3585
      Begin VB.ListBox lst 
         Height          =   4545
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3300
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   255
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmCharSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub carga_charset()

    Dim arr_charset() As String
    Dim inifile As String
    Dim k As Integer
    
    inifile = util.StripPath(App.Path) & "config\set.ini"
    
    get_info_section "set", arr_charset, inifile
    
    For k = 2 To UBound(arr_charset)
        lst.AddItem arr_charset(k)
    Next k
    
End Sub

Private Sub insertar_caracter()

    Dim buf As New cStringBuilder
    
    buf.Append lst.Text
    
    If Not frmMain.ActiveForm Is Nothing Then
        frmMain.ActiveForm.Insertar buf.ToString
    End If
    
    Set buf = Nothing
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If lst.ListIndex <> -1 Then
            Call insertar_caracter
            Unload Me
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
    util.Hourglass hwnd, True
        
    Call carga_charset
        
    util.Hourglass hwnd, False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmCharSet = Nothing
End Sub


