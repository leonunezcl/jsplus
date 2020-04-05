VERSION 5.00
Begin VB.Form frmStyleSheet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Style Sheet Link"
   ClientHeight    =   1725
   ClientLeft      =   4095
   ClientTop       =   4005
   ClientWidth     =   5520
   ControlBox      =   0   'False
   Icon            =   "frmStyleSheet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   5
      Top             =   1245
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   4
      Top             =   1245
      Width           =   1215
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1530
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3570
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   1155
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   15
      Width           =   5385
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   4935
         TabIndex        =   6
         Top             =   525
         Width           =   375
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   510
         Width           =   4770
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Style Sheet URL"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   285
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmStyleSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub crea_style()

    Dim str As New cStringBuilder
    
    If txtFile.Text = "" Then
        str.Append "<link rel=" & Chr$(34) & "stylesheet" & Chr$(34) & " href=" & Chr$(34) & Chr$(34) & " type=" & Chr$(34) & "text/css" & Chr$(34) & ">"
    Else
        str.Append "<link rel=" & Chr$(34) & "stylesheet" & Chr$(34) & " href=" & Chr$(34) & Replace(txtFile.Text, "\", "/") & Chr$(34) & " type=" & Chr$(34) & "text/css" & Chr$(34) & ">"
    End If
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        frmMain.ActiveForm.Insertar str.ToString
    End If
    
    Set str = Nothing
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call crea_style
        Unload Me
    ElseIf Index = 1 Then
        Unload Me
    Else
        Dim Archivo As String
        Dim glosa As String
        
        glosa = "Cascading Style Sheets (*.css)|*.css|"
        glosa = glosa & "All files (*.*)|*.*"
    
        If Cdlg.VBGetOpenFileName(Archivo, , , , , , glosa, , LastPath) Then
            txtFile.Text = Archivo
        End If
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmStyleSheet = Nothing
End Sub


