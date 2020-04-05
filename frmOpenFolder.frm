VERSION 5.00
Begin VB.Form frmOpenFolder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Folder"
   ClientHeight    =   1710
   ClientLeft      =   3960
   ClientTop       =   2280
   ClientWidth     =   7380
   ControlBox      =   0   'False
   Icon            =   "frmOpenFolder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Folder"
      Height          =   1035
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   7245
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   6720
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtFolder 
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   570
         Width           =   6465
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Last opened folder :"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   330
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmOpenFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private maxfilesfolder As String
Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If txtFolder.Text <> "" Then
            Call util.GrabaIni(IniPath(), "open_folder", "path", txtFolder.Text)
            
            Dim afiles() As String
            
            Call get_files_from_folder(txtFolder.Text, afiles)
            
            If UBound(afiles) > CInt(maxfilesfolder) Then
                If Confirma("Confirm open " & UBound(afiles) & " files ? (Max : " & maxfilesfolder & ")") = vbNo Then
                    Exit Sub
                End If
            End If
            
            Me.Hide
            frmMain.open_folder txtFolder.Text
            
            MsgBox "To open file select File Manager Panel from Main Window", vbInformation
            
            Unload Me
        Else
            txtFolder.SetFocus
        End If
    ElseIf Index = 1 Then
        Unload Me
    Else
        txtFolder.Text = util.BrowseFolder(hwnd)
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
        
    txtFolder.Text = util.LeeIni(IniPath(), "open_folder", "path")
        
    maxfilesfolder = util.LeeIni(IniPath, "maxfilesfolder", "value")
    If maxfilesfolder = "" Then maxfilesfolder = "10"
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOpenFolder = Nothing
End Sub


