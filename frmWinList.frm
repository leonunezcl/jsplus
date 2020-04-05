VERSION 5.00
Begin VB.Form frmWinList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Window List"
   ClientHeight    =   4470
   ClientLeft      =   5250
   ClientTop       =   2115
   ClientWidth     =   7605
   Icon            =   "frmWinList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ListBox lst 
      Height          =   3570
      Left            =   45
      TabIndex        =   0
      Top             =   315
      Width           =   7485
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   7080
      TabIndex        =   4
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Opened Documents ..."
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
      Left            =   45
      TabIndex        =   1
      Top             =   75
      Width           =   1920
   End
End
Attribute VB_Name = "frmWinList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call activar_window
    End If
    
    Unload Me
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Dim k As Integer
    Dim File As New cFile
    Dim frm As Form
    Dim fvisible As Boolean
    
    util.CenterForm Me
    
    util.Hourglass hwnd, True
    
    For k = 1 To Files.Files.count
    
        Set File = New cFile
        Set File = Files.Files.ITem(k)
        
        fvisible = False
        
        For Each frm In Forms
            If TypeName(frm) = "frmEdit" Then
                If frm.Caption = File.Caption Then
                    fvisible = True
                    Exit For
                End If
            End If
        Next
        
        If fvisible Then
            If Len(File.filename) > 0 Then
                If File.Ftp Then
                    lst.AddItem File.Caption
                Else
                    lst.AddItem File.filename
                End If
            Else
                lst.AddItem File.Caption
            End If
            lst.ItemData(lst.NewIndex) = File.IdDoc
        End If
        
        Set File = Nothing
    Next k
    
    lblTot.Caption = CStr(lst.ListCount) & " documents"
    
    Debug.Print "load"
                
    util.Hourglass hwnd, False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmWinList = Nothing
End Sub


Private Sub activar_window()

    Dim frm As Form
    
    util.Hourglass hwnd, True
    
    If lst.ListIndex <> -1 Then
        For Each frm In Forms
            If TypeName(frm) = "frmEdit" Then
                If CInt(frm.Tag) = lst.ItemData(lst.ListIndex) Then
                    Me.Hide
                    On Error Resume Next
                    frm.SetFocus
                    frm.txtCode.SetFocus
                    Err = 0
                    Exit For
                End If
            End If
        Next
    End If
    
    util.Hourglass hwnd, False
    
End Sub
