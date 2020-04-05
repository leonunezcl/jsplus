VERSION 5.00
Begin VB.Form frmLstFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add file to Project ..."
   ClientHeight    =   3315
   ClientLeft      =   3555
   ClientTop       =   4530
   ClientWidth     =   6465
   Icon            =   "frmLstFiles.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   5040
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Add"
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CheckBox chkSel 
      Caption         =   "Select All"
      Height          =   225
      Left            =   45
      TabIndex        =   2
      Top             =   3060
      Width           =   1095
   End
   Begin VB.ListBox lstFiles 
      Height          =   2790
      Left            =   30
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   240
      Width           =   4875
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Files in Project"
      Height          =   195
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   1020
   End
End
Attribute VB_Name = "frmLstFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSel_Click()
   
   Dim ret As Boolean
   Dim k As Integer
   
   util.Hourglass hwnd, True
   
   If chkSel.Value = 1 Then
      ret = True
   Else
      ret = False
   End If
   
   For k = 0 To lstFiles.ListCount - 1
      lstFiles.Selected(k) = ret
   Next k
   
   util.Hourglass hwnd, False
   
End Sub

Private Sub cmd_Click(Index As Integer)

    Dim k As Integer
    Dim j As Integer
    
    Dim pFile As cFile
    
    If Index = 0 Then
        If lstFiles.ListIndex <> -1 Then
volver:
            For j = 0 To lstFiles.ListCount - 1
                If lstFiles.Selected(j) Then
                    With Files
                        For k = 1 To .Files.count
                            Set pFile = New cFile
                            Set pFile = .Files.ITem(k)
                            If pFile.IdDoc = lstFiles.ItemData(j) Then
                                'verificar que archivo a agregar no exista
                                ProjectMan.AddFile pFile
                                Exit For
                            End If
                            Set pFile = Nothing
                        Next k
                    End With
                    lstFiles.RemoveItem j
                    GoTo volver
                End If
            Next j
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Dim k As Integer
    Dim pFile As cFile
    
    util.CenterForm Me
    util.Hourglass hwnd, True
        
    With Files
        For k = 1 To .Files.count
            Set pFile = New cFile
            Set pFile = .Files.ITem(k)
            
            If Len(pFile.filename) > 0 Then
                lstFiles.AddItem pFile.filename
                lstFiles.ItemData(lstFiles.NewIndex) = pFile.IdDoc
            End If
        Next k
    End With
    
    Debug.Print "load"
    
    util.Hourglass hwnd, False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print Me.Name & " unload"
    Set frmLstFiles = Nothing
End Sub


