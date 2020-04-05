VERSION 5.00
Begin VB.Form frmProMan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Project Manager"
   ClientHeight    =   3135
   ClientLeft      =   4185
   ClientTop       =   3600
   ClientWidth     =   6870
   Icon            =   "frmProMan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Exit"
      Height          =   375
      Index           =   3
      Left            =   5520
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Save"
      Height          =   375
      Index           =   2
      Left            =   5520
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Remove"
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Add"
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.ListBox lstFiles 
      Height          =   2790
      Left            =   75
      TabIndex        =   1
      Top             =   285
      Width           =   5325
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Files in Project"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   1020
   End
End
Attribute VB_Name = "frmProMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cargar_archivos()

    Dim k As Integer
    Dim pFile As cFile
    
    lstFiles.Clear
    
    With ProjectMan
        For k = 1 To .Files.count
            Set pFile = New cFile
            Set pFile = .Files.ITem(k)
            lstFiles.AddItem pFile.filename
            lstFiles.ItemData(lstFiles.NewIndex) = pFile.IdDoc
        Next k
    End With
    
End Sub


Private Sub cmd_Click(Index As Integer)

   If Index = 0 Then
      'agregar archivo
      frmLstFiles.Show vbModal
      Call cargar_archivos
   ElseIf Index = 1 Then
       'eliminar archivo
       If lstFiles.ListIndex <> -1 Then
           Call ProjectMan.DeleteFile(lstFiles.Text)
           lstFiles.RemoveItem lstFiles.ListIndex
       Else
           MsgBox "You must select a file to remove.", vbCritical
       End If
   ElseIf Index = 2 Then
       'guardar
       If lstFiles.ListCount - 1 > -1 Then
           Call ProjectMan.SaveProject
           Unload Me
       Else
           MsgBox "You must add files to the current project.", vbCritical
       End If
   ElseIf Index = 3 Then
       'salir
       Unload Me
   End If
    
End Sub

Private Sub Form_Load()
    
    Me.Caption = ProjectMan.nombre
    
    util.CenterForm Me
    
    util.Hourglass hwnd, True
    
    Call cargar_archivos
        
    util.Hourglass hwnd, False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print Me.Name & " unload"
    Set frmProMan = Nothing
End Sub


