VERSION 5.00
Begin VB.Form frmStatements 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Statements"
   ClientHeight    =   2970
   ClientLeft      =   5580
   ClientTop       =   4335
   ClientWidth     =   3780
   Icon            =   "frmStatements.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Edit"
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   5430
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1350
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Frame fra 
      Caption         =   "Available Statements"
      Height          =   2325
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   3645
      Begin VB.ListBox lstStmt 
         Height          =   1620
         Left            =   90
         TabIndex        =   1
         Top             =   240
         Width           =   3420
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Double click to insert "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   2070
         Width           =   1890
      End
   End
End
Attribute VB_Name = "frmStatements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub InsertarStmt()

    Dim sFile As String
        
    If lstStmt.ListIndex = -1 Then
        Exit Sub
    End If
    
    Select Case lstStmt.ListIndex + 1
        Case 1  'do
            sFile = "jswhile.txt"
        Case 2  'for
            sFile = "jsfor.txt"
        Case 3  'for in
            sFile = "jsforin.txt"
        Case 4  'function
            sFile = "jsfunction.txt"
        Case 5  'if..else
            sFile = "jsifthenelse.txt"
        Case 6  'switch
            sFile = "jsswitch.txt"
        Case 7  'try...catch
            sFile = "jstrycatch.txt"
        Case 8  'var
            frmNewVar.Show vbModal
        Case 9  'while
            sFile = "jswhile.txt"
        Case 10 'with
            sFile = "jswith.txt"
    End Select
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        frmMain.ActiveForm.Insertar frmMain.LoadSnipet(sFile)
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)

   If Index = 0 Then
      Call InsertarStmt
      Unload Me
   ElseIf Index = 2 Then
      Select Case lstStmt.ListIndex + 1
        Case 1  'do
            frmCrearPlantilla.template_file = "jswhile.txt"
            frmCrearPlantilla.Show vbModal
        Case 2  'for
            frmCrearPlantilla.template_file = "jsfor.txt"
            frmCrearPlantilla.Show vbModal
        Case 3  'for in
            frmCrearPlantilla.template_file = "jsforin.txt"
            frmCrearPlantilla.Show vbModal
        Case 4  'function
            frmCrearPlantilla.template_file = "jsfunction.txt"
            frmCrearPlantilla.Show vbModal
        Case 5  'if..else
            frmCrearPlantilla.template_file = "jsifthenelse.txt"
            frmCrearPlantilla.Show vbModal
        Case 6  'switch
            frmCrearPlantilla.template_file = "jsswitch.txt"
            frmCrearPlantilla.Show vbModal
        Case 7  'try...catch
            frmCrearPlantilla.template_file = "jstrycatch.txt"
            frmCrearPlantilla.Show vbModal
        Case 9  'while
            frmCrearPlantilla.template_file = "jswhile.txt"
            frmCrearPlantilla.Show vbModal
        Case 10 'with
            frmCrearPlantilla.template_file = "jswith.txt"
            frmCrearPlantilla.Show vbModal
      End Select
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Call util.CenterForm(Me)
    
    lstStmt.AddItem "do...while"
    lstStmt.AddItem "for"
    lstStmt.AddItem "for...in"
    lstStmt.AddItem "function"
    lstStmt.AddItem "if...then...else"
    lstStmt.AddItem "switch"
    lstStmt.AddItem "try...catch"
    lstStmt.AddItem "var"
    lstStmt.AddItem "while"
    lstStmt.AddItem "with"
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmStatements = Nothing
End Sub


Private Sub lstStmt_DblClick()
    Call InsertarStmt
End Sub

Private Sub lstStmt_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Call InsertarStmt
    End If
    
End Sub


