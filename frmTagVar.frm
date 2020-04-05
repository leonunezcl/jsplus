VERSION 5.00
Begin VB.Form frmTagVar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Titulo"
   ClientHeight    =   5070
   ClientLeft      =   3525
   ClientTop       =   2325
   ClientWidth     =   6690
   Icon            =   "frmTagVar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtHelp 
      Height          =   1515
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2925
      Width           =   6525
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   90
      TabIndex        =   0
      Top             =   315
      Width           =   6525
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   2730
      Width           =   330
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Variables"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   705
   End
End
Attribute VB_Name = "frmTagVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public titulo As String
Public file As String
Public prefijo As String
Public lang As String
Private arr_help() As String

Private Sub insertar_variable()

    Dim src As New cStringBuilder
    
    If List1.ListIndex <> -1 Then
        If frmMain.ActiveForm.Name = "frmEdit" Then
            If lang = "ASP" Then
                src.Append prefijo & "(" & Chr$(34) & List1.List(List1.ListIndex) & Chr$(34) & ")"
            ElseIf lang = "PHP" Then
                src.Append prefijo & "[" & Chr$(34) & List1.List(List1.ListIndex) & Chr$(34) & "]"
            ElseIf lang = "SSI" Then
                src.Append "<!--" & prefijo & Chr$(34) & List1.List(List1.ListIndex) & Chr$(34) & " -->"
            End If
            frmMain.ActiveForm.Insertar src.ToString
        End If
    End If
    
    Set src = Nothing
    
End Sub

Public Sub open_varfile()

    Dim nFreeFile As Long
    Dim linea As String
    Dim k As Integer
    
    nFreeFile = FreeFile
    
    file = util.StripPath(App.Path) & "config\" & file
    
    If ArchivoExiste2(file) Then
        k = 1
        Open file For Input As #nFreeFile
            Do While Not EOF(nFreeFile)
                Line Input #nFreeFile, linea
                ReDim Preserve arr_help(k)
                List1.AddItem Explode(linea, 1, "|")
                arr_help(k) = Explode(linea, 2, "|")
                k = k + 1
            Loop
        Close #nFreeFile
        
        List1.ListIndex = 0
    Else
        MsgBox "File not found :" & file, vbCritical
    End If
    
End Sub


Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call insertar_variable
        Unload Me
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

    util.CenterForm Me
    
    util.Hourglass hwnd, True
        
    ReDim arr_help(0)
    
    Me.Caption = titulo
    
    Call open_varfile
    
    util.Hourglass hwnd, False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTagVar = Nothing
End Sub


Private Sub List1_Click()
    txtHelp.Text = arr_help(List1.ListIndex + 1)
End Sub


