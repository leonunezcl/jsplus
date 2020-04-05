VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "My VB Program"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":0000
      Height          =   1095
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim IRes As Integer
IRes = IntegrityOK
If IRes = 1 Then
   MsgBox App.EXEName & ".exe has not been tampered with :-)", vbInformation + vbOKOnly, "CRC32 Check OK!"
ElseIf IRes = -1 Then
   MsgBox App.EXEName & ".exe doesn't have a CRC footer!", vbExclamation + vbOKOnly, "CRC32 Error"
Else
   MsgBox UCase(App.EXEName) & ".EXE HAS BEEN TAMPERED WITH!", vbExclamation + vbOKOnly, "CRC32 ALARM"
End If
End Sub
