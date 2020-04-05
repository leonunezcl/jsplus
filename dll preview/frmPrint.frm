VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir"
   ClientHeight    =   2085
   ClientLeft      =   4155
   ClientTop       =   5385
   ClientWidth     =   4170
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   1680
      Width           =   975
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Selección de Página"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtEnd 
         Height          =   315
         Left            =   3000
         TabIndex        =   10
         Text            =   "0"
         Top             =   990
         Width           =   855
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "Rango de Páginas"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1740
      End
      Begin VB.TextBox txtStart 
         Height          =   315
         Left            =   3000
         TabIndex        =   5
         Text            =   "0"
         Top             =   630
         Width           =   855
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "Todas las páginas"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1725
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "Página Actual"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label lblEnd 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Término:"
         Height          =   195
         Left            =   2280
         TabIndex        =   11
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "Página:"
         Height          =   255
         Left            =   2115
         TabIndex        =   8
         Top             =   300
         Width           =   780
      End
      Begin VB.Label lblPages 
         Alignment       =   2  'Center
         Caption         =   "0 / 0"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lblStart 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicio:"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   660
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================
' Project:      Enhance Print object
' Author:       edward moth
' Copyright:    © 2000 qbd software ltd
'
' ==============================================================
' Module:       frmPrint
' Purpose:      User options for printing
' ==============================================================

Option Explicit

Private mvarStart As Integer
Private mvarEnd As Integer
Private mvarCurrent As Integer
Private mvarMax As Integer
Private mvarPrint As Boolean


Private Sub cmdCancel_Click()

mvarPrint = False
Me.Hide

End Sub

Private Sub cmdPrint_Click()
Dim lStart As Integer, lEnd As Integer
Dim bEnable As Boolean

bEnable = True
lStart = Val(txtStart.Text)
lEnd = Val(txtStart.Text)

If lStart = 0 Or lEnd = 0 Then
  bEnable = False
End If
If lStart > lEnd Then
  bEnable = False
End If
If lStart <> CInt(lStart) Then
  bEnable = False
End If
If lEnd <> CInt(lEnd) Then
  bEnable = False
End If

If optPrint(0).Value Then
  bEnable = True
  mvarStart = mvarCurrent
  mvarEnd = mvarCurrent
ElseIf optPrint(1).Value Then
  bEnable = True
  mvarStart = 1
  mvarEnd = mvarMax
ElseIf optPrint(2).Value Then
  mvarStart = lStart
  mvarEnd = lEnd
End If

If Not bEnable Then
  MsgBox "Favor ingrese un rango de páginas válido.", vbOKOnly
  mvarPrint = False
Else
  mvarPrint = True
  Me.Hide
End If

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Form_Activate()

optPrint(0).Value = True
optPrint(1).Enabled = CBool(mvarMax > 1)
optPrint(2).Enabled = CBool(mvarMax > 1)
txtStart.Enabled = CBool(mvarMax > 1)
txtEnd.Enabled = CBool(mvarMax > 1)
lblPages.Caption = mvarCurrent & " / " & mvarMax
txtStart.Text = "1"
txtEnd.Text = mvarMax

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmPrint = Nothing
End Sub


Private Sub optPrint_Click(Index As Integer)

txtStart.Locked = CBool(Index <> 2)
txtEnd.Locked = CBool(Index <> 2)
If Index = 0 Then
  txtStart.Text = mvarCurrent
  txtEnd.Text = mvarCurrent
Else
  txtStart.Text = 1
  txtEnd.Text = mvarMax
End If
End Sub

Public Property Get PrintDoc() As Boolean
PrintDoc = mvarPrint
End Property


Public Property Let PageCurrent(ByVal vNewValue As Integer)
mvarCurrent = vNewValue
End Property

Public Property Get PageStart() As Integer
PageStart = mvarStart
End Property

Public Property Get PageEnd() As Integer
PageEnd = mvarEnd
End Property

Public Property Let PageMax(ByVal vNewValue As Integer)
mvarMax = vNewValue
End Property
