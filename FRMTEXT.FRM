VERSION 5.00
Begin VB.Form frmHtmlText 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TextBox"
   ClientHeight    =   2775
   ClientLeft      =   4395
   ClientTop       =   2700
   ClientWidth     =   6435
   Icon            =   "FRMTEXT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2775
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&More"
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Events"
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   13
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Events"
      Height          =   690
      Index           =   3
      Left            =   45
      TabIndex        =   6
      Top             =   1545
      Width           =   6330
      Begin VB.CommandButton cmd 
         Caption         =   "&Delete"
         Height          =   375
         Index           =   4
         Left            =   5640
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cboEvents 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   705
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   255
         Width           =   4860
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   7
         Top             =   285
         Width           =   510
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   1410
      Index           =   0
      Left            =   45
      TabIndex        =   5
      Top             =   60
      Width           =   6330
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   870
         TabIndex        =   0
         Top             =   285
         Width           =   5325
      End
      Begin VB.TextBox txtValor 
         Height          =   285
         Left            =   870
         TabIndex        =   3
         Top             =   975
         Width           =   5325
      End
      Begin VB.TextBox txtTam 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   870
         MaxLength       =   3
         TabIndex        =   1
         Top             =   630
         Width           =   495
      End
      Begin VB.TextBox txtCm 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5565
         TabIndex        =   2
         Top             =   630
         Width           =   615
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   11
         Top             =   285
         Width           =   420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   10
         Top             =   1005
         Width           =   405
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Length"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   9
         Top             =   630
         Width           =   495
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   4815
         TabIndex        =   8
         Top             =   630
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmHtmlText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tipo_texto As String

Private Function InsertarTexto() As Boolean

    On Error GoTo ErrorInsertarTexto

    Dim src As New cStringBuilder
    
    If txtNombre.Text = "" Then
        MsgBox "The name control please.", vbCritical
        txtNombre.SetFocus
        Exit Function
    End If
    
    src.Append "<input type=" & Chr$(34) & tipo_texto & Chr$(34)
    src.Append " name=" & Chr$(34) & txtNombre.Text & Chr$(34)
    If txtTam.Text = "" Then txtTam.Text = "40"
    src.Append " size=" & Chr$(34) & txtTam.Text & Chr$(34)
    If txtCm.Text = "" Then txtCm.Text = "255"
    src.Append " maxlength=" & Chr$(34) & txtCm.Text & Chr$(34)
    src.Append " value=" & Chr$(34) & txtValor.Text & Chr$(34)
    
    'agregar atributos especiales
    src.Append CComHtmlAttrib.Output
        
    'agregar los eventos
    src.Append CEventos.Output
    
    src.Append ">"
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(src.ToString)
    End If
    
    InsertarTexto = True
    
    Exit Function
    
ErrorInsertarTexto:
    MsgBox "InsertarTexto : " & Err & " " & Error$, vbCritical
    
End Function

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If InsertarTexto() Then
            Unload Me
        End If
    ElseIf Index = 1 Then
        Unload Me
    ElseIf Index = 2 Then
        frmEvents.html_tag = tipo_texto
        frmEvents.Show vbModal
        Call CEventos.Attach(Me.cboEvents)
    ElseIf Index = 3 Then
        frmCommonHtml.html_tag = tipo_texto
        frmCommonHtml.Show vbModal
    ElseIf Index = 4 Then
        If cboEvents.ListIndex <> -1 Then
            CEventos.Remove cboEvents.Text
            cboEvents.RemoveItem cboEvents.ListIndex
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
        
    util.SetNumber txtTam.hwnd
    util.SetNumber txtCm.hwnd
    
    CEventos.Clear
    CComHtmlAttrib.Clear
    
    Debug.Print "load"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmHtmlText = Nothing
End Sub


