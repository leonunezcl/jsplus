VERSION 5.00
Begin VB.Form frmInsForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form"
   ClientHeight    =   3750
   ClientLeft      =   5010
   ClientTop       =   3645
   ClientWidth     =   6630
   Icon            =   "frmInsForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&More"
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   18
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Events"
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   17
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   16
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   15
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Events"
      Height          =   690
      Index           =   3
      Left            =   45
      TabIndex        =   13
      Top             =   2415
      Width           =   6540
      Begin VB.CommandButton cmd 
         Caption         =   "&Delete"
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cboEvents 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   705
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   255
         Width           =   4935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   14
         Top             =   285
         Width           =   510
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   2310
      Index           =   2
      Left            =   45
      TabIndex        =   9
      Top             =   60
      Width           =   6555
      Begin VB.Frame fra 
         Caption         =   "Send Method"
         Height          =   750
         Index           =   1
         Left            =   855
         TabIndex        =   12
         Top             =   1455
         Width           =   5565
         Begin VB.OptionButton optEnv 
            Caption         =   "Get"
            Height          =   240
            Index           =   0
            Left            =   1350
            TabIndex        =   6
            Top             =   300
            Width           =   780
         End
         Begin VB.OptionButton optEnv 
            Caption         =   "Post"
            Height          =   240
            Index           =   1
            Left            =   3495
            TabIndex        =   7
            Top             =   300
            Value           =   -1  'True
            Width           =   780
         End
      End
      Begin VB.Frame fra 
         Caption         =   "Select protocol to use"
         Height          =   810
         Index           =   0
         Left            =   855
         TabIndex        =   11
         Top             =   600
         Width           =   5565
         Begin VB.OptionButton opt 
            Caption         =   "None"
            Height          =   225
            Index           =   0
            Left            =   165
            TabIndex        =   1
            Top             =   330
            Width           =   1050
         End
         Begin VB.OptionButton opt 
            Caption         =   "Http"
            Height          =   225
            Index           =   1
            Left            =   4380
            TabIndex        =   5
            Top             =   330
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.OptionButton opt 
            Caption         =   "Ftp"
            Height          =   225
            Index           =   2
            Left            =   3540
            TabIndex        =   4
            Top             =   330
            Width           =   675
         End
         Begin VB.OptionButton opt 
            Caption         =   "Email"
            Height          =   225
            Index           =   3
            Left            =   1335
            TabIndex        =   2
            Top             =   330
            Width           =   780
         End
         Begin VB.OptionButton opt 
            Caption         =   "News"
            Height          =   225
            Index           =   4
            Left            =   2520
            TabIndex        =   3
            Top             =   330
            Width           =   750
         End
      End
      Begin VB.TextBox txtAccion 
         Height          =   285
         Left            =   855
         TabIndex        =   0
         Top             =   255
         Width           =   5565
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Action"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   285
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmInsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Function InsertarFormulario() As Boolean

    On Error GoTo ErrorInsertarFormulario
    
    Dim src As New cStringBuilder
    Dim Accion As String
    
    src.Append "<form"
    
    Accion = txtAccion.Text
    
    If opt(1).Value Then
        If Accion = "" Then
            Accion = "http://www.mypage.com"
        End If
    ElseIf opt(2).Value Then
        If Accion = "" Then
            Accion = "ftp://user:password@ftp.mypage.com/"
        End If
    ElseIf opt(3).Value Then
        If Accion = "" Then
            Accion = "mailto:email@mypage.com"
        End If
    ElseIf opt(4).Value Then
        If Accion = "" Then
            Accion = "news:alt.newsgroup"
        End If
    End If
    
    src.Append " action=" & Chr$(34) & Accion & Chr$(34) & " method=" & Chr$(34)
    If optEnv(0).Value Then
        src.Append "get"
    Else
        src.Append "post"
    End If
            
    src.Append Chr$(34)
    
    src.Append " ENCTYPE=" & Chr$(34) & "text/plain" & Chr$(34)
    
    'agregar atributos especiales
    If Len(Trim$(CComHtmlAttrib.Output)) > 0 Then
        src.Append " " & Trim$(CComHtmlAttrib.Output)
    End If
    
    'agregar los eventos
    If Len(Trim$(CEventos.Output)) > 0 Then
        src.Append CEventos.Output
    End If
    
    src.Append ">" & vbNewLine & vbNewLine
    
    src.Append "</form>"
    
    Call frmMain.ActiveForm.Insertar(src.ToString)
    
    Set src = Nothing
    
    InsertarFormulario = True
    
    Exit Function
ErrorInsertarFormulario:
    MsgBox "InsertarFormulario : " & Err & " " & Error$, vbCritical
    
End Function

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If InsertarFormulario() Then
            Unload Me
        End If
    ElseIf Index = 1 Then
        Unload Me
    ElseIf Index = 2 Then
        frmEvents.html_tag = "image"
        frmEvents.Show vbModal
        Call CEventos.Attach(Me.cboEvents)
    ElseIf Index = 3 Then
        frmCommonHtml.html_tag = "image"
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
        
    CEventos.Clear
    CComHtmlAttrib.Clear
    
    Debug.Print "load"
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmInsForm = Nothing
End Sub


