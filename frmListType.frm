VERSION 5.00
Begin VB.Form frmListType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List"
   ClientHeight    =   4620
   ClientLeft      =   4305
   ClientTop       =   3210
   ClientWidth     =   5865
   Icon            =   "frmListType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Remove"
      Height          =   285
      Index           =   3
      Left            =   5055
      TabIndex        =   16
      Top             =   2460
      Width           =   735
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Add"
      Height          =   255
      Index           =   2
      Left            =   5055
      TabIndex        =   15
      Top             =   1920
      Width           =   720
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3045
      TabIndex        =   14
      Top             =   4140
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1005
      TabIndex        =   13
      Top             =   4140
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   1755
      Index           =   0
      Left            =   75
      TabIndex        =   7
      Top             =   90
      Width           =   5760
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   255
         Index           =   5
         Left            =   5370
         TabIndex        =   18
         Top             =   630
         Width           =   315
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   255
         Index           =   4
         Left            =   5385
         TabIndex        =   17
         Top             =   1335
         Width           =   315
      End
      Begin VB.TextBox txtImgBul 
         Height          =   285
         Left            =   735
         TabIndex        =   1
         Top             =   615
         Width           =   4590
      End
      Begin VB.TextBox txtImgSty 
         Height          =   285
         Left            =   735
         TabIndex        =   3
         Top             =   1335
         Width           =   4590
      End
      Begin VB.ComboBox cboStyle 
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   735
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   975
         Width           =   3090
      End
      Begin VB.ComboBox cboBullets 
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   735
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   255
         Width           =   3090
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Image"
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   12
         Top             =   630
         Width           =   435
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Image"
         Height          =   195
         Index           =   11
         Left            =   60
         TabIndex        =   10
         Top             =   1350
         Width           =   435
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Style"
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   9
         Top             =   1005
         Width           =   345
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bullet"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   300
         Width           =   390
      End
   End
   Begin VB.TextBox txtItem 
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Top             =   1905
      Width           =   4530
   End
   Begin VB.ListBox lstItems 
      ForeColor       =   &H00404040&
      Height          =   1620
      Left            =   75
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   2445
      Width           =   4920
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   195
      Index           =   3
      Left            =   75
      TabIndex        =   11
      Top             =   1920
      Width           =   300
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List Items"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   6
      Top             =   2250
      Width           =   660
   End
End
Attribute VB_Name = "frmListType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mycaption As String
Public mytype As Integer
Private Sub crea_lista()

    Dim src As New cStringBuilder
    Dim k As Integer
    
    If mytype = 0 Then  'ordenada
        src.Append "<ol"
        
        If cboBullets.ListIndex > 0 Then
            If cboBullets.ListIndex = 1 Then
                src.Append " type=" & Chr$(34) & "l" & Chr$(34)
            ElseIf cboBullets.ListIndex = 2 Then
                src.Append " type=" & Chr$(34) & "a" & Chr$(34)
            ElseIf cboBullets.ListIndex = 3 Then
                src.Append " type=" & Chr$(34) & "A" & Chr$(34)
            ElseIf cboBullets.ListIndex = 4 Then
                src.Append " type=" & Chr$(34) & "i" & Chr$(34)
            ElseIf cboBullets.ListIndex = 5 Then
                src.Append " type=" & Chr$(34) & "I" & Chr$(34)
            End If
        End If
        
        If cboStyle.ListIndex <> -1 Then
            src.Append " style=" & Chr$(34) & "list-style-type: " & cboStyle.Text
            If txtImgSty.Text <> "" Then
                src.Append "; list-style-image: url('" & txtImgSty.Text & "')"
            End If
            src.Append Chr$(34)
        ElseIf txtImgSty.Text <> "" Then
            src.Append " style=" & Chr$(34) & "list-style-image: url('" & txtImgSty.Text & "')" & Chr$(34)
        End If
        src.Append ">" & vbNewLine
    Else
        src.Append "<ul"
        
        If cboBullets.ListIndex > 0 Then
            src.Append " type=" & Chr$(34) & cboBullets.Text & Chr$(34)
        End If
        
        If txtImgBul.Text <> "" Then
            src.Append " imgsrc=" & Chr$(34) & txtImgBul.Text & Chr$(34)
        End If
        
        If cboStyle.ListIndex <> -1 Then
            src.Append " style=" & Chr$(34) & "list-style-type: " & cboStyle.Text
            If txtImgSty.Text <> "" Then
                src.Append "; list-style-image: url('" & txtImgSty.Text & "')"
            End If
            src.Append Chr$(34)
        ElseIf txtImgSty.Text <> "" Then
            src.Append " style=" & Chr$(34) & "list-style-image: url('" & txtImgSty.Text & "')" & Chr$(34)
        End If
        src.Append ">" & vbNewLine
    End If
    
    For k = 0 To lstItems.ListCount - 1
        src.Append "<li>" & lstItems.List(k) & vbNewLine
    Next k
    
    If mytype = 0 Then
        src.Append "</ol>" & vbNewLine
    Else
        src.Append "</ul>" & vbNewLine
    End If
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(src.ToString)
    End If
    
    Set src = Nothing
    
End Sub


Private Sub cmd_Click(Index As Integer)

    Dim glosa As String
    Dim Archivo As String

    If Index = 0 Then
        Call crea_lista
        Unload Me
    ElseIf Index = 1 Then
        Unload Me
    ElseIf Index = 2 Then
        If Len(Trim$(txtItem.Text)) > 0 Then
            lstItems.AddItem txtItem.Text
            txtItem.SetFocus
        End If
    ElseIf Index = 3 Then
        If lstItems.ListIndex <> -1 Then
            lstItems.RemoveItem lstItems.ListIndex
        End If
    ElseIf Index = 4 Then
    
        glosa = "CompuServe Graphics Interchange (*.gif)|*.gif|"
        glosa = glosa & "JPG (*.jpg)|*.jpg|"
        glosa = glosa & "All files (*.*)|*.*"
        
        If Not Cdlg.VBGetOpenFileName(Archivo, , , , , , glosa, , App.Path, "Open Image", "gif") Then
            Exit Sub
        End If
            
        txtImgSty.Text = util.VBArchivoSinPath(Archivo)
    ElseIf Index = 5 Then
    
        glosa = "CompuServe Graphics Interchange (*.gif)|*.gif|"
        glosa = glosa & "JPG (*.jpg)|*.jpg|"
        glosa = glosa & "All files (*.*)|*.*"
        
        If Not Cdlg.VBGetOpenFileName(Archivo, , , , , , glosa, , App.Path, "Open Image", "gif") Then
            Exit Sub
        End If
            
        txtImgBul.Text = util.VBArchivoSinPath(Archivo)
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
        
    Me.Caption = mycaption
        
    If mytype = 0 Then  'ordenada
        lbl(1).Caption = "Numbers"
        cboBullets.AddItem "default"
        cboBullets.AddItem "decimal"
        cboBullets.AddItem "lower-alpha"
        cboBullets.AddItem "upper-alpha"
        cboBullets.AddItem "lower-roman"
        cboBullets.AddItem "upper-roman"
        txtImgBul.Enabled = False
        cmd(5).Enabled = False
    Else                'desordenada
        cboBullets.AddItem "default"
        cboBullets.AddItem "disk"
        cboBullets.AddItem "square"
        cboBullets.AddItem "circle"
    End If
    
    cboStyle.AddItem "default"
    cboStyle.AddItem "decimal"
    cboStyle.AddItem "upper-alpha"
    cboStyle.AddItem "lower-alpha"
    cboStyle.AddItem "upper-roman"
    cboStyle.AddItem "lower-roman"
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmListType = Nothing
End Sub


Private Sub txtItem_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        cmd_Click 2
        txtItem.SetFocus
    End If
    
End Sub


