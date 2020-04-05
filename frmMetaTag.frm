VERSION 5.00
Begin VB.Form frmMetaTag 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MetaTag Wizard"
   ClientHeight    =   7800
   ClientLeft      =   4095
   ClientTop       =   2085
   ClientWidth     =   8715
   Icon            =   "frmMetaTag.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   23
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   22
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   6885
      Index           =   0
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   8610
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   8115
         TabIndex        =   24
         Top             =   5955
         Width           =   375
      End
      Begin VB.TextBox txtTell 
         Height          =   285
         Left            =   105
         TabIndex        =   8
         Top             =   5325
         Width           =   8400
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   105
         TabIndex        =   0
         Top             =   480
         Width           =   8400
      End
      Begin VB.TextBox txtKeywords 
         Height          =   285
         Left            =   105
         TabIndex        =   1
         Top             =   1065
         Width           =   8400
      End
      Begin VB.TextBox txtSiteDes 
         Height          =   285
         Left            =   105
         TabIndex        =   2
         Top             =   1680
         Width           =   8400
      End
      Begin VB.TextBox txtPublisher 
         Height          =   285
         Left            =   105
         TabIndex        =   3
         Top             =   2310
         Width           =   8400
      End
      Begin VB.TextBox txtLang 
         Height          =   285
         Left            =   105
         TabIndex        =   4
         Top             =   2910
         Width           =   8400
      End
      Begin VB.ComboBox cboSearchE 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   4110
         Width           =   8400
      End
      Begin VB.ComboBox cboFollow 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   4710
         Width           =   8400
      End
      Begin VB.TextBox txtCopyright 
         Height          =   285
         Left            =   105
         TabIndex        =   5
         Top             =   3495
         Width           =   8400
      End
      Begin VB.TextBox txtIcon 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   5925
         Width           =   7935
      End
      Begin VB.Image imgico 
         Height          =   540
         Left            =   120
         Top             =   6240
         Width           =   885
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title Tag (use major keywords in title)"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   20
         Top             =   240
         Width           =   2610
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keywords/Phrases (separated by commas and not spaces)"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   19
         Top             =   855
         Width           =   4155
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Site Description"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   18
         Top             =   1455
         Width           =   1110
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Publisher's Name"
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   17
         Top             =   2100
         Width           =   1215
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Content Language"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   16
         Top             =   2700
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allow search engine to index the page "
         Height          =   195
         Index           =   5
         Left            =   105
         TabIndex        =   15
         Top             =   3900
         Width           =   2745
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allow search engine to follow links to other pages "
         Height          =   195
         Index           =   6
         Left            =   105
         TabIndex        =   14
         Top             =   4500
         Width           =   3525
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tell search engine when to return "
         Height          =   195
         Index           =   7
         Left            =   105
         TabIndex        =   13
         Top             =   5100
         Width           =   2415
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         Height          =   195
         Index           =   8
         Left            =   105
         TabIndex        =   12
         Top             =   3285
         Width           =   660
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Link to Icon Image"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   5700
         Width           =   1320
      End
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Browser Compatibility"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   5805
      TabIndex        =   21
      Top             =   7440
      Width           =   1485
   End
   Begin VB.Image imgIE 
      Height          =   255
      Left            =   7410
      Top             =   7425
      Width           =   300
   End
   Begin VB.Image imgFX 
      Height          =   255
      Left            =   7725
      Top             =   7425
      Width           =   300
   End
   Begin VB.Image imgNE 
      Height          =   255
      Left            =   8040
      Top             =   7425
      Width           =   300
   End
   Begin VB.Image imgOP 
      Height          =   255
      Left            =   8370
      Top             =   7425
      Width           =   300
   End
End
Attribute VB_Name = "frmMetaTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function metatag() As Boolean

    Dim src As New cStringBuilder
    
    src.Append "<TITLE>" & txtTitle.Text & "</TITLE>" & vbNewLine
    src.Append "<META http-equiv=" & Chr$(34) & "Content-Type" & Chr$(34) & " CONTENT=" & Chr$(34) & "text/html;" & Chr$(34) & ">" & vbNewLine ' charset=iso-8859-1" & Chr$(34) & ">" & vbNewLine
    src.Append "<META NAME=" & Chr$(34) & "DESCRIPTION" & Chr$(34) & " CONTENT=" & Chr$(34) & txtSiteDes.Text & Chr$(34) & ">" & vbNewLine
    src.Append "<META NAME=" & Chr$(34) & "KEYWORDS" & Chr$(34) & " CONTENT=" & Chr$(34) & txtKeywords.Text & Chr$(34) & ">" & vbNewLine
    src.Append "<META NAME=" & Chr$(34) & "Generator" & Chr$(34) & " CONTENT=" & Chr$(34) & "JavaScript Plus! - http://www.vbsoftware.cl" & Chr$(34) & ">" & vbNewLine
    src.Append "<META NAME=" & Chr$(34) & "PUBLISHER" & Chr$(34) & " CONTENT=" & Chr$(34) & txtPublisher.Text & Chr$(34) & ">" & vbNewLine
    src.Append "<META NAME=" & Chr$(34) & "LANGUAGE" & Chr$(34) & " CONTENT=" & Chr$(34) & txtLang.Text & Chr$(34) & ">" & vbNewLine
    src.Append "<META NAME=" & Chr$(34) & "COPYRIGHT" & Chr$(34) & " CONTENT=" & Chr$(34) & txtCopyright.Text & Chr$(34) & ">" & vbNewLine
    src.Append "<META NAME=" & Chr$(34) & "ROBOTS" & Chr$(34) & " CONTENT=" & Chr$(34) & cboSearchE.Text & Chr$(34) & ">" & vbNewLine
    src.Append "<META NAME=" & Chr$(34) & "ROBOTS" & Chr$(34) & " CONTENT=" & Chr$(34) & cboFollow.Text & Chr$(34) & ">" & vbNewLine
    src.Append "<META NAME=" & Chr$(34) & "REVISIT-AFTER" & Chr$(34) & " CONTENT=" & Chr$(34) & txtTell.Text & Chr$(34) & ">" & vbNewLine
    src.Append "<link rel=" & Chr$(34) & "SHORTCUT ICON" & Chr$(34) & " href=" & Chr$(34) & txtIcon.Text & Chr$(34) & ">" & vbNewLine

    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(src.ToString)
    End If
    
    Set src = Nothing
    
    metatag = True
    
End Function

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If metatag() Then
            Unload Me
        End If
    ElseIf Index = 2 Then
        Dim Archivo As String
        Dim glosa As String
        
        glosa = "CompuServe Graphics Interchange (*.gif)|*.gif|"
        glosa = glosa & "jpg (*.jpg)|*.jpg|"
        glosa = glosa & "ico (*.ico)|*.ico|"
        glosa = glosa & "All files (*.*)|*.*"
    
        If Cdlg.VBGetOpenFileName(Archivo, , , , , , glosa, , LastPath) Then
            txtIcon.Text = Archivo
            On Error Resume Next
            Set imgico.Picture = Nothing
            imgico.Picture = LoadPicture(Archivo)
            If Err <> 0 Then
                MsgBox "File Error " & Err & " " & Error$, vbCritical
            End If
            Err = 0
        End If
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
        
    Set imgIE.Picture = LoadResPicture(1007, vbResBitmap)
    Set imgFX.Picture = LoadResPicture(1008, vbResBitmap)
    Set imgNE.Picture = LoadResPicture(1009, vbResBitmap)
    Set imgOP.Picture = LoadResPicture(1010, vbResBitmap)
    
    cboSearchE.AddItem "index"
    cboSearchE.AddItem "noindex"
    
    cboFollow.AddItem "follow"
    cboFollow.AddItem "nofollow"
        
    Debug.Print "load : " & Me.Name
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call clear_memory(Me)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload : " & Me.Name
    Set frmMetaTag = Nothing
End Sub


