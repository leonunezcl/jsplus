VERSION 5.00
Begin VB.Form frmPageTran 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Page Transitions"
   ClientHeight    =   2385
   ClientLeft      =   2865
   ClientTop       =   6075
   ClientWidth     =   5415
   Icon            =   "frmPageTran.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   1455
      Index           =   0
      Left            =   45
      TabIndex        =   4
      Top             =   105
      Width           =   5295
      Begin VB.OptionButton opt 
         Caption         =   "Exiting"
         Height          =   270
         Index           =   1
         Left            =   2400
         TabIndex        =   3
         Top             =   1005
         Width           =   1080
      End
      Begin VB.OptionButton opt 
         Caption         =   "Entering"
         Height          =   270
         Index           =   0
         Left            =   900
         TabIndex        =   2
         Top             =   1005
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.ComboBox cboEfect 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   285
         Width           =   4335
      End
      Begin VB.TextBox txtTime 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   900
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "3"
         Top             =   645
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Page Event"
         Height          =   390
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   930
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   705
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   5
         Top             =   345
         Width           =   360
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
      Index           =   3
      Left            =   3465
      TabIndex        =   8
      Top             =   2115
      Width           =   1485
   End
   Begin VB.Image imgIE 
      Height          =   255
      Left            =   5070
      Top             =   2100
      Width           =   300
   End
End
Attribute VB_Name = "frmPageTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub create_efect()

    Dim src As New cStringBuilder
    
    If cboEfect.ListIndex = -1 Then Exit Sub
    
    src.Append "<html>" & vbNewLine
    src.Append "<head>" & vbNewLine
    src.Append "<meta http-equiv=" & Chr$(34) & "Content-Type" & Chr$(34) & "CONTENT=" & Chr$(34) & "text/html;" & Chr$(34) & ">" & vbNewLine
    src.Append "<META http-equiv="
    If opt(0).Value Then
        src.Append Chr$(34) & "Page-Enter" & Chr$(34)
    Else
        src.Append Chr$(34) & "Page-Exit" & Chr$(34)
    End If
    
    src.Append " content="
    src.Append Chr$(34) & "revealTrans(Duration=" & txtTime.Text & ","
    src.Append "Transition=" & cboEfect.ListIndex & ")" & Chr$(34) & ">"
    src.Append vbNewLine
    src.Append "<title>Testing Page</title>" & vbNewLine
    src.Append "</head>" & vbNewLine
    src.Append "<body>" & vbNewLine
    src.Append "<p>Have a nice day!" & vbNewLine
    src.Append "</body>" & vbNewLine
    src.Append "</html>" & vbNewLine
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(src.ToString)
    End If
    
    Set src = Nothing
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call create_efect
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
    util.SetNumber Me.txtTime
        
    Set imgIE.Picture = LoadResPicture(1007, vbResBitmap)
    
    cboEfect.AddItem "Shrinking Box"
    cboEfect.AddItem "Growing Box"
    cboEfect.AddItem "Shrinking Circle"
    cboEfect.AddItem "Growing Circle"
    cboEfect.AddItem "Wipes Up"
    cboEfect.AddItem "Wipes Down"
    cboEfect.AddItem "Wipes Right"
    cboEfect.AddItem "Wipes Left"
    cboEfect.AddItem "Right Moving"
    cboEfect.AddItem "Downward Moving"
    cboEfect.AddItem "Right Moving Boxes"
    cboEfect.AddItem "Downward Moving Boxes"
    cboEfect.AddItem "Pixels 'Dissolve'"
    cboEfect.AddItem "Horizontal Curtain Closing"
    cboEfect.AddItem "Horizontal Curtain Opening"
    cboEfect.AddItem "Vertical Curtain Closing"
    cboEfect.AddItem "Vertical Curtain Opening"
    cboEfect.AddItem "Strips away previous screen going Left-Down"
    cboEfect.AddItem "Strips away previous screen going Left-Up"
    cboEfect.AddItem "Strips away previous screen going Right-Down"
    cboEfect.AddItem "Strips away previous screen going Right-Up"
    cboEfect.AddItem "Horizontal Bars 'Dissolve' Screen"
    cboEfect.AddItem "Vertical Bars 'Dissolve' Screen"
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmPageTran = Nothing
End Sub


