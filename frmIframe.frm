VERSION 5.00
Begin VB.Form frmIframe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IFRAME Wizard"
   ClientHeight    =   5160
   ClientLeft      =   5535
   ClientTop       =   2670
   ClientWidth     =   4350
   Icon            =   "frmIframe.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   24
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   23
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   4200
      Index           =   3
      Left            =   90
      TabIndex        =   9
      Top             =   75
      Width           =   4185
      Begin VB.Frame fra 
         Caption         =   "Settings"
         Height          =   750
         Index           =   0
         Left            =   75
         TabIndex        =   17
         Top             =   1095
         Width           =   3975
         Begin VB.TextBox txtH 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   705
            MaxLength       =   5
            TabIndex        =   2
            Top             =   285
            Width           =   915
         End
         Begin VB.TextBox txtW 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2580
            MaxLength       =   5
            TabIndex        =   3
            Top             =   270
            Width           =   915
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Height"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   19
            Top             =   285
            Width           =   465
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Width"
            Height          =   195
            Index           =   1
            Left            =   2025
            TabIndex        =   18
            Top             =   285
            Width           =   420
         End
      End
      Begin VB.TextBox txtUrl 
         Height          =   285
         Left            =   75
         TabIndex        =   1
         Top             =   780
         Width           =   3960
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1035
         TabIndex        =   0
         Top             =   210
         Width           =   2985
      End
      Begin VB.Frame fra 
         Caption         =   "Spacing"
         Height          =   705
         Index           =   1
         Left            =   75
         TabIndex        =   14
         Top             =   1905
         Width           =   3975
         Begin VB.TextBox txtHS 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   975
            MaxLength       =   5
            TabIndex        =   4
            Top             =   225
            Width           =   915
         End
         Begin VB.TextBox txtVS 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2580
            MaxLength       =   5
            TabIndex        =   5
            Top             =   225
            Width           =   915
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Horizontal"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   16
            Top             =   285
            Width           =   705
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Vertical"
            Height          =   195
            Index           =   3
            Left            =   1995
            TabIndex        =   15
            Top             =   285
            Width           =   525
         End
      End
      Begin VB.Frame fra 
         Caption         =   "Extra"
         Height          =   1410
         Index           =   2
         Left            =   75
         TabIndex        =   10
         Top             =   2670
         Width           =   3990
         Begin VB.ComboBox cboBorder 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   255
            Width           =   2955
         End
         Begin VB.ComboBox cboAlign 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   615
            Width           =   2955
         End
         Begin VB.ComboBox cboScr 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   975
            Width           =   2955
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Border"
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   13
            Top             =   255
            Width           =   465
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Align"
            Height          =   195
            Index           =   7
            Left            =   135
            TabIndex        =   12
            Top             =   615
            Width           =   345
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Scrollbars"
            Height          =   195
            Index           =   8
            Left            =   135
            TabIndex        =   11
            Top             =   975
            Width           =   690
         End
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "URL of page to appear in frame"
         Height          =   195
         Index           =   4
         Left            =   75
         TabIndex        =   21
         Top             =   570
         Width           =   2235
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Iframe Name"
         Height          =   195
         Index           =   5
         Left            =   75
         TabIndex        =   20
         Top             =   255
         Width           =   900
      End
   End
   Begin VB.Image imgOP 
      Height          =   255
      Left            =   3975
      Top             =   4815
      Width           =   300
   End
   Begin VB.Image imgNE 
      Height          =   255
      Left            =   3645
      Top             =   4815
      Width           =   300
   End
   Begin VB.Image imgFX 
      Height          =   255
      Left            =   3330
      Top             =   4815
      Width           =   300
   End
   Begin VB.Image imgIE 
      Height          =   255
      Left            =   3015
      Top             =   4815
      Width           =   300
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Browser Compatibility"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   1410
      TabIndex        =   22
      Top             =   4830
      Width           =   1485
   End
End
Attribute VB_Name = "frmIframe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function creariframe() As Boolean

    Dim src As New cStringBuilder
    Dim fb As String
    
    If cboBorder.ListIndex = 0 Then
        fb = "0"
    Else
        fb = "1"
    End If
    
    src.Append "<html>" & vbNewLine
    src.Append "<head>" & vbNewLine
    src.Append "<meta http-equiv=" & Chr$(34) & "Content-Type" & Chr$(34) & "CONTENT=" & Chr$(34) & "text/html;" & Chr$(34) & ">" & vbNewLine
    src.Append "<title>Testing Page</title>" & vbNewLine
    src.Append "</head>" & vbNewLine
    src.Append "<body>" & vbNewLine
    src.Append "<iframe name=" & Chr$(34) & txtName.Text & Chr$(34) & " height=" & Chr$(34) & txtH.Text & Chr$(34) & " width=" & Chr$(34) & txtW.Text & Chr$(34) & _
               " src=" & Chr$(34) & txtUrl.Text & Chr$(34) & " border=" & Chr$(34) & fb & Chr$(34) & " frameborder=" & Chr$(34) & fb & Chr$(34) & _
               " scrolling=" & Chr$(34) & cboScr.Text & Chr$(34) & " align=" & Chr$(34) & cboAlign.Text & Chr$(34) & _
               " hspace=" & Chr$(34) & txtHS.Text & Chr$(34) & " vspace=" & Chr$(34) & txtVS.Text & Chr$(34) & _
               "></iframe>" & vbNewLine
    src.Append "</body>" & vbNewLine
    src.Append "</html>" & vbNewLine
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(src.ToString)
    End If
    
    Set src = Nothing
    
    creariframe = True
    
End Function

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If creariframe() Then
            Unload Me
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
        
    util.SetNumber txtH.hwnd
    util.SetNumber txtW.hwnd
    util.SetNumber txtHS.hwnd
    util.SetNumber txtVS.hwnd
    
    cboBorder.AddItem "no"
    cboBorder.AddItem "yes"
    cboBorder.ListIndex = 0
    
    cboAlign.AddItem "left"
    cboAlign.AddItem "center"
    cboAlign.AddItem "right"
    cboAlign.AddItem "top"
    cboAlign.AddItem "middle"
    cboAlign.AddItem "bottom"
    cboAlign.ListIndex = 0
    
    cboScr.AddItem "no"
    cboScr.AddItem "yes"
    cboScr.AddItem "auto"
    cboScr.ListIndex = 2
    
    Set imgIE.Picture = LoadResPicture(1007, vbResBitmap)
    Set imgFX.Picture = LoadResPicture(1008, vbResBitmap)
    Set imgNE.Picture = LoadResPicture(1009, vbResBitmap)
    Set imgOP.Picture = LoadResPicture(1010, vbResBitmap)
    
    Debug.Print "load : " & Me.Name
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload : " & Me.Name
    Set frmIframe = Nothing
End Sub


