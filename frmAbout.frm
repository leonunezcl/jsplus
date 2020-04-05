VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   5340
   ClientLeft      =   3480
   ClientTop       =   3135
   ClientWidth     =   6495
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   4200
      Left            =   0
      ScaleHeight     =   4170
      ScaleWidth      =   1080
      TabIndex        =   9
      Top             =   0
      Width           =   1110
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.vbsoftware.cl"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4035
      MouseIcon       =   "frmAbout.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Tag             =   "http://www.vbsoftware.cl"
      Top             =   3525
      Width           =   1890
   End
   Begin VB.Label lblsupport 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mailto:support@vbsoftware.cl"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4035
      MouseIcon       =   "frmAbout.frx":0316
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Tag             =   "mailto:support@vbsoftware.cl"
      Top             =   3780
      Width           =   2175
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0620
      ForeColor       =   &H00C00000&
      Height          =   1020
      Index           =   1
      Left            =   135
      TabIndex        =   10
      Top             =   4305
      Width           =   4800
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line2 
      X1              =   1110
      X2              =   6435
      Y1              =   4185
      Y2              =   4185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This software is registered to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1305
      TabIndex        =   8
      Top             =   2925
      Width           =   2535
   End
   Begin VB.Label lblv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3270
      TabIndex        =   7
      Top             =   1110
      Width           =   1725
   End
   Begin VB.Label lblreg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special Edition for Centennial College - Toronto Ontario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1305
      TabIndex        =   6
      Top             =   3135
      Width           =   4755
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   7
      Left            =   3135
      Picture         =   "frmAbout.frx":0742
      Top             =   1260
      Width           =   480
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "JavaScript Plus!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   17
      Left            =   3105
      TabIndex        =   5
      Top             =   855
      Width           =   1920
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Santiago, Chile"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   3690
      TabIndex        =   4
      Top             =   2355
      Width           =   1065
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "VBSoftware"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   3690
      TabIndex        =   3
      Top             =   2145
      Width           =   840
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   3690
      Picture         =   "frmAbout.frx":0A4C
      Top             =   2595
      Width           =   315
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2002-2009 Luis Nunez"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Index           =   16
      Left            =   3690
      TabIndex        =   2
      Top             =   1545
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "All rights reserved."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   15
      Left            =   3690
      TabIndex        =   1
      Top             =   1935
      Width           =   1365
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1755
      Left            =   1305
      Top             =   870
      Width           =   1680
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Student Edition"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   3690
      TabIndex        =   0
      Top             =   1320
      Width           =   1290
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Dim fi As String
    Dim ini As String
    Dim us As String
    Dim na As String
    
    util.CenterForm Me
        
    'SetLayered hwnd, True
    
    If tipo_version = 1 Then
        lbl(0).Caption = "Developer Edition"
        Set Image1.Picture = LoadResPicture(1006, vbResBitmap)
    ElseIf tipo_version = 2 Then    'comercial
        lbl(0).Caption = "Commercial Edition"
        Set Image1.Picture = LoadResPicture(1005, vbResBitmap)
    ElseIf tipo_version = 3 Then    'educacional
        lbl(0).Caption = "Student Edition"
        Set Image1.Picture = LoadResPicture(1004, vbResBitmap)
    ElseIf tipo_version = 4 Then        'testing
        lbl(0).Caption = "Testing Edition"
        Set Image1.Picture = LoadResPicture(1006, vbResBitmap)
    End If
            
    lblv.Caption = App.Major & "." & App.Minor & "." & App.Revision
    
    fi = Base64Encode(Chr$(114) & Chr$(101) & Chr$(103) & Chr$(117) & Chr$(115) & Chr$(101) & Chr$(114) & Chr$(46) & Chr$(105) & Chr$(110) & Chr$(105))
    ini = Base64Encode(util.StripPath(App.Path)) & fi
    
    If tipo_version < 4 Then        '
        If ArchivoExiste2(Base64Decode(ini)) Then
            us = Chr$(117) & Chr$(115) & Chr$(101) & Chr$(114)
            na = Chr$(110) & Chr$(97) & Chr$(109) & Chr$(101)
            lblreg.Caption = util.LeeIni(Base64Decode(ini), us, na)
        End If
    Else
        lblreg.Caption = "Only for beta testers"
    End If
        
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmAbout = Nothing
End Sub


Private Sub lblsupport_Click()
    util.ShellFunc lblsupport.Tag, vbNormalFocus
End Sub

Private Sub lblURL_Click()
    util.ShellFunc lblURL.Tag, vbNormalFocus
End Sub


