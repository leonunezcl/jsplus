VERSION 5.00
Begin VB.Form frmtrial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Welcome to JavaScript Plus!"
   ClientHeight    =   5250
   ClientLeft      =   2790
   ClientTop       =   3195
   ClientWidth     =   6645
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   Icon            =   "frmAcerca.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdComprar 
      Caption         =   "&Buy"
      Height          =   375
      Left            =   3720
      TabIndex        =   24
      Tag             =   "http://www.regnow.com/softsell/nph-softsell.cgi?item=12453-1"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1440
      TabIndex        =   23
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UNREGISTERED VERSION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4335
      TabIndex        =   22
      Top             =   990
      Width           =   2265
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Beta 1"
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
      Left            =   4245
      TabIndex        =   21
      Top             =   525
      Visible         =   0   'False
      Width           =   1725
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
      Left            =   4245
      TabIndex        =   20
      Top             =   300
      Width           =   1725
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
      Left            =   4080
      TabIndex        =   19
      Top             =   30
      Width           =   1920
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   7
      Left            =   6045
      Picture         =   "frmAcerca.frx":000C
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Santiago, Chile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   18
      Top             =   720
      Width           =   1305
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "VBSoftware"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   510
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   60
      Picture         =   "frmAcerca.frx":0316
      Top             =   945
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   16
      Left            =   60
      TabIndex        =   16
      Top             =   75
      Width           =   2910
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   60
      TabIndex        =   15
      Top             =   300
      Width           =   1605
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   6
      Left            =   585
      Picture         =   "frmAcerca.frx":03A8
      Top             =   3855
      Width           =   240
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Supports Windows Vista, Windows XP, 2000, NT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   915
      TabIndex        =   14
      Top             =   3870
      Width           =   3930
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   5
      Left            =   585
      Picture         =   "frmAcerca.frx":053A
      Top             =   3495
      Width           =   240
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Secure Online Order Form with SSL encryption via RegNow "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   13
      Left            =   915
      TabIndex        =   13
      Top             =   3495
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      X1              =   -30
      X2              =   6570
      Y1              =   4695
      Y2              =   4695
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   12
      Left            =   540
      TabIndex        =   12
      Top             =   4770
      Width           =   45
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Get the last version, updates or bugs fixed for FREE."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   915
      TabIndex        =   11
      Top             =   4410
      Width           =   3825
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   4
      Left            =   585
      Picture         =   "frmAcerca.frx":0636
      Top             =   4200
      Width           =   240
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Free Updates"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   915
      TabIndex        =   10
      Top             =   4200
      Width           =   1125
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order now from REGNOW"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3420
      MouseIcon       =   "frmAcerca.frx":0724
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Tag             =   "http://www.regnow.com/softsell/nph-softsell.cgi?item=12453-1"
      Top             =   1350
      Width           =   2085
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Get the full version immediately."
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
      Index           =   9
      Left            =   600
      TabIndex        =   8
      Top             =   1350
      Width           =   2760
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   8
      Left            =   780
      TabIndex        =   7
      Top             =   2865
      Width           =   45
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   3
      Left            =   585
      Picture         =   "frmAcerca.frx":0A2E
      Top             =   3165
      Width           =   240
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Unconditional 30-Day Money-Back Guarantee"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   915
      TabIndex        =   6
      Top             =   3165
      Width           =   3855
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "download immediately the full version after placing your order"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   915
      TabIndex        =   5
      Top             =   2850
      Width           =   4440
   End
   Begin VB.Image img 
      Height          =   210
      Index           =   2
      Left            =   585
      Picture         =   "frmAcerca.frx":0C87
      Top             =   2640
      Width           =   210
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Instant Delivery"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   915
      TabIndex        =   4
      Top             =   2640
      Width           =   1380
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "save time and increase productivity for just $30"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   915
      TabIndex        =   3
      Top             =   2355
      Width           =   3435
   End
   Begin VB.Image img 
      Height          =   210
      Index           =   1
      Left            =   585
      Picture         =   "frmAcerca.frx":105D
      Top             =   2130
      Width           =   210
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Best Value"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   915
      TabIndex        =   2
      Top             =   2130
      Width           =   885
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Get a powerfull, low cost and easy to use JavaScript Editor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   915
      TabIndex        =   1
      Top             =   1860
      Width           =   4245
   End
   Begin VB.Image img 
      Height          =   210
      Index           =   0
      Left            =   585
      Picture         =   "frmAcerca.frx":1433
      Top             =   1665
      Width           =   210
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Right Choice"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   915
      TabIndex        =   0
      Top             =   1665
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   0
      Top             =   0
      Width           =   6645
   End
End
Attribute VB_Name = "frmtrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
    Unload Me
End Sub


Private Sub cmdComprar_Click()
    util.ShellFunc cmdComprar.Tag, vbNormalFocus
End Sub



Private Sub Form_Load()
    
    util.CenterForm Me
    
    Image1.Picture = LoadResPicture(1003, vbResBitmap)
        
    lblv.Caption = App.Major & "." & App.Minor & "." & App.Revision
        
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmtrial = Nothing
End Sub


Private Sub lblURL_Click()
    util.ShellFunc lblURL.Tag, vbNormalFocus
End Sub




