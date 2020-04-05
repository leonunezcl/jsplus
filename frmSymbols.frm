VERSION 5.00
Begin VB.Form frmSymbols 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Icon Help"
   ClientHeight    =   4230
   ClientLeft      =   4170
   ClientTop       =   5430
   ClientWidth     =   6360
   ControlBox      =   0   'False
   Icon            =   "frmSymbols.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   2430
      TabIndex        =   15
      Top             =   3795
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "JavaScript Plus! Icon Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   6285
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CSS"
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
         Index           =   2
         Left            =   4905
         TabIndex        =   14
         Top             =   285
         Width           =   375
      End
      Begin VB.Image img 
         Height          =   315
         Index           =   10
         Left            =   2730
         Top             =   1485
         Width           =   315
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "DHTML Event"
         Height          =   195
         Index           =   10
         Left            =   3285
         TabIndex        =   13
         Top             =   1590
         Width           =   1035
      End
      Begin VB.Image img 
         Height          =   315
         Index           =   9
         Left            =   2730
         Top             =   1050
         Width           =   315
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Attribute"
         Height          =   195
         Index           =   9
         Left            =   3285
         TabIndex        =   12
         Top             =   1215
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "HTML"
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
         Index           =   1
         Left            =   2640
         TabIndex        =   11
         Top             =   285
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "JavaScript"
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
         Index           =   0
         Left            =   105
         TabIndex        =   10
         Top             =   285
         Width           =   915
      End
      Begin VB.Image img 
         Height          =   315
         Index           =   8
         Left            =   180
         Top             =   3180
         Width           =   315
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Object"
         Height          =   195
         Index           =   8
         Left            =   720
         TabIndex        =   9
         Top             =   3315
         Width           =   465
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "CSS Tag"
         Height          =   195
         Index           =   7
         Left            =   5490
         TabIndex        =   8
         Top             =   675
         Width           =   645
      End
      Begin VB.Image img 
         Height          =   315
         Index           =   7
         Left            =   4980
         Top             =   525
         Width           =   315
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "HTML Tag"
         Height          =   195
         Index           =   6
         Left            =   3285
         TabIndex        =   7
         Top             =   810
         Width           =   780
      End
      Begin VB.Image img 
         Height          =   315
         Index           =   6
         Left            =   2730
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Event"
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   6
         Top             =   2805
         Width           =   420
      End
      Begin VB.Image img 
         Height          =   315
         Index           =   5
         Left            =   180
         Top             =   2685
         Width           =   315
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Constant"
         Height          =   195
         Index           =   4
         Left            =   735
         TabIndex        =   5
         Top             =   2370
         Width           =   630
      End
      Begin VB.Image img 
         Height          =   315
         Index           =   4
         Left            =   180
         Top             =   2250
         Width           =   315
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Collection"
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   4
         Top             =   1935
         Width           =   690
      End
      Begin VB.Image img 
         Height          =   315
         Index           =   3
         Left            =   180
         Top             =   1815
         Width           =   315
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Method/Function"
         Height          =   195
         Index           =   2
         Left            =   735
         TabIndex        =   3
         Top             =   1515
         Width           =   1230
      End
      Begin VB.Image img 
         Height          =   315
         Index           =   2
         Left            =   180
         Top             =   1395
         Width           =   315
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Property"
         Height          =   195
         Index           =   1
         Left            =   735
         TabIndex        =   2
         Top             =   1095
         Width           =   585
      End
      Begin VB.Image img 
         Height          =   315
         Index           =   1
         Left            =   180
         Top             =   990
         Width           =   315
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "JavaScript Object"
         Height          =   195
         Index           =   0
         Left            =   735
         TabIndex        =   1
         Top             =   720
         Width           =   1260
      End
      Begin VB.Image img 
         Height          =   315
         Index           =   0
         Left            =   180
         Top             =   600
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

    util.Hourglass hwnd, True
    
    util.CenterForm Me
        
    'javascript object
    img(0).Picture = LoadResPicture(199, vbResIcon)
    
    'javascript property
    img(1).Picture = LoadResPicture(193, vbResIcon)
    img(9).Picture = LoadResPicture(193, vbResIcon)
    
    'javascript method
    img(2).Picture = LoadResPicture(191, vbResIcon)
    
    'javascript collection
    img(3).Picture = LoadResPicture(253, vbResIcon)
    
    'javascript constant
    img(4).Picture = LoadResPicture(194, vbResIcon)
    
    'event
    img(5).Picture = LoadResPicture(195, vbResIcon)
    img(10).Picture = LoadResPicture(254, vbResIcon)
    
    'html tag
    img(6).Picture = LoadResPicture(200, vbResIcon)
    
    'html tag
    img(7).Picture = LoadResPicture(264, vbResIcon)
    
    'object
    img(8).Picture = LoadResPicture(263, vbResIcon)
    
    util.Hourglass hwnd, False
        
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmSymbols = Nothing
End Sub


