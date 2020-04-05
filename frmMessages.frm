VERSION 5.00
Object = "{5F37140E-C836-11D2-BEF8-525400DFB47A}#1.1#0"; "vbalTab6.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmMessages 
   Caption         =   "Output"
   ClientHeight    =   2010
   ClientLeft      =   1440
   ClientTop       =   3390
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   4995
   Begin vbalTabStrip6.TabControl tabMsg 
      Height          =   1695
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   2990
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabAlign        =   2
      HotTrack        =   -1  'True
      FlatButtons     =   -1  'True
      CoolTabs        =   1
      Begin vbAcceleratorSGrid6.vbalGrid griErr 
         Height          =   915
         Left            =   2895
         TabIndex        =   2
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1614
         NoHorizontalGridLines=   -1  'True
         NoVerticalGridLines=   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         DisableIcons    =   -1  'True
      End
      Begin vbAcceleratorSGrid6.vbalGrid griMsg 
         Height          =   885
         Left            =   195
         TabIndex        =   1
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1561
         NoHorizontalGridLines=   -1  'True
         NoVerticalGridLines=   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         DisableIcons    =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Resize2()
    Form_Resize
End Sub


Private Sub Form_Load()

    With tabMsg
        .ImageList = frmMain.m_MainImg.hIml
        .AddTab "Output", 0, , "Output"
        .AddTab "Errors", 1, , "Errors"
        .TabAlign = etaBottom
        .Rebuild
    End With

    griMsg.ZOrder 0
End Sub


Private Sub Form_Resize()
    
    If WindowState <> vbMinimized Then
        Dim clefT As Integer
        clefT = 270
        tabMsg.Move clefT, 0, ScaleWidth - clefT, ScaleHeight
        griMsg.Move 10, 0, ScaleWidth - 10, ScaleHeight - 350
        griErr.Move 10, 0, ScaleWidth - 10, ScaleHeight - 350
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmMessages = Nothing
End Sub


Private Sub tabMsg_TabClick(ByVal lTab As Long)

    If lTab = 1 Then
        griMsg.ZOrder 0
    ElseIf lTab = 2 Then
        griErr.ZOrder 0
    End If
    
End Sub


