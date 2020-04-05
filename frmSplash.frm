VERSION 5.00
Object = "{866F095F-113F-4DC1-B803-F4CF4AFC96EE}#1.0#0"; "vbspgbbar.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5400
   ClientLeft      =   5295
   ClientTop       =   3405
   ClientWidth     =   5880
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin vbsprgbar.ucProgressBar pgb 
      Height          =   225
      Left            =   495
      TabIndex        =   5
      Top             =   4830
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   12937777
   End
   Begin VB.Label lblv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "VERSION"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   30
      TabIndex        =   7
      Top             =   5190
      Width           =   720
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Starting ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   5130
      Width           =   5730
   End
   Begin VB.Label lble 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Santiago de Chile"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   4095
      TabIndex        =   4
      Top             =   1050
      Width           =   1245
   End
   Begin VB.Label lble 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   4095
      TabIndex        =   3
      Top             =   840
      Width           =   840
   End
   Begin VB.Label lble 
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
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   3
      Left            =   4095
      TabIndex        =   2
      Top             =   225
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label lble 
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   4095
      TabIndex        =   1
      Top             =   630
      Width           =   1365
   End
   Begin VB.Label lble 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Evaluation Edition"
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
      Index           =   1
      Left            =   4095
      TabIndex        =   0
      Top             =   15
      Width           =   1500
   End
   Begin VB.Image imgSplash 
      Height          =   4725
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Top             =   0
      Width           =   5835
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
    
    Dim k As Integer
    
    util.CenterForm Me
    util.Hourglass hwnd, True
    'SetLayered hwnd, True
    
    #If LITE = 1 Then
        'Me.Width = 5800
    #Else
        
        For k = 0 To 4
            lble(k).Visible = False
        Next k
                
    #End If
    
    util.CenterForm Me
    
    'Me.imgSplash = LoadResPicture(1001, vbResBitmap)
    
    lblv.Caption = "V." & App.Major & "." & App.Minor & "." & App.Revision & "-Jan 2009"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
    util.Hourglass hwnd, False
    Set imgSplash = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmSplash = Nothing
End Sub




