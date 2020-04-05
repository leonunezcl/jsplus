VERSION 5.00
Begin VB.Form frmRunMacro 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Run Macro"
   ClientHeight    =   3585
   ClientLeft      =   5400
   ClientTop       =   2205
   ClientWidth     =   3855
   Icon            =   "frmRunMacro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   3210
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4410
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.ListBox lstMacro 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   75
      TabIndex        =   0
      Top             =   300
      Width           =   3705
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   1
      Left            =   2250
      TabIndex        =   2
      Top             =   3120
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Cancel"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePos      =   3
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   0
      Left            =   135
      TabIndex        =   3
      Top             =   3120
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Ok"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePos      =   3
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select a macro to run"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   75
      TabIndex        =   4
      Top             =   45
      Width           =   1515
   End
End
Attribute VB_Name = "frmRunMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub RunMacro()

    Dim Index As Integer
    'Dim Indice As Integer
    
    If lstMacro.ListIndex = -1 Then
        MsgBox "Select macro to run", vbCritical
        Exit Sub
    End If
    
    Index = lstMacro.ListIndex + 1
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
    Select Case Index
        Case 1
            frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro1
        Case 2
            frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro2
        Case 3
            frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro3
        Case 4
            frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro4
        Case 5
            frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro5
        Case 6
            frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro6
        Case 7
            frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro7
        Case 8
            frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro8
        Case 9
            frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro9
        Case 10
            frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro10
    End Select
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        RunMacro
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
    
    Dim k As Integer
        
    set_color_form Me
    'SetLayered hwnd, True
    
    For k = 1 To 10
        lstMacro.AddItem "Macro " & k
    Next k
    
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    cmd(0).Refresh
    cmd(1).Refresh
    
    Debug.Print "load"
    
    DrawXPCtl Me
            
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmRunMacro = Nothing
End Sub


