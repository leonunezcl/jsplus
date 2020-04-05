VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   3315
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTF1 
      Height          =   945
      Left            =   150
      TabIndex        =   0
      Tag             =   $"frmMain.frx":0000
      Top             =   240
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1667
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":008B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfTmp 
      Height          =   510
      Left            =   1620
      TabIndex        =   1
      Top             =   255
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   900
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":0102
   End
   Begin CodeSenseCtl.CodeSense txtCode 
      Height          =   1095
      Left            =   1560
      OleObjectBlob   =   "frmMain.frx":0184
      TabIndex        =   2
      Top             =   1065
      Width           =   1620
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "html"
      DialogTitle     =   "WonderHTML"
      Filter          =   $"frmMain.frx":02EA
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FileName As String
Private Sub Form_Load()

    If Len(FileName) > 0 Then
        txtCode.OpenFile FileName
    End If
    
    RTF1.Font.Name = txtCode.Font.Name
    RTF1.Font.Size = txtCode.Font.Size
    RTF1.SelIndent = 45 'just a little
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmMain = Nothing
End Sub


