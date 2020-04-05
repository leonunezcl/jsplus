VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmXmlExplorer 
   Caption         =   "XML Explorer"
   ClientHeight    =   8130
   ClientLeft      =   2835
   ClientTop       =   1950
   ClientWidth     =   7200
   Icon            =   "frmXmlExplorer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   7200
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   2250
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   13361
      _Version        =   393217
      Indentation     =   265
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   0
      Left            =   2985
      TabIndex        =   1
      Top             =   7650
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
End
Attribute VB_Name = "frmXmlExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Filename As String
Private Sub cmd_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

    On Error GoTo ErrHandler
    
    Dim x As XMLTree.XMLToTree
    
    util.CenterForm Me
    
    util.Hourglass hwnd, True
    
    Set x = New XMLTree.XMLToTree

    x.PlantTree TreeView1, Filename
    
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    
    cmd(0).Refresh
    
    Debug.Print "load"
    
    DrawXPCtl Me
    
    util.Hourglass hwnd, False
    
ErrHandler:
    util.Hourglass hwnd, False
    
End Sub


Private Sub Form_Resize()

    If WindowState <> vbMinimized Then
        TreeView1.Move 0, 0, Me.Width - 100, Me.Height - 1000
        cmd(0).Move Me.Width / 2 - 800, TreeView1.Height + 100
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmXmlExplorer = Nothing
End Sub


