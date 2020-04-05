VERSION 5.00
Object = "{462EF1F4-16AF-444F-9DEE-F41BEBEC2FD8}#1.1#0"; "vbalODCL6.ocx"
Begin VB.Form frmMarkup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Markup Explorer"
   ClientHeight    =   4095
   ClientLeft      =   4680
   ClientTop       =   3660
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   273
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbohtml 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   105
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   45
      Width           =   3420
   End
   Begin ODCboLst6.OwnerDrawComboList lstObj 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   375
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3413
      ExtendedUI      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   4
      MaxLength       =   0
   End
   Begin ODCboLst6.OwnerDrawComboList lstEle 
      Height          =   1935
      Left            =   90
      TabIndex        =   2
      Top             =   2160
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3413
      ExtendedUI      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   4
      MaxLength       =   0
   End
End
Attribute VB_Name = "frmMarkup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()

    If WindowState <> vbMinimized Then
        On Error Resume Next
        Dim Top As Integer
        Top = 17
                        
        cbohtml.Move 0, Top, ScaleWidth
        lstObj.Move 0, (cbohtml.Height + 1 + Top), ScaleWidth, ScaleHeight / 2
        lstEle.Move 0, (cbohtml.Height + lstObj.Height + 1 + Top), ScaleWidth, ((ScaleHeight / 2) - cbohtml.Height) - Top
        Err = 0
    End If
    
End Sub


Public Sub Resize2()
    Form_Resize
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmMarkup = Nothing
End Sub


