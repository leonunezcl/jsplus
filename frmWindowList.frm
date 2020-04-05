VERSION 5.00
Begin VB.Form frmWindowList 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Window List"
   ClientHeight    =   4620
   ClientLeft      =   5220
   ClientTop       =   3420
   ClientWidth     =   7470
   Icon            =   "frmWindowList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   5205
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5730
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.ListBox lstWin 
      Height          =   4545
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   5820
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   0
      Left            =   5940
      TabIndex        =   1
      Top             =   45
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
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   1
      Left            =   5940
      TabIndex        =   3
      Top             =   510
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
End
Attribute VB_Name = "frmWindowList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub activa_ventana()

    If lstWin.ListIndex = -1 Then
        Exit Sub
    End If
        
    Dim frm As Form
    Dim file As New cFile
    Dim k As Integer
    
    For k = 1 To Files.count
        If Files.Files(k).Filename = lstWin.List(lstWin.ListIndex) Then
            For Each frm In Forms
                If TypeName(frm) = "frmEdit" Then
                    If CInt(frm.tag) = Files.Files(k).IdDoc Then
                        Call frmMain.activar_editor(frm.hwnd)
                        GoTo salir
                    End If
                End If
            Next
        End If
    Next k
    
    Exit Sub
salir:
    Unload Me
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call activa_ventana
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Dim file As New cFile
    Dim k As Integer
    
    util.CenterForm Me
    
    For k = 1 To Files.count
        Set file = New cFile
        Set file = Files.Files(k)
        If Len(file.Filename) > 0 Then
            lstWin.AddItem file.Filename
        End If
        Set file = Nothing
    Next k
    
    util.CenterForm Me
    
    set_color_form Me
    
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    cmd(0).Refresh
    cmd(0).Refresh
    
    DrawXPCtl Me
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmWindowList = Nothing
End Sub


