VERSION 5.00
Begin VB.Form frmLanMan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Languages"
   ClientHeight    =   2805
   ClientLeft      =   4485
   ClientTop       =   3525
   ClientWidth     =   4095
   Icon            =   "frmLanMan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   2610
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5070
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.ListBox lstFiles 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   15
      TabIndex        =   1
      Top             =   300
      Width           =   2520
   End
   Begin jsplus.MyButton cmdAdd 
      Height          =   405
      Left            =   2595
      TabIndex        =   2
      Top             =   315
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Add"
      AccessKey       =   "A"
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
   Begin jsplus.MyButton cmdRem 
      Height          =   405
      Left            =   2595
      TabIndex        =   4
      Top             =   1245
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Remove"
      AccessKey       =   "R"
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
   Begin jsplus.MyButton cmdEdit 
      Height          =   405
      Left            =   2595
      TabIndex        =   3
      Top             =   780
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Edit"
      AccessKey       =   "E"
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
   Begin jsplus.MyButton cmdExit 
      Height          =   405
      Left            =   2595
      TabIndex        =   5
      Top             =   1710
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "E&xit"
      AccessKey       =   "x"
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Installed Files"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   945
   End
End
Attribute VB_Name = "frmLanMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cargar_lenguajes()

    Dim archivo As String
    Dim k As Integer
    
    lstFiles.Clear
    
    archivo = Dir(App.path & "\languages\")
    k = 1
    Do Until archivo = ""
        If LCase$(VBA.Right$(archivo, 4)) = ".def" Then
            lstFiles.AddItem archivo
        End If
        archivo = Dir()
    Loop
    
End Sub

Private Sub cmdAdd_Click()

    frmEdtLan.Show vbModal
    
End Sub

Private Sub cmdEdit_Click()

    If lstFiles.ListIndex <> -1 Then
        frmEdtLan.file = lstFiles.Text
        frmEdtLan.Show vbModal
    End If
    
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdRem_Click()

    If lstFiles.ListIndex <> -1 Then
        If Confirma("Are you sure to remove this file") = vbYes Then
            util.Hourglass hwnd, True
            util.BorrarArchivo util.StripPath(App.path) & "languages\" & lstFiles.Text
            Call cargar_lenguajes
            ListaLangs.Load
            util.Hourglass hwnd, False
        End If
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
    
    util.Hourglass hwnd, True
    
    set_color_form Me
    
    'SetLayered hwnd, True
    
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    
    Call cargar_lenguajes
    
    cmdAdd.Refresh
    cmdEdit.Refresh
    cmdRem.Refresh
    cmdExit.Refresh
    
    Debug.Print "load : " & Me.Name
    
    DrawXPCtl Me
    
    util.Hourglass hwnd, False
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Debug.Print "unload :" & Me.Name
    Set frmLanMan = Nothing
    
End Sub


