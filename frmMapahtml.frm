VERSION 5.00
Object = "{283862C9-F6FD-4704-8691-3487E0FDCFFD}#1.0#0"; "MyButton.ocx"
Begin VB.Form frmMapa 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Symbol"
   ClientHeight    =   6540
   ClientLeft      =   2475
   ClientTop       =   2610
   ClientWidth     =   7455
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   ShowInTaskbar   =   0   'False
   Begin vb6projectMyButton.MyButton cmd 
      Height          =   450
      Index           =   0
      Left            =   5970
      TabIndex        =   10
      Top             =   1290
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   794
      SPN             =   "MyButtonDefSkin"
      Text            =   "MyButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   2610
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3585
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Frame fra 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Character List"
      ForeColor       =   &H00FF8080&
      Height          =   2550
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   240
      Width           =   5745
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   8
         Top             =   270
         Width           =   255
      End
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   6495
   End
   Begin vb6projectMyButton.MyButton cmd 
      Height          =   450
      Index           =   1
      Left            =   5985
      TabIndex        =   11
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   794
      SPN             =   "MyButtonDefSkin"
      Text            =   "MyButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Double click to insert in active document"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   75
      TabIndex        =   6
      Top             =   45
      Width           =   3510
   End
   Begin VB.Label lblcodigo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ALT"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3465
      TabIndex        =   5
      Top             =   3045
      Width           =   300
   End
   Begin VB.Label lblDescrip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIP"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   3045
      Width           =   705
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Alt+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   3465
      TabIndex        =   3
      Top             =   2805
      Width           =   345
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   2805
      Width           =   975
   End
   Begin VB.Label lblchar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   5925
      TabIndex        =   1
      Top             =   330
      Width           =   1395
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public tipo_insercion As Integer

Private ultima As Integer

Private Type eHtml
    traduccion As String
    caracter As String
    ayuda As String
End Type
Private arr_html() As eHtml
Private Sub CargarMapa()

    Dim k As Integer
    Dim j As Integer
    Dim t As Integer
    Dim c As Integer
    Dim l As Integer
    Dim p As Integer
    Dim num
    Dim Linea As String
    
    num = Util.LeeIni(IniPath, "mapahtml", "num")
    
    If num = "" Or Not IsNumeric(num) Then
        MsgBox "Error on file " & App.Title & ".ini"
        Exit Sub
    End If
    
    ReDim arr_html(0)
    c = 0
    For k = 1 To CInt(num)
        Linea = Util.LeeIni(IniPath, "mapahtml", "car" & k)
        
        If Len(Linea) > 0 Then
            ReDim Preserve arr_html(c)
            arr_html(c).traduccion = Util.Explode(Linea, 1, ";")
            arr_html(c).caracter = Util.Explode(Linea, 2, ";")
            arr_html(c).ayuda = Util.Explode(Linea, 3, ";")
            c = c + 1
        End If
    Next k
    
    t = pic(0).Top
    l = pic(0).Left
    c = 0
    
    For j = 0 To UBound(arr_html)
        
        If j > 0 Then
            Load pic(j)
        End If
        
        If c = 0 Then
            pic(j).Left = pic(0).Left
        Else
            pic(j).Left = pic(c - 1).Left + pic(j - 1).Height
        End If
        
        pic(j).Top = t
        pic(j).Width = pic(0).Width
        pic(j).Visible = True
        pic(j).CurrentX = 60
        pic(j).CurrentY = 3
        pic(j).Font.Size = 10
        
        pic(j).Print arr_html(j).caracter
        pic(j).tag = c & "#" & arr_html(j).caracter & "#" & arr_html(j).traduccion & "#" & arr_html(j).ayuda
        
        If c > 19 Then
            c = 0
            t = t + pic(0).Height
        Else
            c = c + 1
        End If
    Next j
    
End Sub


Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        'If tipo_insercion = 0 Then
        '    Call InsertarHtml(lblchar.Caption)
        'Else
   '         Call InsertarHtml(lblcodigo.Caption & ";")
        'End If
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
    
    Util.Hourglass hwnd, True
    
    Util.CenterForm Me
    
    ultima = -1
    
    lblDescrip.Caption = ""
    lblcodigo.Caption = ""
    lblchar.Caption = ""
    
    Call CargarMapa
    
    MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    cmd(0).Refresh
    cmd(1).Refresh
    
    Util.Hourglass hwnd, False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMapa = Nothing
End Sub


Private Sub pic_DblClick(Index As Integer)
    cmd_Click 0
End Sub

Private Sub pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim num As Integer
    
    If ultima = Index Then Exit Sub
    
    If ultima > -1 Then
        pic(ultima).BorderStyle = 0
        pic(ultima).Print pic(Index).tag
    End If
        
    pic(Index).BorderStyle = 1
    
    pic(Index).Print Util.Explode(pic(Index).tag, 2, "#")
    
    If InStr(Util.Explode(pic(Index).tag, 2, "#"), "&") Then
        lblchar.Caption = "&" & Util.Explode(pic(Index).tag, 2, "#")
    Else
        lblchar.Caption = Util.Explode(pic(Index).tag, 2, "#")
    End If
    
    lblDescrip.Caption = Util.Explode(pic(Index).tag, 4, "#")
    lblcodigo.Caption = Util.Explode(pic(Index).tag, 3, "#")
    ultima = Index
    
End Sub


