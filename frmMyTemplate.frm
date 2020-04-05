VERSION 5.00
Object = "{FCFAF346-DE8A-4FB6-8612-5000548EFDC7}#2.0#0"; "vbsListView6.ocx"
Begin VB.Form frmMyTemplate 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My Template"
   ClientHeight    =   2715
   ClientLeft      =   2910
   ClientTop       =   3045
   ClientWidth     =   6120
   Icon            =   "frmMyTemplate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   4455
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6105
      Visible         =   0   'False
      Width           =   2250
   End
   Begin vbalListViewLib6.vbalListViewCtl lvwTem 
      Height          =   2310
      Left            =   45
      TabIndex        =   0
      Top             =   225
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   4075
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      View            =   1
      MultiSelect     =   -1  'True
      LabelEdit       =   0   'False
      AutoArrange     =   0   'False
      Appearance      =   0
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   0
      Left            =   4590
      TabIndex        =   1
      Top             =   195
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Save"
      AccessKey       =   "S"
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
      Left            =   4590
      TabIndex        =   2
      Top             =   675
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&New"
      AccessKey       =   "N"
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
      Index           =   2
      Left            =   4590
      TabIndex        =   3
      Top             =   1155
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
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   3
      Left            =   4590
      TabIndex        =   4
      Top             =   1620
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
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   4
      Left            =   4590
      TabIndex        =   5
      Top             =   2115
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
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Available Templates"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   45
      TabIndex        =   6
      Top             =   15
      Width           =   1425
   End
End
Attribute VB_Name = "frmMyTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyTemFile As String
Private Sub cargar_templates()

    Dim arr_sec() As String
    Dim k As Integer
    
    get_info_section "mytemplates", arr_sec, MyTemFile
    
    For k = 1 To UBound(arr_sec)
        lvwTem.ListItems.Add , , Explode(arr_sec(1), 1, "|") 'nombre
        lvwTem.ListItems(k).SubItems(1).Caption = Explode(arr_sec(1), 2, "|") 'archivo
    Next k
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        'open
    ElseIf Index = 1 Then
        'new
    ElseIf Index = 2 Then
        'edit
    ElseIf Index = 3 Then
        'remove
    ElseIf Index = 4 Then
        'exit
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Util.CenterForm Me
    
    Util.Hourglass hwnd, True
    
    set_color_form Me
    
    With lvwTem
        .Columns.Add , "k1", "Name", , 2000
        .Columns.Add , "k2", "File", , 2400
    End With
    
    MyTemFile = Util.StripPath(App.Path) & "mytemplates.ini"
    
    Call cargar_templates
    
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    cmd(0).Refresh
    cmd(1).Refresh
    cmd(2).Refresh
    cmd(3).Refresh
    cmd(4).Refresh
    
    Debug.Print "load"
    
    DrawXPCtl Me
    
    Util.Hourglass hwnd, False
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmMyTemplate = Nothing
End Sub


