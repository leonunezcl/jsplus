VERSION 5.00
Object = "{E7106799-3A07-4335-80BA-4F20E8E5E2E9}#2.0#0"; "vbsODCL6.ocx"
Begin VB.UserControl vbsXHTML 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   45
      ScaleHeight     =   255
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   300
      Width           =   3405
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "document"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   855
      End
   End
   Begin ODCboLst6.OwnerDrawComboList lstObj 
      Height          =   1695
      Left            =   75
      TabIndex        =   2
      ToolTipText     =   "Double clic to insert in active document"
      Top             =   645
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   2990
      Sorted          =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      ClientDraw      =   1
      Style           =   4
      MaxLength       =   0
   End
End
Attribute VB_Name = "vbsXHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Img As cVBALImageList
Public Event ItemSelected(ByVal Atributo As String)
Public Sub Load()

    Dim Archivo As String
    Dim nFreeFile As String
    Dim k As Integer
    Dim linea As String
    
    Archivo = util.StripPath(App.Path) & "config\xhtml.ini"
    
    BuildImageList
    
    ReDim arr_data_xml(0)
    k = 0
    nFreeFile = FreeFile
    
    lstObj.ImageList = m_Img.hIml
    
    Open Archivo For Input As #nFreeFile
        Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, linea
            lstObj.AddItemAndData Explode(linea, 1, "|"), 0
            
            If k > 0 Then
                ReDim Preserve arr_data_xml(k)
            End If
            arr_data_xml(k).Tag = Explode(linea, 1, "|")
            arr_data_xml(k).help = Explode(linea, 2, "|")
            k = k + 1
        Loop
    Close #nFreeFile
    
    lstObj.ListIndex = 0
    
End Sub



Private Sub BuildImageList()
    
    Set m_Img = New cVBALImageList
    
    With m_Img
        .IconSizeX = 16: .IconSizeY = 16: .ColourDepth = ILC_COLOR24
        .Create
        .AddFromResourceID 244, App.hInstance, IMAGE_ICON, "k1"
    End With
   
End Sub

Private Sub lstObj_Click()

    If lstObj.ListIndex <> -1 Then
        lbl.Caption = lstObj.Text
        lbl.Caption = arr_data_xml(lstObj.ListIndex).Tag
    End If
        
End Sub


Private Sub lstObj_DblClick()

    RaiseEvent ItemSelected(arr_data_xml(lstObj.ListIndex).help)
    
End Sub


Private Sub UserControl_Resize()

    On Error Resume Next
    
    pic.Move 0, 0, UserControl.Width - 15
    lstObj.Move 0, pic.Height + 1, pic.Width, UserControl.Height - pic.Height - 1
        
    Err = 0
    
End Sub


