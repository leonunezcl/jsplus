VERSION 5.00
Object = "{E7106799-3A07-4335-80BA-4F20E8E5E2E9}#2.0#0"; "vbsODCL6.ocx"
Object = "{98F993CC-3598-405A-9E9A-0D2CF198B250}#2.0#0"; "vbsDkTb6.ocx"
Begin VB.UserControl vbSDhtml 
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LockControls    =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   4800
   Begin VB.Frame fraHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   45
      TabIndex        =   5
      Top             =   4350
      Width           =   3240
      Begin VB.Label lblItemHelp 
         Caption         =   "Label1"
         Height          =   720
         Left            =   105
         TabIndex        =   6
         Top             =   255
         Width           =   1875
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   2625
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
      Left            =   15
      TabIndex        =   2
      Top             =   900
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
   Begin ODCboLst6.OwnerDrawComboList lstEle 
      Height          =   1290
      Left            =   -15
      TabIndex        =   3
      ToolTipText     =   "Double clic to insert in active document"
      Top             =   2970
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   2275
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
      Style           =   4
      MaxLength       =   0
   End
   Begin vbalDkTb6.vbalDockContainer vbalDockContainer1 
      Align           =   1  'Align Top
      Height          =   30
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   53
      AllowUndock     =   0   'False
      LockToolbars    =   -1  'True
   End
End
Attribute VB_Name = "vbSDhtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_cMenu As cPopupMenu
Attribute m_cMenu.VB_VarHelpID = -1
Private m_IniFile As String
Private m_Img As cVBALImageList

Private fLoading As Boolean

Public Event ElementClicked(ByVal Value As String)
Public Event TagClicked(ByVal Value As String)

Private Type arr_elements_det
    Atributo As String
    help As String
End Type

Private Type arr_elements
    Tag As String
    arr_elementos() As arr_elements_det
End Type
Private arr_dhtml() As arr_elements

Private Sub lstEle_Click()

    Dim k As Integer
    Dim j As Integer
    
    If lstObj.Text <> "" Then
        If lstEle.Text <> "" Then
            For k = 1 To UBound(arr_dhtml)
                If arr_dhtml(k).Tag = lstObj.Text Then
                    For j = 1 To UBound(arr_dhtml(k).arr_elementos)
                        If arr_dhtml(k).arr_elementos(j).Atributo = lstEle.Text Then
                            lblItemHelp.Caption = arr_dhtml(k).arr_elementos(j).help
                            Exit For
                        End If
                    Next j
                End If
            Next k
        End If
    End If
End Sub

Private Sub lstEle_DblClick()

    If lstEle.Text <> "" Then
        RaiseEvent ElementClicked(lstEle.Text)
    End If
    
End Sub


Private Sub lstObj_Click()

    'Dim V As Variant
    Dim j As Integer
    Dim k As Integer
    Dim e As Integer
    Dim Tag As String
    Dim elems As String
    Dim sSections() As String
    Dim v2 As Variant
    Dim icono As Integer
    
    If fLoading Then Exit Sub
    
    e = 1
    If lstObj.ListCount > -1 Then
        Tag = lstObj.Text
        For j = 1 To UBound(arr_dhtml)
            If arr_dhtml(j).Tag = Tag Then
                lbl.Caption = Tag
                lstEle.Clear
                fraHelp.Caption = "Help"
                lblItemHelp.Caption = vbNullString
                
                If UBound(arr_dhtml(j).arr_elementos) = 0 Then
                    util.Hourglass hwnd, True
                    'cargar los elementos del objeto
                    v2 = util.LeeIni(m_IniFile, arr_dhtml(j).Tag, "num")
                    If v2 < 81 Then
                        'carga optimizada
                        get_info_section arr_dhtml(j).Tag, sSections, m_IniFile
        
                        For k = 2 To UBound(sSections)
                            If Len(sSections(k)) > 0 Then
                                ReDim Preserve arr_dhtml(j).arr_elementos(e)
                                arr_dhtml(j).arr_elementos(e).Atributo = Explode(sSections(k), 1, "|")
                                arr_dhtml(j).arr_elementos(e).help = Explode(sSections(k), 2, "|")
                                e = e + 1
                            End If
                        Next k
                    Else
                        For k = 2 To v2
                            elems = util.LeeIni(m_IniFile, arr_dhtml(j).Tag, "ele" & k)
                            If Len(elems) > 0 Then
                                ReDim Preserve arr_dhtml(j).arr_elementos(e)
                                arr_dhtml(j).arr_elementos(e).Atributo = Explode(elems, 1, "|")
                                arr_dhtml(j).arr_elementos(e).help = Explode(elems, 2, "|")
                                e = e + 1
                            End If
                        Next k
                    End If
                    util.Hourglass hwnd, False
                End If
                
                icono = lstObj.ListIndex
                
                If icono = 0 Then
                    icono = 0 'colecciones
                ElseIf icono = 1 Then
                    icono = 1   'eventos
                ElseIf icono = 2 Then
                    icono = 2   'metodos
                ElseIf icono = 3 Then
                    icono = 3   'objetos
                ElseIf icono = 4 Then
                    icono = 4   'propiedades
                End If
                        
                For k = 1 To UBound(arr_dhtml(j).arr_elementos)
                    lstEle.AddItemAndData arr_dhtml(j).arr_elementos(k).Atributo, icono
                Next k
                lstEle.ListIndex = 0
                Exit For
            End If
        Next j
    End If

End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    
    LockWindowUpdate hwnd
    
    pic.Move 5, 1, UserControl.Width - 15
    lstObj.Move 0, pic.Height + 1, UserControl.Width, 1500
    lstEle.Move 0, lstObj.Height + pic.Height + 1, UserControl.Width, UserControl.Height - (lstObj.Height) - 2500
    fraHelp.Move 0, lstObj.Height + lstEle.Height + pic.Height + 1, UserControl.Width, UserControl.Height - (lstObj.Height - lstEle.Height - pic.Height) - 100
    lblItemHelp.Move 50, 200, fraHelp.Width - 200, fraHelp.Height - 50
    Err = 0
    LockWindowUpdate False
    
End Sub


Public Property Get inifile() As String
    inifile = m_IniFile
End Property

Public Property Let inifile(ByVal pIniFile As String)
    m_IniFile = pIniFile
End Property


Public Sub Load()

    On Error GoTo ErrorLoad
    
    Dim V As Variant
    Dim j As Integer
    Dim C As Integer
    Dim ele As String
    Dim Tag As String
        
    fLoading = True
    
    V = util.LeeIni(m_IniFile, "dhtml", "num")
    
    BuildImageList
    
    lstObj.ImageList = m_Img.hIml
    lstEle.ImageList = m_Img.hIml
    
    ReDim arr_dhtml(0)
    C = 1
    For j = 1 To V
        ele = LeeIni(m_IniFile, "dhtml", "ele" & j)
        
        If Len(ele) > 0 Then
            ReDim Preserve arr_dhtml(C)
            ReDim Preserve arr_dhtml(C).arr_elementos(0)
                        
            ele = Explode(ele, 1, "|")
            arr_dhtml(C).Tag = ele
            
            If j = 1 Then
                lstObj.AddItemAndData ele, 0 'colecciones
            ElseIf j = 2 Then
                lstObj.AddItemAndData ele, 1   'eventos
            ElseIf j = 3 Then
                lstObj.AddItemAndData ele, 2   'metodos
            ElseIf j = 4 Then
                lstObj.AddItemAndData ele, 3   'objetos
            ElseIf j = 5 Then
                lstObj.AddItemAndData ele, 4   'propiedades
            End If
            C = C + 1
        End If
    Next j
    
    fLoading = False
    
    lstObj.ListIndex = 0
    
    Exit Sub
ErrorLoad:
    debug_startup "error :" & Error$ & " numero :" & Err
    
End Sub


Private Sub BuildImageList()
    
    Set m_Img = New cVBALImageList
    
    With m_Img
        .IconSizeX = 16: .IconSizeY = 16: .ColourDepth = ILC_COLOR24
        .Create
        .AddFromResourceID 253, App.hInstance, IMAGE_ICON, "k1"
        .AddFromResourceID 195, App.hInstance, IMAGE_ICON, "k2"
        .AddFromResourceID 192, App.hInstance, IMAGE_ICON, "k3"
        .AddFromResourceID 263, App.hInstance, IMAGE_ICON, "k4"
        .AddFromResourceID 193, App.hInstance, IMAGE_ICON, "k5"
    End With
   
End Sub
