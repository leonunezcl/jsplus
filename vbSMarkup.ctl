VERSION 5.00
Object = "{2128BF45-F895-4206-84CD-F4DE2DD8D6B1}#2.0#0"; "vbsTbar6.ocx"
Object = "{E7106799-3A07-4335-80BA-4F20E8E5E2E9}#2.0#0"; "vbsODCL6.ocx"
Object = "{98F993CC-3598-405A-9E9A-0D2CF198B250}#2.0#0"; "vbsDkTb6.ocx"
Begin VB.UserControl vbSMarkup 
   Appearance      =   0  'Flat
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7410
   ScaleHeight     =   7605
   ScaleWidth      =   7410
   ToolboxBitmap   =   "vbSMarkup.ctx":0000
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
      Left            =   105
      TabIndex        =   6
      Top             =   4620
      Width           =   3240
      Begin VB.Label lblItemHelp 
         Caption         =   "Label1"
         Height          =   720
         Left            =   105
         TabIndex        =   7
         Top             =   255
         Width           =   1875
      End
   End
   Begin VB.PictureBox picGeneral 
      Align           =   1  'Align Top
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   7350
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   7410
      Begin vbalTBar6.cToolbar tbrTools 
         Height          =   270
         Left            =   660
         Top             =   45
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   476
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
      TabIndex        =   2
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
         TabIndex        =   3
         Top             =   30
         Width           =   855
      End
   End
   Begin ODCboLst6.OwnerDrawComboList lstObj 
      Height          =   1695
      Left            =   15
      TabIndex        =   0
      ToolTipText     =   "Double clic to insert in active document"
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
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Double clic to insert in active document"
      Top             =   2955
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   2275
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
      Style           =   4
      MaxLength       =   0
   End
   Begin vbalDkTb6.vbalDockContainer vbalDockContainer1 
      Align           =   1  'Align Top
      Height          =   30
      Left            =   0
      TabIndex        =   4
      Top             =   375
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   53
      AllowUndock     =   0   'False
      LockToolbars    =   -1  'True
   End
   Begin VB.Menu mnuPop 
      Caption         =   "MenuPop"
      Visible         =   0   'False
      Begin VB.Menu mnuInsert 
         Caption         =   "Insert"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "vbSMarkup"
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

Private Type edesevent
    nombre As String
    descrip As String
End Type

Private arr_standard() As edesevent
Private arr_extended() As edesevent
Private Eventos As String
Private eventfile As String
Public Function get_item_help(ByVal objeto As String, ByVal miembro As String, ByVal tipo As Integer, _
                              ByRef tipomiembro As String, ByRef icono As Integer) As String

    Dim ret As String
    Dim k As Integer
    Dim C As Integer
    Dim tmp As String
    
    For k = 1 To UBound(arr_html)
        If Trim$(LCase$(arr_html(k).Tag)) = objeto Then
            C = k
            Exit For
        End If
    Next k
            
    miembro = LCase$(miembro)
    If C > 0 Then
        For k = 1 To UBound(arr_html(C).elems)
            tmp = Trim$(LCase$(arr_html(C).elems(k).attribute))
            If tmp = miembro Then
                ret = arr_html(C).elems(k).help
                If arr_html(C).elems(k).tipo = 1 Then
                    tipomiembro = "Property"
                    icono = 193
                Else
                    tipomiembro = "Event"
                    icono = 195
                End If
                Exit For
            End If
        Next k
    End If
    
    get_item_help = ret
    
End Function


Private Function ayuda_evento(ByVal Evento As String, ByVal Indice As Integer) As String

    Dim k As Integer
    Dim ret As String
    
    If Indice = 1 Then
        For k = 1 To UBound(arr_standard)
            If arr_standard(k).nombre = Evento Then
                ret = arr_standard(k).descrip
                Exit For
            End If
        Next k
    Else
        For k = 1 To UBound(arr_extended)
            If arr_extended(k).nombre = Evento Then
                ret = arr_extended(k).descrip
                Exit For
            End If
        Next k
    End If
    
    ayuda_evento = ret
    
End Function

Private Sub cargar_eventos()
        
    Dim nFreeFile As Long
    Dim linea As String
    Dim k As Integer
    Dim j As Integer
    
    ReDim arr_standard(0)
    ReDim arr_extended(0)
    
    eventfile = util.StripPath(App.Path) & "config\events.ini"
    
    nFreeFile = FreeFile
    k = 1
    j = 1
    Open eventfile For Input As #nFreeFile
        Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, linea
            linea = Trim$(util.SacarBasura(linea))
            If Left$(linea, 1) <> "[" Then
            If Len(linea) > 0 Then
                If j = 1 Then
                    ReDim Preserve arr_standard(k)
                    arr_standard(k).nombre = util.Explode(linea, 1, "=")
                    arr_standard(k).descrip = util.Explode(linea, 2, "=")
                    k = k + 1
                Else
                    ReDim Preserve arr_extended(k)
                    arr_extended(k).nombre = util.Explode(linea, 1, "=")
                    arr_extended(k).descrip = util.Explode(linea, 2, "=")
                    k = k + 1
                End If
            Else
                j = j + 1
                If j > 2 Then
                    Exit Do
                End If
                k = 1
            End If
            End If
        Loop
    Close #nFreeFile
        
End Sub

Public Sub Prepare()
    
    Dim iMain As Long
    Dim ip As Long
    
    BuildImageList
    
    Set m_cMenu = New cPopupMenu
    m_cMenu.hWndOwner = UserControl.hwnd
    m_cMenu.OfficeXpStyle = True
    m_cMenu.ImageList = m_Img.hIml

    With m_cMenu
        'tools
        iMain = .AddItem("TOOLS", "Tools Toolbar", , , , , , "TOOLSTOOLBAR")
        ip = .AddItem("Insert", "Inserts tag", , iMain, 0, , , "TOOLS:INS")
        ip = .AddItem("Help", "Item Help", , iMain, 1, , , "TOOLS:HELP")
    End With

    With tbrTools
        .ImageSource = CTBExternalImageList
        .SetImageList m_Img, CTBImageListNormal
        .DrawStyle = CTBDrawOfficeXPStyle
        .CreateToolbar 16, True, True, True
        .CreateFromMenu2 m_cMenu, CTBToolbarStyle, "TOOLSTOOLBAR"
    End With
    
    With vbalDockContainer1
        .Add "TOOLS", tbrTools.ToolbarWidth, tbrTools.ToolbarHeight, frmMain.getVerticalHeight(tbrTools), frmMain.getVerticalWidth(tbrTools), "Tools"
        .Capture "TOOLS", tbrTools.hwnd
    End With
        
End Sub


Public Sub Load()

    On Error GoTo ErrorLoad
    
    Dim V As Variant
    Dim j As Integer
    Dim k As Integer
    Dim C As Integer
    Dim e As Integer
    Dim ele As String
    Dim elems As String
    Dim Tag As String
    Dim help As String
    Dim Archivo As String
    Dim sSections() As String
    Dim ev As Integer
    Dim streven As String
    
    fLoading = True
    
    Call cargar_eventos
    
    V = util.LeeIni(m_IniFile, "html_help", "num")
    
    Archivo = util.StripPath(App.Path) & "config\htmlitemhelp.ini"
    
    lstObj.ImageList = m_Img.hIml
    lstEle.ImageList = m_Img.hIml
    e = 1
    
    ReDim arr_html(0)
    C = 1
    For j = 3 To V 'UBound(sSections)
        ele = LeeIni(m_IniFile, "html_help", "ele" & j)
        
        If Len(ele) > 0 Then
            help = util.Explode(ele, 5, "|")
            
            elems = util.Explode(ele, 7, "|")
            
            Tag = util.Explode(ele, 2, "|")
            
            ele = Tag
            
            'eliminar espacios en blanco
            If ele <> "<!-- -->" Then
                If InStr(ele, " ") Then
                    ele = VBA.Left$(ele, InStr(ele, " ") - 1)
                End If
                ele = Replace(ele, "<", "")
                ele = Replace(ele, ">", "")
            End If
            
            lstObj.AddItemAndData ele, 2
            
            ReDim Preserve arr_html(C)
            arr_html(C).Tag = ele
            arr_html(C).HTML = Replace(Tag, " ", "")
            arr_html(C).help = help
            
            'leer los atributos de este
            ReDim arr_html(C).elems(0)
            
            get_info_section ele, sSections, Archivo
                        
            e = 1
            For k = 3 To UBound(sSections)
                ReDim Preserve arr_html(C).elems(e)
                arr_html(C).elems(e).attribute = Explode(sSections(k), 1, "#")
                arr_html(C).elems(e).tipo = Explode(sSections(k), 2, "#")
                arr_html(C).elems(e).help = Explode(sSections(k), 3, "#")
                e = e + 1
            Next k

            'cargar los eventos standard del tag
            Eventos = Trim$(util.SacarBasura(util.LeeIni(eventfile, ele, "standard")))
            
            If Len(Eventos) > 0 Then
                For ev = 1 To 30
                    If Len(util.Explode(Eventos, ev, ",")) > 0 Then
                        streven = Trim$(util.Explode(Eventos, ev, ","))
                        ReDim Preserve arr_html(C).elems(e)
                        arr_html(C).elems(e).attribute = streven
                        arr_html(C).elems(e).tipo = 3
                        arr_html(C).elems(e).help = ayuda_evento(streven, 1)
                        arr_html(C).elems(e).icono = 1
                        e = e + 1
                    Else
                        Exit For
                    End If
                Next ev
            End If
                                        
            'cargar los eventos extendidos del tag
            Eventos = Trim$(util.SacarBasura(util.LeeIni(eventfile, ele, "iexplorer")))
            
            If Len(Eventos) > 0 Then
                For ev = 1 To 30
                    If Len(util.Explode(Eventos, ev, ",")) > 0 Then
                        streven = Trim$(util.Explode(Eventos, ev, ","))
                        ReDim Preserve arr_html(C).elems(e)
                        arr_html(C).elems(e).attribute = streven
                        arr_html(C).elems(e).tipo = 3
                        arr_html(C).elems(e).help = ayuda_evento(streven, 2)
                        arr_html(C).elems(e).icono = 2
                        e = e + 1
                    Else
                        Exit For
                    End If
                Next ev
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
        .AddFromResourceID 244, App.hInstance, IMAGE_ICON, "k1"
        .AddFromResourceID 241, App.hInstance, IMAGE_ICON, "k2"
        .AddFromResourceID 200, App.hInstance, IMAGE_ICON, "k3"
        .AddFromResourceID 193, App.hInstance, IMAGE_ICON, "k4"
        .AddFromResourceID 195, App.hInstance, IMAGE_ICON, "k5"
        .AddFromResourceID 254, App.hInstance, IMAGE_ICON, "k6"
    End With
   
End Sub
Private Sub lstEle_Click()

    Dim k As Integer
    Dim j As Integer
    
    If lstObj.Text <> "" Then
        If lstEle.Text <> "" Then
            For k = 1 To UBound(arr_html)
                If arr_html(k).Tag = lstObj.Text Then
                    For j = 1 To UBound(arr_html(k).elems)
                        If arr_html(k).elems(j).attribute = lstEle.Text Then
                            lblItemHelp.Caption = arr_html(k).elems(j).help
                            Exit For
                        End If
                    Next j
                End If
            Next k
        End If
    End If
    
End Sub

Private Sub lstEle_DblClick()

    If Len(lstEle.Text) > 0 Then
        RaiseEvent ElementClicked(lstEle.Text & "=" & Chr$(34) & Chr$(34))
    End If
    
End Sub


Private Sub lstObj_Click()

    Dim j As Integer
    Dim k As Integer
    Dim Tag As String
    Dim elem As String
    Dim tipo As String
    
    If fLoading Then Exit Sub
    
    If lstObj.ListCount > -1 Then
        Tag = lstObj.Text
        lblItemHelp.Caption = vbNullString
        For j = 1 To UBound(arr_html)
            If arr_html(j).Tag = Tag Then
                If lstObj.ListIndex > 1 Then
                    If InStr(1, Tag, "<") Then
                        Tag = Mid$(Tag, 2)
                        If InStr(1, Tag, ">") Then
                            lbl.Caption = VBA.Left$(Tag, InStr(2, Tag, ">") - 1)
                        Else
                            lbl.Caption = Tag
                        End If
                    Else
                        lbl.Caption = Tag
                    End If
                Else
                    lbl.Caption = Tag
                End If
                lstEle.Clear
                                                
                For k = 3 To UBound(arr_html(j).elems)
                    elem = arr_html(j).elems(k).attribute
                    tipo = arr_html(j).elems(k).tipo
                    If tipo = "1" Then
                        lstEle.AddItemAndData elem, 3
                    Else
                        If arr_html(j).elems(k).icono = 1 Then
                            lstEle.AddItemAndData elem, 4
                        Else
                            lstEle.AddItemAndData elem, 5
                        End If
                    End If
                Next k
                lstEle.ListIndex = 0
                Exit For
            End If
        Next j
    End If
    
End Sub

Private Sub lstObj_DblClick()

    If Len(lstObj.Text) > 0 Then
        RaiseEvent TagClicked(arr_html(lstObj.ListIndex + 1).HTML)
    End If
    
End Sub

Private Sub lstObj_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuPop
    End If
    
End Sub

Private Sub mnuHelp_Click()

    If Len(lstObj.Text) > 0 Then
        frmHtmlHelp.elem = lstObj.ListIndex + 3
        
        frmHtmlHelp.Show vbModal
    End If
    
End Sub

Private Sub mnuInsert_Click()

    If Len(lstObj.Text) > 0 Then
        RaiseEvent TagClicked(arr_html(lstObj.ListIndex + 1).HTML)
    End If
    
End Sub

Private Sub tbrTools_ButtonClick(ByVal lButton As Long)

    Select Case tbrTools.ButtonKey(lButton)
        Case "TOOLS:INS"
            Call mnuInsert_Click
        Case "TOOLS:HELP"
            Call mnuHelp_Click
    End Select
    
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    
    LockWindowUpdate hwnd
    
    pic.Move 5, picGeneral.Height, UserControl.Width - 15
    lstObj.Move 0, pic.Height + picGeneral.Height + 1, UserControl.Width, 3000
    lstEle.Move 0, pic.Height + picGeneral.Height + lstObj.Height, UserControl.Width, UserControl.Height - (lstObj.Height + pic.Height + picGeneral.Height) - 2500
    fraHelp.Move 0, pic.Height + picGeneral.Height + lstObj.Height + lstEle.Height, UserControl.Width, UserControl.Height - (lstObj.Height + pic.Height + picGeneral.Height - lstEle.Height) - 100
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

