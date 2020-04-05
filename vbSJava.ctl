VERSION 5.00
Object = "{2128BF45-F895-4206-84CD-F4DE2DD8D6B1}#2.0#0"; "vbsTbar6.ocx"
Object = "{E7106799-3A07-4335-80BA-4F20E8E5E2E9}#2.0#0"; "vbsODCL6.ocx"
Object = "{8C44B082-B582-4258-9E2C-7D9383CE7DF4}#1.0#0"; "vbsTreeView6.ocx"
Object = "{98F993CC-3598-405A-9E9A-0D2CF198B250}#2.0#0"; "vbsDkTb6.ocx"
Begin VB.UserControl vbSJava 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vbSJava.ctx":0000
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
      ScaleWidth      =   4740
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   4800
      Begin vbalTBar6.cToolbar tbrTools 
         Height          =   270
         Left            =   660
         Top             =   45
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   476
      End
   End
   Begin vbalTreeViewLib6.vbalTreeView tvwJs 
      Height          =   1260
      Left            =   585
      TabIndex        =   3
      Top             =   3090
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2223
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
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   585
      ScaleHeight     =   255
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   2775
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
      Left            =   570
      TabIndex        =   2
      Top             =   1065
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   2990
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
   Begin vbalDkTb6.vbalDockContainer vbalDockContainer1 
      Align           =   1  'Align Top
      Height          =   30
      Left            =   0
      TabIndex        =   5
      Top             =   375
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   53
      AllowUndock     =   0   'False
      LockToolbars    =   -1  'True
   End
End
Attribute VB_Name = "vbSJava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Img As cVBALImageList
Private WithEvents m_cMenu As cPopupMenu
Attribute m_cMenu.VB_VarHelpID = -1
Private m_IniFile As String

Private Type ejshelp
    miembro As String
    tipo As Integer
    help As String
End Type

Private Type eJs
    Tag As String
    prop() As ejshelp
    meth() As ejshelp
    eve() As ejshelp
    cons() As ejshelp
    coll() As ejshelp
    obj() As ejshelp
End Type

Private arr_js() As eJs

Private nodProp As cTreeViewNode
Private nodMeth As cTreeViewNode
Private nodEven As cTreeViewNode
Private nodCons As cTreeViewNode
Private nodColl As cTreeViewNode
Private nodObj As cTreeViewNode

Private m_Loaded As Boolean
Private fLoading As Boolean

Public Event ElementClicked(ByVal Value As String)
Private Sub BuildImageList()
    
    Set m_Img = New cVBALImageList
    
    With m_Img
        .IconSizeX = 16: .IconSizeY = 16: .ColourDepth = ILC_COLOR24
        .Create
        .AddFromResourceID 244, App.hInstance, IMAGE_ICON, "k1"
        .AddFromResourceID 241, App.hInstance, IMAGE_ICON, "k2"
        .AddFromResourceID 199, App.hInstance, IMAGE_ICON, "k3"
        .AddFromResourceID 193, App.hInstance, IMAGE_ICON, "k4"
        .AddFromResourceID 192, App.hInstance, IMAGE_ICON, "k5"
        .AddFromResourceID 195, App.hInstance, IMAGE_ICON, "k6"
        .AddFromResourceID 194, App.hInstance, IMAGE_ICON, "k7"
        .AddFromResourceID 253, App.hInstance, IMAGE_ICON, "k8"
        .AddFromResourceID 263, App.hInstance, IMAGE_ICON, "k9"
    End With
   
End Sub

Private Sub ayudajs()

    Dim j As Integer
    Dim k As Integer
    Dim ini As String
    Dim tipo As String
    Dim miembro As String
    Dim w As String
    Dim glosa As String
    Dim objeto As String
    Dim sSections() As String
    
    If Not tvwJs.SelectedItem Is Nothing Then
        w = tvwJs.SelectedItem.Text
        
        If InStr(w, "(") Then
            w = Left$(w, InStr(w, "(") - 1)
        End If
        
        If InStr(w, "Properties") Or InStr(w, "Events") Or InStr(w, "Methods") Then
            Exit Sub
        End If
        
        ini = util.StripPath(App.Path) & "config\jshelp.ini"
            
        If Not ArchivoExiste2(ini) Then
            MsgBox "File not found : " & ini, vbCritical
            Exit Sub
        End If
        
        For j = 1 To UBound(udtObjetos)
            objeto = udtObjetos(j)
            get_info_section objeto, sSections, ini
            
            For k = 2 To UBound(sSections)
                glosa = sSections(k)
                If Len(glosa) > 0 Then
                    miembro = util.Explode(glosa, 1, "#")
                    tipo = util.Explode(glosa, 2, "#")
                    glosa = util.Explode(glosa, 3, "#")
                    
                    If Len(miembro) > 0 Then
                        If InStr(miembro, "(") Then
                            miembro = Left$(miembro, InStr(miembro, "(") - 1)
                        End If
                        If LCase$(miembro) = LCase$(w) Then
                            util.Hourglass hwnd, False
                            frmItemHelp.member = miembro
                            frmItemHelp.mtype = tipo
                            frmItemHelp.mhelp = glosa
                            frmItemHelp.Show vbModal
                            Exit Sub
                        End If
                    End If
                End If
            Next k
        Next j
    End If

End Sub

Public Function get_item_help(ByVal objeto As String, ByVal miembro As String, ByVal tipo As Integer, _
                              ByRef tipomiembro As String, ByRef icono As Integer) As String

    Dim ret As String
    Dim k As Integer
    Dim C As Integer
    Dim tmp As String
    
    objeto = Trim$(LCase$(objeto))
    For k = 1 To UBound(arr_js)
        If Trim$(LCase$(arr_js(k).Tag)) = objeto Then
            C = k
            Exit For
        End If
    Next k
            
    miembro = LCase$(miembro)
    If C > 0 Then
        If tipo = 1 Then    'propiedad
            For k = 1 To UBound(arr_js(C).prop)
                tmp = Trim$(LCase$(arr_js(C).prop(k).miembro))
                If tmp = miembro Then
                    ret = arr_js(C).prop(k).help
                    tipomiembro = "Property"
                    icono = 193
                    Exit For
                End If
            Next k
        ElseIf tipo = 2 Then 'metodo
            For k = 1 To UBound(arr_js(C).meth)
                tmp = Trim$(LCase$(arr_js(C).meth(k).miembro))
                If InStr(tmp, "(") Then
                    tmp = Left$(tmp, InStr(1, tmp, "(") - 1)
                End If
                If tmp = miembro Then
                    ret = arr_js(C).meth(k).help
                    tipomiembro = "Method"
                    icono = 191
                    Exit For
                End If
            Next k
        ElseIf tipo = 3 Then 'evento
            For k = 1 To UBound(arr_js(C).eve)
                tmp = Trim$(LCase$(arr_js(C).eve(k).miembro))
                If InStr(tmp, "(") Then
                    tmp = Left$(tmp, InStr(1, tmp, "(") - 1)
                End If
                If tmp = miembro Then
                    tipomiembro = "Event"
                    ret = arr_js(C).eve(k).help
                    icono = 195
                    Exit For
                End If
            Next k
        ElseIf tipo = 4 Then 'constante
            For k = 1 To UBound(arr_js(C).cons)
                tmp = Trim$(LCase$(arr_js(C).cons(k).miembro))
                If tmp = miembro Then
                    tipomiembro = "Constant"
                    ret = arr_js(C).cons(k).help
                    icono = 194
                    Exit For
                End If
            Next k
        ElseIf tipo = 5 Then 'coleccion
            For k = 1 To UBound(arr_js(C).coll)
                tmp = Trim$(LCase$(arr_js(C).coll(k).miembro))
                If tmp = miembro Then
                    tipomiembro = "Collection"
                    ret = arr_js(C).coll(k).help
                    icono = 253
                    Exit For
                End If
            Next k
        ElseIf tipo = 6 Then 'objeto
            For k = 1 To UBound(arr_js(C).obj)
                tmp = Trim$(LCase$(arr_js(C).obj(k).miembro))
                If tmp = miembro Then
                    tipomiembro = "Object"
                    ret = arr_js(C).obj(k).help
                    icono = 263
                    Exit For
                End If
            Next k
        End If
    End If
    
    get_item_help = ret
    
End Function

Public Sub Prepare()
    
    Dim iMain As Long
    Dim ip As Long
        
    BuildImageList
    
    Set m_cMenu = New cPopupMenu
    m_cMenu.hWndOwner = UserControl.hwnd
    m_cMenu.OfficeXpStyle = True
    m_cMenu.ImageList = m_Img.hIml

    With m_cMenu
        iMain = .AddItem("TOOLS", "Tools Toolbar", , , , , , "TOOLSTOOLBAR")
        ip = .AddItem("Insert", "Inserts tag", , iMain, 0, , , "TOOLS:INS")
        ip = .AddItem("Help", "Item Help", , iMain, 1, , , "TOOLS:HELP")
    End With

    With tbrTools
        .ImageSource = CTBExternalImageList
        .SetImageList m_Img.hIml, CTBImageListNormal
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
    
    Dim j As Integer
    Dim k As Integer
    
    Dim ce As Integer
    Dim cm As Integer
    Dim cp As Integer
    Dim co As Integer
    Dim ccol As Integer
    Dim ccon As Integer
    
    Dim C As Long
    
    Dim ele As String
    Dim info As String
    Dim tipo As String
    
    fLoading = True
        
    Dim sSections() As String
    Dim Atributos() As String
    
    ReDim arr_js(0)
    
    get_info_section "objetos", sSections, m_IniFile
    
    C = 1
    
    lstObj.ImageList = m_Img.hIml
    For j = 2 To UBound(sSections)
        ele = sSections(j)
        If Len(ele) > 0 Then
            lstObj.AddItemAndData ele, 2
            ReDim Preserve arr_js(C)
            ReDim arr_js(C).eve(0)
            ReDim arr_js(C).meth(0)
            ReDim arr_js(C).prop(0)
            ReDim arr_js(C).cons(0)
            ReDim arr_js(C).coll(0)
            ReDim arr_js(C).obj(0)
                        
            arr_js(C).Tag = ele
            
            'leer los eventos,methodos y propiedades del objeto
            get_info_section ele, Atributos, m_IniFile
    
            If UBound(Atributos) > 0 Then
                ce = 1: cm = 1: cp = 1: ccon = 1: ccol = 1: co = 1
                For k = 2 To UBound(Atributos)
                    info = Atributos(k)
                    
                    If Len(info) >= 0 Then
                        tipo = util.Explode(info, 2, "#")
                        
                        If tipo = 1 Then    'propiedad
                            ReDim Preserve arr_js(C).prop(cp)
                            arr_js(C).prop(cp).miembro = util.Explode(info, 1, "#")
                            arr_js(C).prop(cp).help = util.Explode(info, 3, "#")
                            arr_js(C).prop(cp).tipo = 1
                            cp = cp + 1
                        ElseIf tipo = 2 Then 'metodo
                            ReDim Preserve arr_js(C).meth(cm)
                            arr_js(C).meth(cm).miembro = util.Explode(info, 1, "#")
                            arr_js(C).meth(cm).help = util.Explode(info, 3, "#")
                            arr_js(C).meth(cm).tipo = 2
                            cm = cm + 1
                        ElseIf tipo = 3 Then 'evento
                            ReDim Preserve arr_js(C).eve(ce)
                            arr_js(C).eve(ce).miembro = util.Explode(info, 1, "#")
                            arr_js(C).eve(ce).help = util.Explode(info, 3, "#")
                            arr_js(C).eve(ce).tipo = 3
                            ce = ce + 1
                        ElseIf tipo = 4 Then 'constante
                            ReDim Preserve arr_js(C).cons(ccon)
                            arr_js(C).cons(ccon).miembro = util.Explode(info, 1, "#")
                            arr_js(C).cons(ccon).help = util.Explode(info, 3, "#")
                            arr_js(C).cons(ccon).tipo = 4
                            ccon = ccon + 1
                        ElseIf tipo = 5 Then 'coleccion
                            ReDim Preserve arr_js(C).coll(ccol)
                            arr_js(C).coll(ccol).miembro = util.Explode(info, 1, "#")
                            arr_js(C).coll(ccol).help = util.Explode(info, 3, "#")
                            arr_js(C).coll(ccol).tipo = 5
                            ccol = ccol + 1
                        ElseIf tipo = 6 Then 'objeto
                            ReDim Preserve arr_js(C).obj(co)
                            arr_js(C).obj(co).miembro = util.Explode(info, 1, "#")
                            arr_js(C).obj(co).help = util.Explode(info, 3, "#")
                            arr_js(C).obj(co).tipo = 6
                            co = co + 1
                        End If
                    End If
                Next k
            End If
            C = C + 1
        End If
    Next j
            
    tvwJs.ImageList = m_Img.hIml
    
    Set nodProp = tvwJs.Nodes.Add(, etvwChild, "kprop", "Properties", 3)
    
    Set nodMeth = tvwJs.Nodes.Add(, etvwChild, "kmeth", "Methods", 4)
    
    Set nodEven = tvwJs.Nodes.Add(, etvwChild, "keven", "Events", 5)
        
    Set nodCons = tvwJs.Nodes.Add(, etvwChild, "kcons", "Constants", 6)
    
    Set nodColl = tvwJs.Nodes.Add(, etvwChild, "kcoll", "Collection", 7)
    
    Set nodObj = tvwJs.Nodes.Add(, etvwChild, "kobj", "Objects", 8)
    
    fLoading = False
    
    lstObj.ListIndex = 0
    
    m_Loaded = True
    
    Exit Sub
    
ErrorLoad:
    'debug_startup "Load -> error : " & Error$ & " numero :" & Err
    
End Sub

Private Sub lstObj_Click()

    Dim j As Integer
    Dim k As Integer
    Dim Tag As String
    
    If fLoading Then Exit Sub
    
    If lstObj.ListCount > -1 Then
        Tag = lstObj.Text
        For j = 1 To UBound(arr_js)
            If arr_js(j).Tag = Tag Then
                lbl.Caption = Tag
                
                nodProp.Delete
                nodMeth.Delete
                nodEven.Delete
                nodCons.Delete
                nodColl.Delete
                nodObj.Delete
                
                Set nodProp = tvwJs.Nodes.Add(, etvwChild, "kprop", "Properties", 3)
                nodProp.Bold = True
                
                Set nodMeth = tvwJs.Nodes.Add(, etvwChild, "kmeth", "Methods", 4)
                nodMeth.Bold = True
                
                Set nodEven = tvwJs.Nodes.Add(, etvwChild, "keven", "Events", 5)
                nodEven.Bold = True
                
                Set nodCons = tvwJs.Nodes.Add(, etvwChild, "kcons", "Constants", 6)
                nodCons.Bold = True
                
                Set nodColl = tvwJs.Nodes.Add(, etvwChild, "kcoll", "Collection", 7)
                nodColl.Bold = True
                
                Set nodObj = tvwJs.Nodes.Add(, etvwChild, "kobj", "Objects", 8)
                nodObj.Bold = True
                
                For k = 1 To UBound(arr_js(j).prop)
                    nodProp.AddChildNode "kprop" & k, arr_js(j).prop(k).miembro, 3
                Next k
                
                For k = 1 To UBound(arr_js(j).meth)
                    nodMeth.AddChildNode "kmeth" & k, arr_js(j).meth(k).miembro, 4
                Next k
                
                For k = 1 To UBound(arr_js(j).eve)
                    nodEven.AddChildNode "keven" & k, arr_js(j).eve(k).miembro, 5
                Next k
                
                For k = 1 To UBound(arr_js(j).cons)
                    nodCons.AddChildNode "kcons" & k, arr_js(j).cons(k).miembro, 6
                Next k
                
                For k = 1 To UBound(arr_js(j).coll)
                    nodColl.AddChildNode "kcoll" & k, arr_js(j).coll(k).miembro, 7
                Next k
                
                For k = 1 To UBound(arr_js(j).obj)
                    nodObj.AddChildNode "kobj" & k, arr_js(j).obj(k).miembro, 8
                Next k
                
                nodProp.Expanded = True
                nodMeth.Expanded = True
                nodEven.Expanded = True
                nodCons.Expanded = True
                nodColl.Expanded = True
                nodObj.Expanded = True
                
                nodProp.Selected = True
                Exit For
            End If
        Next j
    End If

End Sub
Private Sub tbrTools_ButtonClick(ByVal lButton As Long)

    Select Case tbrTools.ButtonKey(lButton)
        Case "TOOLS:INS"
            If Not tvwJs.SelectedItem Is Nothing Then
                tvwJs_NodeDblClick tvwJs.SelectedItem
            End If
        Case "TOOLS:HELP"
            Call ayudajs
    End Select
    
End Sub

Private Sub tvwJs_NodeDblClick(Node As vbalTreeViewLib6.cTreeViewNode)
    
    If Not Node Is Nothing Then
        If InStr(Node.Text, "Properties") Or InStr(Node.Text, "Events") Or InStr(Node.Text, "Methods") Then
            Exit Sub
        ElseIf InStr(Node.Text, "Constants") Or InStr(Node.Text, "Collection") Or InStr(Node.Text, "Objects") Then
            Exit Sub
        End If
        
        RaiseEvent ElementClicked(lstObj.Text & "." & Node.Text)
    End If
    
End Sub


Private Sub UserControl_Resize()

    On Error Resume Next
    
    LockWindowUpdate hwnd
    pic.Move 5, picGeneral.Height, UserControl.Width - 15
    lstObj.Move 0, pic.Height + picGeneral.Height + 1, UserControl.Width, 3500
    tvwJs.Move 0, pic.Height + picGeneral.Height + lstObj.Height, UserControl.Width, UserControl.Height - (lstObj.Height + pic.Height + picGeneral.Height)
    Err = 0
    
    LockWindowUpdate False
    Err = 0
    
End Sub



Public Property Get inifile() As String
    inifile = m_IniFile
End Property

Public Property Let inifile(ByVal pIniFile As String)
    m_IniFile = pIniFile
End Property



Public Property Get Loaded() As Boolean
    Loaded = m_Loaded
End Property

Public Property Let Loaded(ByVal pLoaded As Boolean)
    Loaded = pLoaded
End Property

