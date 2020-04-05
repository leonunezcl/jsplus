VERSION 5.00
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Begin VB.UserControl QuickHelp 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   ScaleHeight     =   3600
   ScaleWidth      =   4110
   ToolboxBitmap   =   "QuickHelp.ctx":0000
   Begin vbalTreeViewLib6.vbalTreeView tvwJs 
      Height          =   1260
      Left            =   375
      TabIndex        =   0
      Top             =   420
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
End
Attribute VB_Name = "QuickHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_IniFile As String
Private m_Img As cVBALImageList

Private nodOpe As cTreeViewNode
Private nodCon As cTreeViewNode
Private nodLen As cTreeViewNode
Private nodFun As cTreeViewNode
Public Property Let IniFile(ByVal pIniFile As String)
    m_IniFile = pIniFile
End Property


Public Property Get JScVBALImageList() As cVBALImageList
    Set JScVBALImageList = m_Img
End Property



Public Property Set JScVBALImageList(ByVal pcVBALImageList As cVBALImageList)
    Set m_Img = pcVBALImageList
End Property

Public Property Get IniFile() As String
    IniFile = m_IniFile
End Property


Public Sub Load()

    Dim sSections() As String
    Dim sSections2() As String
    
    Dim k As Integer
    Dim j As Integer
    Dim Nodo As cTreeViewNode
    
    tvwJs.ImageList = m_Img.hIml
    Set nodOpe = tvwJs.Nodes.Add(, etvwChild, "kope", "Operators", 123)
    Set nodCon = tvwJs.Nodes.Add(, etvwChild, "kcon", "Control", 120)
    Set nodLen = tvwJs.Nodes.Add(, etvwChild, "klan", "Language", 121)
    Set nodFun = tvwJs.Nodes.Add(, etvwChild, "kfun", "Functions", 122)
    
    'operadores
    get_info_section "operators", sSections, m_IniFile
        
    For k = 2 To UBound(sSections)
        Set Nodo = nodOpe.AddChildNode("kope-" & k, sSections(k), 123)
        get_info_section sSections(k), sSections2, m_IniFile
        For j = 2 To UBound(sSections2)
            Nodo.AddChildNode "kope-" & k & "-" & j, Util.Explode(sSections2(j), 1, "#"), 124
        Next j
        Set Nodo = Nothing
    Next k
    
    'control
    get_info_section "control", sSections, m_IniFile
    For k = 2 To UBound(sSections)
        nodCon.AddChildNode "kcon" & k, sSections(k), 120
    Next k
    
    'lenguaje
    get_info_section "language", sSections, m_IniFile
    For k = 2 To UBound(sSections)
        nodLen.AddChildNode "klan" & k, Util.Explode(sSections(k), 1, "#"), 121
    Next k
    
    'funciones
    get_info_section "functions", sSections, m_IniFile
    For k = 2 To UBound(sSections)
        nodFun.AddChildNode "kfun" & k, Util.Explode(sSections(k), 1, "#"), 122
    Next k
    
    nodOpe.Expanded = True
    nodCon.Expanded = True
    nodLen.Expanded = True
    nodFun.Expanded = True
    
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    
    LockWindowUpdate hWnd
    tvwJs.Move 0, 0, UserControl.Width, UserControl.Height
    LockWindowUpdate False
    Err = 0
    
End Sub


