VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{2128BF45-F895-4206-84CD-F4DE2DD8D6B1}#2.0#0"; "vbsTbar6.ocx"
Object = "{98F993CC-3598-405A-9E9A-0D2CF198B250}#2.0#0"; "vbsDkTb6.ocx"
Begin VB.UserControl ctlProject 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2100
      Top             =   2625
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
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
      ScaleWidth      =   4740
      TabIndex        =   1
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
   Begin MSComctlLib.TreeView tvwProject 
      Height          =   1380
      Left            =   285
      TabIndex        =   0
      Top             =   615
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   2434
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin vbalDkTb6.vbalDockContainer vbalDockContainer1 
      Align           =   1  'Align Top
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   375
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   53
      AllowUndock     =   0   'False
      LockToolbars    =   -1  'True
   End
End
Attribute VB_Name = "ctlProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Img As cVBALImageList
Private WithEvents m_cMenu As cPopupMenu
Attribute m_cMenu.VB_VarHelpID = -1

Public Sub AddFile(ByVal Category As String, ByVal Filename As String)
      
   tvwProject.Nodes.Add Category, tvwChild, Filename, VBArchivoSinPath(Filename)
   tvwProject.Nodes(Filename).Tag = Filename
   
End Sub


Public Function GetCategoryKey(ByVal Category As String) As String

   Dim ret As String
   
   Select Case Category
      Case "JavaScript Files"
         ret = "js"
      Case "HTML Files"
         ret = "html"
      Case "CSS Files"
         ret = "css"
      Case "XML Files"
         ret = "xml"
      Case "Image Files"
         ret = "img"
      Case "Extra Files"
         ret = "ot"
   End Select
   
   GetCategoryKey = ret
   
End Function

Public Property Get JScVBALImageList() As cVBALImageList
    Set JScVBALImageList = m_Img
End Property
Public Property Set JScVBALImageList(ByVal pcVBALImageList As cVBALImageList)
    Set m_Img = pcVBALImageList
End Property
Public Sub Load()

   Dim iMain As Long
    Dim ip As Long
    
    Set m_cMenu = New cPopupMenu
    m_cMenu.hWndOwner = UserControl.hwnd
    m_cMenu.OfficeXpStyle = True
    m_cMenu.ImageList = frmMain.m_MainImg.hIml

    With m_cMenu
        'tools
        iMain = .AddItem("TOOLS", "Tools Toolbar", , , , , , "TOOLSTOOLBAR")
        ip = .AddItem("New", "New Project", , iMain, 194, , , "PROJECT:NEW")
        ip = .AddItem("Open", "Open Project", , iMain, 194, , , "PROJECT:OPEN")
        ip = .AddItem("Save", "Save Project", , iMain, 184, , , "PROJECT:SAVE")
        ip = .AddItem("Publish", "Publish Project", , iMain, 181, , , "PROJECT:PUBLISH")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Add", "Add File", , iMain, 171, , , "PROJECT:ADD")
        ip = .AddItem("AddF", "Add Folder", , iMain, 178, , , "PROJECT:ADDFOLDER")
        ip = .AddItem("Remove ", "Remove File", , iMain, 172, , , "PROJECT:REMOVE")
        ip = .AddItem("-", , , iMain)
    End With

    With tbrTools
        .ImageSource = CTBExternalImageList
        .SetImageList frmMain.m_MainImg, CTBImageListNormal
        .DrawStyle = CTBDrawOfficeXPStyle
        .CreateToolbar 16, True, True, True
        ' Now we create the toolbar:
        .CreateFromMenu2 m_cMenu, CTBToolbarStyle, "TOOLSTOOLBAR"
    End With
    
    With vbalDockContainer1
        .Add "TOOLS", tbrTools.ToolbarWidth, tbrTools.ToolbarHeight, frmMain.getVerticalHeight(tbrTools), frmMain.getVerticalWidth(tbrTools), "Tools"
        .Capture "TOOLS", tbrTools.hwnd
    End With
    
   Call LoadTreeView
   
End Sub


Public Sub LoadTreeView()

   tvwProject.Nodes.Clear
   
   tvwProject.Nodes.Add , , "root", "Project"
   tvwProject.Nodes.Add "root", tvwChild, "js", "JavaScript Files"
   tvwProject.Nodes.Add "root", tvwChild, "html", "HTML Files"
   tvwProject.Nodes.Add "root", tvwChild, "css", "CSS Files"
   tvwProject.Nodes.Add "root", tvwChild, "xml", "XML Files"
   tvwProject.Nodes.Add "root", tvwChild, "img", "Image Files"
   tvwProject.Nodes.Add "root", tvwChild, "ot", "Extra Files"
   tvwProject.Nodes("root").Expanded = True
   
End Sub

Public Sub SetProjectName(ByVal Name As String)
   tvwProject.Nodes("root").Text = "Project - " & Name
   tvwProject.Nodes("root").Tag = Name
End Sub

Private Sub tbrTools_ButtonClick(ByVal lButton As Long)

   Select Case tbrTools.ButtonKey(lButton)
      Case "PROJECT:NEW"
         Call ProjectMan.NewProject
      Case "PROJECT:OPEN"
      
      Case "PROJECT:SAVE"
      
      Case "PROJECT:PUBLISH"
      
      Case "PROJECT:ADD"
      
      Case "PROJECT:ADDFOLDER"
      
      Case "PROJECT:REMOVE"
   
   End Select
   
End Sub


Private Sub UserControl_Resize()

   On Error Resume Next
   
   LockWindowUpdate hwnd
   tvwProject.Move 0, picGeneral.Height + 10, UserControl.Width, UserControl.Height - picGeneral.Height
   LockWindowUpdate False
   
   Err = 0
   
End Sub

