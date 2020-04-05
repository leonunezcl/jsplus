VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{2128BF45-F895-4206-84CD-F4DE2DD8D6B1}#2.0#0"; "vbsTbar6.ocx"
Object = "{98F993CC-3598-405A-9E9A-0D2CF198B250}#2.0#0"; "vbsDkTb6.ocx"
Begin VB.UserControl HTMLTree 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4140
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "HTMLTree.ctx":0000
            Key             =   ""
         EndProperty
      EndProperty
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
         Left            =   120
         Top             =   45
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   476
      End
   End
   Begin MSComctlLib.TreeView tview 
      Height          =   2295
      Left            =   300
      TabIndex        =   0
      Top             =   1170
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4048
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   178
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
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
   End
End
Attribute VB_Name = "HTMLTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents MyTree As clsTree
Attribute MyTree.VB_VarHelpID = -1
Private m_Img As cVBALImageList
Private WithEvents m_cMenu As cPopupMenu
Attribute m_cMenu.VB_VarHelpID = -1
Public Property Get JScVBALImageList() As cVBALImageList
    Set JScVBALImageList = m_Img
End Property
Public Property Set JScVBALImageList(ByVal pcVBALImageList As cVBALImageList)
    Set m_Img = pcVBALImageList
End Property
Public Sub Prepare()
    
    Dim iMain As Long
    Dim ip As Long
    
    Set m_cMenu = New cPopupMenu
    m_cMenu.hWndOwner = UserControl.hwnd
    m_cMenu.OfficeXpStyle = True
    m_cMenu.ImageList = frmMain.m_MainImg.hIml

    With m_cMenu
        'tools
        iMain = .AddItem("TOOLS", "Tools Toolbar", , , , , , "TOOLSTOOLBAR")
        ip = .AddItem("Expand", "Expand Tree", , iMain, 171, , , "TREE:EXPAND")
        ip = .AddItem("Collapse", "Collapse Tree", , iMain, 172, , , "TREE:COLLAPSE")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Configure", "Configure", , iMain, 176, , , "TREE:CONFIGURE")
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
       
End Sub

Public Sub Load(ByVal Archivo As String)

   Dim lCnt As Long
   Dim ext As String
   Dim IniFileTree As String
   Dim arr_ext() As String
   Dim k As Integer
   Dim ret As New cStringBuilder
   
   Exit Sub
   
   If InStr(Archivo, ".") > 0 Then
      ext = Mid$(Archivo, InStr(Archivo, ".") + 1)
      
      IniFileTree = StripPath(App.Path) & "treeext.ini"
   
      get_info_section "tree", arr_ext, IniFileTree
   
      For k = 1 To UBound(arr_ext)
         If arr_ext(k) = ext Then
         
            Set ret = get_tags(Archivo)
            
            Dim nFreeFile As Long
            
            nFreeFile = FreeFile
            
            Open StripPath(App.Path) & "htmltemp.txt" For Output As #nFreeFile
               Print #nFreeFile, ret.ToString
            Close #nFreeFile
            
            MyTree.InitializeHTML
              
            If Len(ret.ToString) > 0 Then
               LockWindowUpdate tview.hwnd
            
               MyTree.ProduceTree tview
            
               For lCnt = 1 To tview.Nodes.count
                  tview.Nodes(lCnt).Expanded = True
               Next lCnt
            
               tview.SelectedItem = tview.Nodes(1)
               
               LockWindowUpdate False
            End If
            
            Set ret = Nothing
            
            HTML = ""
            Exit For
         End If
      Next k
   End If
    
End Sub

Private Sub tbrTools_ButtonClick(ByVal lButton As Long)

   Dim k As Integer
   
   util.Hourglass hwnd, True
   LockWindowUpdate tview.hwnd
   
   Select Case tbrTools.ButtonKey(lButton)
        Case "TREE:COLLAPSE"
            For k = 1 To tview.Nodes.count
               tview.Nodes(k).Expanded = False
            Next k
        Case "TREE:EXPAND"
            For k = 1 To tview.Nodes.count
               tview.Nodes(k).Expanded = True
            Next k
        Case "TREE:CONFIGURE"
            frmSetTreeExp.Show vbModal
    End Select
    
    LockWindowUpdate False
    util.Hourglass hwnd, False
    
End Sub

Private Sub tview_NodeClick(ByVal Node As MSComctlLib.Node)

'   Dim r As CodeSenseCtl.range
'
'   Set r = frmMain.ActiveForm.txtCode.GetSel(True)
'
'   r.StartColNo = MyTree.ReturnPos(val(Node.key)) - 1
'   r.StartLineNo = MyTree.ReturnLine(val(Node.key)) - 1
'   r.EndLineNo = MyTree.ReturnLine(val(Node.key)) - 1
'
'   Call frmMain.ActiveForm.txtCode.SetSel(r, True)
'   frmMain.ActiveForm.txtCode.SetFocus
  
End Sub


Private Sub UserControl_Initialize()
   Set MyTree = New clsTree
End Sub

Private Sub UserControl_Resize()

   On Error Resume Next
   
   LockWindowUpdate hwnd
   tview.Move 0, picGeneral.Height + 1, UserControl.Width, UserControl.Height - picGeneral.Height - 245
   LockWindowUpdate False
   
   Err = 0
   
End Sub


Private Sub UserControl_Terminate()
   Set MyTree = Nothing
End Sub


