VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmViewTree 
   Caption         =   "XML Tree Viewer"
   ClientHeight    =   8910
   ClientLeft      =   2955
   ClientTop       =   1875
   ClientWidth     =   11400
   ControlBox      =   0   'False
   Icon            =   "ViewTree.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   11400
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   14295
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   8670
      Visible         =   0   'False
      Width           =   2250
   End
   Begin CodeSenseCtl.CodeSense txtCode 
      Height          =   1095
      Left            =   17340
      OleObjectBlob   =   "ViewTree.frx":0442
      TabIndex        =   3
      Top             =   2640
      Width           =   1620
   End
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   2250
      Left            =   12615
      TabIndex        =   2
      Top             =   2580
      Width           =   4500
      ExtentX         =   7937
      ExtentY         =   3969
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.TreeView TreeView 
      Height          =   7620
      Left            =   90
      TabIndex        =   0
      Top             =   645
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   13441
      _Version        =   393217
      Indentation     =   441
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11430
      Top             =   3015
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewTree.frx":05AE
            Key             =   "FClosed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewTree.frx":0A00
            Key             =   "FOpen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewTree.frx":31B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewTree.frx":3604
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ViewTree.frx":391E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   8070
      Left            =   30
      TabIndex        =   1
      Top             =   270
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   14235
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "TreeView"
            Object.ToolTipText     =   "Graphical View"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            Object.ToolTipText     =   "Preview in Browser"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Source"
            Object.ToolTipText     =   "XML Source Code"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin jsplus.MyButton cmd 
      Height          =   390
      Index           =   0
      Left            =   5355
      TabIndex        =   5
      Top             =   8475
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   688
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Exit"
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
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   30
      TabIndex        =   6
      Top             =   15
      Width           =   480
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "File"
      Begin VB.Menu mnuLoadXML 
         Caption         =   "Open File ..."
      End
      Begin VB.Menu mnuSaveXML 
         Caption         =   "Save File ..."
      End
      Begin VB.Menu mnuCloseXML 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmViewTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Filename As String
Private currTreeNode As MSComctlLib.node
Private objXMLTree As XMLTree

Private Sub cmd_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Activate()
    TreeView.ZOrder 0
End Sub

Private Sub Form_Load()

    util.Hourglass hWnd, True
    util.CenterForm Me
    set_color_form Me
    
    web1.Move TreeView.Left, TreeView.Top, TreeView.Width, TreeView.Height
    txtCode.Move TreeView.Left, TreeView.Top, TreeView.Width, TreeView.Height
    
    Set objXMLTree = New XMLTree
    objXMLTree.SetTreeView TreeView
    
    Set TreeView.ImageList = ImageList1
    
    objXMLTree.OpenXMLFile Filename
    web1.Navigate2 "file:///" & Filename
    txtCode.OpenFile Filename
    lblFile.Caption = Filename
    
    ListaLangs.SetLang Filename, txtCode
    txtCode_SelChange txtCode
    
    Debug.Print "Load : " & Me.Name
    
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    cmd(0).Refresh
    
    util.Hourglass hWnd, False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    TreeView.Nodes.Clear
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmViewTree = Nothing
End Sub

Private Sub mnuCloseXML_Click()

    Unload Me
    
End Sub




Private Sub mnuLoadXML_Click()
    
    Dim Archivo As String
    
    objXMLTree.CloseXMLFile
    
    ' Set CancelError is True
    If Cdlg.VBGetOpenFileName(Archivo, , , , , , "XML files (*.xml)|*.xml", , , "Open XML file", "xml", hWnd) Then
        txtCode.OpenFile Archivo
        objXMLTree.OpenXMLFile "file:///" & Archivo
        web1.Navigate2 "file:///" & Archivo
        
        Set objXMLTree = New XMLTree
        objXMLTree.SetTreeView TreeView
        
        Filename = Archivo
        lblFile.Caption = Filename
    End If
            
End Sub

Private Sub mnuSaveXML_Click()
    
    util.Hourglass hWnd, True
    
    txtCode.SaveFile Filename, False
    
    Set objXMLTree = Nothing
    
    Set objXMLTree = New XMLTree
    objXMLTree.SetTreeView TreeView
           
    objXMLTree.OpenXMLFile Filename
    web1.Navigate2 "file:///" & Filename
        
    util.Hourglass hWnd, False
    
End Sub

Private Sub tabMain_Click()

    If tabMain.SelectedItem.Index = 1 Then
        TreeView.ZOrder 0
    ElseIf tabMain.SelectedItem.Index = 2 Then
        web1.ZOrder 0
    Else
        txtCode.ZOrder 0
    End If
    
End Sub

Private Sub TreeView_AfterLabelEdit(Cancel As Integer, NewString As String)

    'you can edit the labels, they refer to XML attributes

    Set currTreeNode = TreeView.Nodes.ITem(objXMLTree.CurrKey)
    
    Select Case currTreeNode.Image
        Case 1, 2
    
        Case 3
    
        Case 4, 5
    
    End Select

End Sub

Private Sub TreeView_Collapse(ByVal node As MSComctlLib.node)
    'switches the open/close folder image
    Select Case node.Image
        Case 1
            node.Image = 2
        Case 2
            node.Image = 1
        Case 3
    End Select
End Sub

Private Sub TreeView_Expand(ByVal node As MSComctlLib.node)
    'switches the open/close folder image
    Select Case node.Image
        Case 1
            node.Image = 2
        Case 2
            node.Image = 1
        Case 3
    End Select
End Sub



Private Function txtCode_RClick(ByVal Control As CodeSenseCtl.ICodeSense) As Boolean
    txtCode_RClick = True
End Function


Private Sub txtCode_SelChange(ByVal Control As CodeSenseCtl.ICodeSense)
    
    On Error Resume Next
    
    Dim r As CodeSenseCtl.IRange
    Dim colorh As Long
    
    Set r = Control.GetSel(True)
        
    colorh = Control.GetColor(cmClrHighlightedLine)
    Call Control.SetColor(cmClrHighlightedLine, Control.GetColor(cmClrWindow))
    Control.HighlightedLine = r.StartLineNo
    DoEvents
    Call Control.SetColor(cmClrHighlightedLine, colorh)
   
    Set r = Nothing
   
    Err = 0
    
End Sub


Private Function txtCode_ShowProps(ByVal Control As CodeSenseCtl.ICodeSense) As Boolean
    txtCode_ShowProps = True
End Function


