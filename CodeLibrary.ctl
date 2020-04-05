VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2128BF45-F895-4206-84CD-F4DE2DD8D6B1}#2.0#0"; "vbsTbar6.ocx"
Object = "{98F993CC-3598-405A-9E9A-0D2CF198B250}#2.0#0"; "vbsDkTb6.ocx"
Begin VB.UserControl CodeLibrary 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
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
      TabIndex        =   2
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
   Begin MSComctlLib.ImageList imlTreview 
      Left            =   1455
      Top             =   2910
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeLibrary.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeLibrary.ctx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CodeLibrary.ctx":0734
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2625
      Top             =   2895
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
            Picture         =   "CodeLibrary.ctx":0ACE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo cboLanguage 
      Height          =   360
      Left            =   270
      TabIndex        =   0
      Top             =   975
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   16777215
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "CodeLibrary.ctx":0E68
      Locked          =   -1  'True
      Text            =   "Language List"
      ImageList       =   "ImageList1"
   End
   Begin MSComctlLib.TreeView tvLibrary 
      Height          =   1230
      Left            =   195
      TabIndex        =   1
      ToolTipText     =   "Double clic to open code"
      Top             =   1440
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   2170
      _Version        =   393217
      Indentation     =   617
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imlTreview"
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "CodeLibrary.ctx":0FCA
   End
   Begin vbalDkTb6.vbalDockContainer vbalDockContainer1 
      Align           =   1  'Align Top
      Height          =   30
      Left            =   0
      TabIndex        =   3
      Top             =   375
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   53
      AllowUndock     =   0   'False
      LockToolbars    =   -1  'True
   End
End
Attribute VB_Name = "CodeLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private CodeLib As New cCodeLibrary
Private m_Img As cVBALImageList
Private WithEvents m_cMenu As cPopupMenu
Attribute m_cMenu.VB_VarHelpID = -1

Public Event FileSelected(ByVal Archivo As String)

Private Sub cboLanguage_Change()
    On Error Resume Next
    Call CodeLib.GetCategories(cboLanguage.SelectedItem.Text, tvLibrary)
    Call CodeLib.GetCategories(cboLanguage.Text, tvLibrary)
    Err = 0
End Sub


Public Sub Prepare()
    
    Dim iMain As Long
    Dim ip As Long
    
    Set m_Img = New cVBALImageList
    
    With m_Img
        .IconSizeX = 16: .IconSizeY = 16: .ColourDepth = ILC_COLOR24
        .Create
        .AddFromResourceID 286, App.hInstance, IMAGE_ICON, "k1"
        .AddFromResourceID 287, App.hInstance, IMAGE_ICON, "k2"
        .AddFromResourceID 288, App.hInstance, IMAGE_ICON, "k3"
    End With
    
    Set m_cMenu = New cPopupMenu
    m_cMenu.hWndOwner = UserControl.hwnd
    m_cMenu.OfficeXpStyle = True
    m_cMenu.ImageList = m_Img.hIml

    With m_cMenu
        'tools
        iMain = .AddItem("TOOLS", "Tools Toolbar", , , , , , "TOOLSTOOLBAR")
        ip = .AddItem("Load", "Load Library", , iMain, 0, , , "LIBRARY:LOAD")
        ip = .AddItem("Refresh", "Refresh Library", , iMain, 1, , , "LIBRARY:REFRESH")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Open", "Open Code", , iMain, 2, , , "LIBRARY:OPEN")
        ip = .AddItem("-", , , iMain)
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


Private Sub cboLanguage_Click()
   On Error Resume Next
   Call CodeLib.GetCategories(cboLanguage.SelectedItem.Text, tvLibrary)
   Call CodeLib.GetCategories(cboLanguage.Text, tvLibrary)
   Err = 0
End Sub

Private Sub tbrTools_ButtonClick(ByVal lButton As Long)
    
   util.Hourglass hwnd, True
   
   Select Case tbrTools.ButtonKey(lButton)
      Case "LIBRARY:LOAD"
         Call CodeLib.GetLanguages(cboLanguage, tvLibrary)
         If cboLanguage.ComboItems.count > 0 Then
            cboLanguage.ComboItems(1).Selected = True
         End If
      Case "LIBRARY:REFRESH"
         cboLanguage.ComboItems.Clear
         tvLibrary.Nodes.Clear
         Call CodeLib.GetLanguages(cboLanguage, tvLibrary)
         If cboLanguage.ComboItems.count > 0 Then
            cboLanguage.ComboItems(1).Selected = True
         End If
      Case "LIBRARY:OPEN"
         Call tvLibrary_DblClick
   End Select
   
   util.Hourglass hwnd, False
    
End Sub

Private Sub tvLibrary_DblClick()

   Dim Archivo As String
       
   If Not tvLibrary.SelectedItem Is Nothing Then
      If tvLibrary.SelectedItem.Tag = "File" Then
         Archivo = CodeLib.DataPath & Replace(tvLibrary.SelectedItem.FullPath, "Console Root", cboLanguage.Text)
    
         If ArchivoExiste2(Archivo) Then
            If Confirma("Do you want to open the selected code into a new window editor") = vbYes Then
               RaiseEvent FileSelected(Archivo)
            Else
               If Not frmMain.ActiveForm Is Nothing Then
                    If InStr(frmMain.ActiveForm.Caption, "*") > 0 Then
                        If Confirma("Do you want to save the changes") = vbYes Then
                            Files.Save frmMain.ActiveForm, False
                        End If
                    End If
               Else
                  Exit Sub
               End If
            End If
            'abrir, cambiar caption y setear el archivo a blanco ...
            frmMain.ActiveForm.txtCode.OpenFile Archivo
            'frmMain.ActiveForm.Caption = util.VBArchivoSinPath(Archivo)
            Call Files.ClearFileName(frmMain.ActiveForm)
         End If
      End If
   End If
    
End Sub


Private Sub UserControl_Initialize()
   CodeLib.DataPath = util.StripPath(App.Path)
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
   
   LockWindowUpdate hwnd
   cboLanguage.Move 0, picGeneral.Height + 1, UserControl.Width, UserControl.Height - picGeneral.Height - 245
   tvLibrary.Move 0, picGeneral.Height + 10 + 345, UserControl.Width, UserControl.Height - picGeneral.Height - 400
   LockWindowUpdate False
   
   Err = 0
   
End Sub


