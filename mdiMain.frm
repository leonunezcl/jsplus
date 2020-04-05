VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   Caption         =   "Browse Library"
   ClientHeight    =   8775
   ClientLeft      =   1785
   ClientTop       =   1875
   ClientWidth     =   9795
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   ScaleHeight     =   8775
   ScaleWidth      =   9795
   WindowState     =   2  'Maximized
   Begin CodeSenseCtl.CodeSense CodeSense1 
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "mdiMain.frx":08CA
      TabIndex        =   5
      Top             =   120
      Width           =   5535
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   8775
      Left            =   5715
      ScaleHeight     =   8775
      ScaleWidth      =   4080
      TabIndex        =   0
      Top             =   0
      Width           =   4080
      Begin VB.Frame fraMid 
         Height          =   7035
         Left            =   60
         TabIndex        =   2
         Top             =   360
         Width           =   3975
         Begin MSComctlLib.ImageList imlTreview 
            Left            =   1320
            Top             =   2760
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
                  Picture         =   "mdiMain.frx":0A30
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":0DCA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "mdiMain.frx":1164
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   2070
            Top             =   360
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
                  Picture         =   "mdiMain.frx":14FE
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageCombo cboLanguage 
            Height          =   360
            Left            =   60
            TabIndex        =   4
            Top             =   210
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
            MouseIcon       =   "mdiMain.frx":1898
            Locked          =   -1  'True
            Text            =   "Language List"
            ImageList       =   "ImageList1"
         End
         Begin MSComctlLib.TreeView tvLibrary 
            Height          =   6345
            Left            =   60
            TabIndex        =   3
            Top             =   630
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   11192
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
            MouseIcon       =   "mdiMain.frx":19FA
         End
      End
      Begin VB.Frame fraTop 
         Height          =   435
         Left            =   60
         TabIndex        =   1
         Top             =   -90
         Width           =   3975
      End
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mSaveAs 
         Caption         =   "Save As.."
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mOther 
      Caption         =   "Other"
      Visible         =   0   'False
      Begin VB.Menu mLW 
         Caption         =   "Language Wizard"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mCW 
         Caption         =   "Category Wizard"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mSS 
         Caption         =   "Security Settings"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mAbout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Folder As New FolderSysObject
Dim File As New FileSysObject

Private Sub cboLanguage_Change()
On Error Resume Next
Call GetCategories(Me.cboLanguage.SelectedItem.Text)
Call GetCategories(Me.cboLanguage.Text)
End Sub

Private Sub cboLanguage_Click()
On Error Resume Next
Call GetCategories(Me.cboLanguage.SelectedItem.Text)
Call GetCategories(Me.cboLanguage.Text)
End Sub

Private Sub Form_Load()
   Call GetLanguages
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmMain = Nothing
End Sub


Private Sub Picture1_Resize()
On Error Resume Next
    Me.fraMid.Width = Me.fraTop.Width
    Me.fraMid.Height = Me.Picture1.Height - 380
    Me.tvLibrary.Height = Me.fraMid.Height - 700
End Sub

Public Sub GetLanguages()
Dim Languages() As String
Dim keyx As String
Dim i As Integer

On Error Resume Next
Languages = Folder.GetDirectories(DataPath)
With Me.cboLanguage.ComboItems
     .Clear
     For i = 1 To UBound(Languages)
         .Add , Languages(i), Languages(i), 1
     Next
End With
Me.tvLibrary.Nodes(1).Expanded = True
End Sub

Public Sub GetCategories(LanguageName As String)
Dim Categories() As String
Dim keys As String
Dim i As Integer

Categories = Folder.GetDirectories(DataPath & LanguageName & "\")
With Me.tvLibrary.Nodes
     .Clear
     'Create Root Node
     Set itn = .Add(, , "root", "Console Root", 1)
     Set itn = Nothing
     
     For i = 1 To UBound(Categories)
         keys = LanguageName & Categories(i)
         Set itn = .Add("root", tvwChild, keys, Categories(i), 2)
         Call GetFiles(LanguageName, Categories(i))
     Next
     Set itn = Nothing
End With
Me.tvLibrary.Nodes(1).Expanded = True
End Sub

Public Sub GetFiles(LanguageName, Category As String)
Dim cFiles() As String
Dim keyx As String
Dim i As Integer

cFiles = File.GetFiles(DataPath & LanguageName & "\" & Category & "\")
On Error Resume Next
Set itn = Nothing
With Me.tvLibrary.Nodes

     For i = 1 To UBound(cFiles)
         keyx = LanguageName & Category & cFiles(i)
         Set itn = .Add(LanguageName & Category, tvwChild, keyx, cFiles(i), 3)
     Next
     Set itn = Nothing
     
End With
End Sub

Public Sub refreshTreeview()
Call cboLanguage_Click
Call cboLanguage_Change
End Sub
