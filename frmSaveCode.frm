VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSaveCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save Code File As..."
   ClientHeight    =   7605
   ClientLeft      =   2055
   ClientTop       =   2130
   ClientWidth     =   9525
   Icon            =   "frmSaveCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   9525
   Begin VB.TextBox txtName 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   3600
      TabIndex        =   10
      Top             =   6660
      Width           =   3450
   End
   Begin VB.Frame Frame4 
      Caption         =   "Save IN"
      Height          =   5955
      Left            =   30
      TabIndex        =   7
      Top             =   1140
      Width           =   3465
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1380
         Top             =   1470
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSaveCode.frx":038A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSaveCode.frx":0724
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSaveCode.frx":0ABE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSaveCode.frx":1058
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvLibrary 
         Height          =   5505
         Left            =   120
         TabIndex        =   8
         Top             =   330
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   9710
         _Version        =   393217
         Indentation     =   441
         LabelEdit       =   1
         Style           =   7
         HotTracking     =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   0
         MousePointer    =   99
         MouseIcon       =   "frmSaveCode.frx":13F2
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Height          =   75
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   3195
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         Height          =   5625
         Left            =   90
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   0
      TabIndex        =   5
      Top             =   900
      Width           =   9555
   End
   Begin VB.Frame Frame2 
      Height          =   5265
      Left            =   3540
      TabIndex        =   3
      Top             =   1140
      Width           =   5865
      Begin MSComctlLib.ListView lvFiles 
         Height          =   4875
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   8599
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         MousePointer    =   99
         MouseIcon       =   "frmSaveCode.frx":1554
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code Files"
            Object.Width           =   8996
         EndProperty
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         Height          =   4935
         Left            =   90
         Top             =   240
         Width           =   5685
      End
   End
   Begin VB.Frame Frame3 
      Height          =   645
      Left            =   7110
      TabIndex        =   0
      Top             =   6450
      Width           =   2265
      Begin VBSCodeLibrary.lvButtons_H lvSave 
         Height          =   375
         Left            =   150
         TabIndex        =   1
         Top             =   180
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         Caption         =   "Save"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VBSCodeLibrary.lvButtons_H lvCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   1110
         TabIndex        =   2
         Top             =   180
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   661
         Caption         =   "Cancel"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   9525
      TabIndex        =   6
      Top             =   0
      Width           =   9525
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save Code File Wizard"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   660
         TabIndex        =   13
         Top             =   180
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   420
         Left            =   150
         Picture         =   "frmSaveCode.frx":16B6
         Top             =   210
         Width           =   420
      End
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      Height          =   345
      Left            =   120
      Top             =   7170
      Width           =   9165
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C0C0&
      Height          =   375
      Left            =   3570
      Top             =   6630
      Width           =   3495
   End
   Begin VB.Label lblpath 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   7200
      Width           =   9165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Name:"
      Height          =   195
      Left            =   3600
      TabIndex        =   11
      Top             =   6450
      Width           =   750
   End
End
Attribute VB_Name = "frmSaveCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Folder As New FolderSysObject
Dim File As New FileSysObject
Dim prent As String

Private Sub Form_Load()

   CenterForm Me
   
   Call GetLanguages
   Me.tvLibrary.Nodes(1).Expanded = True
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmSaveCode = Nothing
End Sub


Private Sub lvCancel_Click()
Unload Me
End Sub

Private Sub GetLanguages()
Dim Languages() As String
Dim keyx As String
Dim i As Integer

On Error Resume Next
Languages = Folder.GetDirectories(DataPath)
With Me.tvLibrary.Nodes
     .Clear
     Set itn = .Add(, , "root", "Console Root", 2)
     For i = 1 To UBound(Languages)
         Set itn = .Add("root", tvwChild, Languages(i), Languages(i), 1)
         Call GetCategories(Languages(i))
     Next
End With
End Sub

Private Sub GetCategories(LanguageName As String)
Dim Categories() As String
Dim keys As String
Dim i As Integer

On Error Resume Next
Categories = Folder.GetDirectories(DataPath & LanguageName & "\")
With Me.tvLibrary.Nodes
     '.Clear
     'Create Root Node)
     'Set itn = Nothing
     
     For i = 1 To UBound(Categories)
         keys = LanguageName & Categories(i)
         Set itn = .Add(LanguageName, tvwChild, keys, Categories(i), 3)
     Next
     Set itn = Nothing
End With
End Sub

Private Sub GetFiles(Data As String)
Dim cFiles() As String
Dim keyx As String
Dim i As Integer

cFiles = File.GetFiles(DataPath & Data)
On Error Resume Next
With Me.lvFiles.ListItems
     .Clear
     For i = 1 To UBound(cFiles)
         keyx = Data & cFiles(i)
         Set itm = .Add(, , cFiles(i), 4, 4)
     Next
     Set itm = Nothing
     
End With
End Sub

Private Sub lvFiles_Click()
On Error Resume Next
Me.txtName.Text = Replace(lvFiles.SelectedItem.Text, fExt, "")
End Sub

Private Sub lvSave_Click()

   Dim c As String

   If txtName.Text = "" Then
      MsgBox "Please Type a Valid Name.", vbCritical, "No Filename Found."
      Exit Sub
   End If

   If ValidName(txtName.Text) = False Then
      MsgBox "Invalid FileName." & vbNewLine & _
          "---------------------------------------------------" & vbNewLine & _
          "a filename must not consist of any of the following" & vbNewLine & _
          Replace(invFileName, "x", " "), vbCritical, "No Filename Found."
      Exit Sub
   End If
   
   If lblpath.Caption = "" Then
      MsgBox "Please select a Location wherein the new code file will be saved.", vbInformation, "Unable to Save"
      Exit Sub
   End If

   If prent = "\Console Root" Then
      MsgBox "Please select a Folder Location wherein the new code file will be saved.", vbInformation, "Unable to Save"
      Exit Sub
   End If
   
   c = lblpath.Caption & "\" & txtName.Text & fExt

   If fso.FileExists(DataPath & c) = True Then
      Msg = MsgBox("The File '" & txtName.Text & "' Already Exists in the library." & vbNewLine & _
                "Replace Existing file?", vbQuestion + vbYesNo, "File Found")
      If Msg = vbYes Then
         Call SaveFileAs(DataPath, c, "")
         'mdiMain.ActiveForm.pathf = c
         'mdiMain.ActiveForm.Caption = File.GetFileName(c)
         'Call mdiMain.refreshTreeview
         Unload Me
      Else
         Unload Me
      End If
   Else
      Call SaveFileAs(DataPath, c, "")
      'mdiMain.ActiveForm.pathf = c
      'mdiMain.ActiveForm.Caption = File.GetFileName(c)
      'Call mdiMain.refreshTreeview
      Unload Me
   End If
   
End Sub

Private Sub tvLibrary_NodeClick(ByVal Node As MSComctlLib.Node)
Dim path As String

On Error Resume Next
prent = "\" & Node.Parent.Text
path = Replace(Node.FullPath, "Console Root", "")
'Me.Caption = path
Call GetFiles(path & "\")
lblpath.Caption = path
End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lvSave_Click
End Sub
