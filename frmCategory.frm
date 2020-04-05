VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Category Wizard"
   ClientHeight    =   6225
   ClientLeft      =   1995
   ClientTop       =   3090
   ClientWidth     =   7245
   Icon            =   "frmCategory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCategory.frx":038A
   ScaleHeight     =   6225
   ScaleWidth      =   7245
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   7245
      TabIndex        =   11
      Top             =   0
      Width           =   7245
      Begin MSComctlLib.ImageList imgLanguage 
         Left            =   4650
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCategory.frx":0C54
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5640
         Top             =   630
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
               Picture         =   "frmCategory.frx":152E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image Image2 
         Height          =   420
         Left            =   240
         Picture         =   "frmCategory.frx":18C8
         Top             =   540
         Width           =   420
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "frmCategory.frx":22B2
         Top             =   450
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CATEGORY WIZARD"
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
         Left            =   150
         TabIndex        =   13
         Top             =   180
         Width           =   2085
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "This wizard allows you to rename, select and remove any category items found on the list"
         ForeColor       =   &H00FF0000&
         Height          =   645
         Left            =   840
         TabIndex        =   12
         Top             =   510
         Width           =   4425
      End
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -30
      TabIndex        =   10
      Top             =   1350
      Width           =   9255
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   90
      TabIndex        =   8
      Top             =   1530
      Width           =   5415
      Begin MSComctlLib.ListView lvCategory 
         Height          =   2865
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5054
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "imgLanguage"
         SmallIcons      =   "imgLanguage"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         MousePointer    =   99
         MouseIcon       =   "frmCategory.frx":2B7C
         NumItems        =   0
      End
      Begin MSComctlLib.ImageCombo cboLanguage 
         Height          =   360
         Left            =   90
         TabIndex        =   14
         Top             =   420
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         MouseIcon       =   "frmCategory.frx":2CDE
         Indentation     =   1
         Locked          =   -1  'True
         Text            =   "Language List"
         ImageList       =   "ImageList1"
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   870
         Width           =   5175
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   3075
         Left            =   90
         Top             =   840
         Width           =   5235
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Langauge List"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   180
         Width           =   1005
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   4605
      Left            =   5550
      TabIndex        =   3
      Top             =   1530
      Width           =   1635
      Begin VBSCodeLibrary.lvButtons_H cmdNew 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   405
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         Caption         =   "New"
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
      Begin VBSCodeLibrary.lvButtons_H cmdRename 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         Caption         =   "Rename"
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
      Begin VBSCodeLibrary.lvButtons_H cmdRemove 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         Caption         =   "Remove"
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
      Begin VBSCodeLibrary.lvButtons_H lvButtons_H4 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   4110
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         Caption         =   "Close"
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
      Begin VBSCodeLibrary.lvButtons_H lvRefresh 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   3690
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         Caption         =   "Refresh"
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
   Begin VB.Frame Frame5 
      Caption         =   "Selected"
      Height          =   585
      Left            =   90
      TabIndex        =   0
      Top             =   5550
      Width           =   5415
      Begin VB.Label lblSelected 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Folder As New FolderSysObject
Dim File As New FileSysObject

Private Sub GetCategories(LanguageName As String)
Dim Categories() As String
Dim keys As String
Dim i As Integer

On Error Resume Next
Categories = Folder.GetDirectories(DataPath & LanguageName & "\")

With Me.lvCategory.ListItems
     .Clear
     Set itm = .Add(, , Categories(1), 1, 1)
     For i = 2 To UBound(Categories)
         Set itm = .Add(, , Categories(i), 1, 1)
     Next
End With
End Sub

Private Sub cmdNew_Click()
Dim s As String

If CheckSelectedLanguage = False Then
   MsgBox "Please Select a Language category first.", vbCritical, "No Selected Language"
   Me.cboLanguage.SetFocus
   Exit Sub
End If
lblSelected.Caption = "None"
inp = InputBox("Type a New Category Name." & vbNewLine & _
               vbNewLine & _
               "Make sure that the new Category name is not yet on the Library.", _
               "Create NEW Category", "User Defined")

If ValidName(CStr(inp)) = False Then
   MsgBox "Invalid Category Name." & vbNewLine & _
          "---------------------------------------------------" & vbNewLine & _
          "a filename must not consist of any of the following" & vbNewLine & _
          Replace(invFileName, "x", " "), vbCritical, "No Filename Found."
   Exit Sub
End If

If CStr(inp) = Empty Then
   Exit Sub
Else
   s = Me.cboLanguage.Text
   If Toggle_Category([New Category], s, "", CStr(inp)) = False Then
      MsgBox "Sorry Unable to Add New Category." & _
             vbNewLine & _
             "This error is caused by either of the following." & vbNewLine & _
             "-----------------------------------------------------------------" & _
             vbNewLine & _
             "(1) A Category having the same name is found on the library." & _
             vbNewLine & _
             "(2) Unable to Create Category due to windows file and folder authorization.", vbCritical, _
             "Language Creation Error."
    Else
       MsgBox "Done", vbInformation
       Call cboLanguage_Click
       Call cboLanguage_Change
    End If
End If
End Sub



Private Sub cmdRemove_Click()

Msg = MsgBox("Deleting this Selected Category will destroy all code files in it." & _
           vbNewLine & "Are you sure?", vbQuestion + vbYesNo, "Confirm Delete")
           
   If Msg = vbYes Then
      If Toggle_Category([Remove Category], Me.cboLanguage.Text, Me.lblSelected.Caption, "") = False Then
         MsgBox "Sorry Unable to DELETE Category." & _
             vbNewLine & _
             "This error is caused by either of the following." & vbNewLine & _
             "-----------------------------------------------------------------" & _
             vbNewLine & _
             "(1) Category Folder not Found." & _
             vbNewLine & _
             "(2) A Category having the same name is found on the library." & _
             vbNewLine & _
             "(3) Unable to Create Category due to windows file and folder authorization.", vbCritical, _
             "Language Delete Error."
      Else
         MsgBox "Category Deleted", vbInformation, "Done"
         Me.lvCategory.ListItems.Clear
         Call cboLanguage_Change
         Call cboLanguage_Click
         lblSelected.Caption = "None"
         
      End If
   End If
End Sub

Private Sub cmdRename_Click()
Dim s As String

If CheckSelectedLanguage = False Then
   MsgBox "Please Select a Language category first.", vbCritical, "No Selected Language"
   Me.cboLanguage.SetFocus
   Exit Sub
End If

If lblSelected.Caption = "" Or lblSelected.Caption = "None" Then
   MsgBox "Please Select the Category you want to rename.", vbExclamation, "No Selected Category"
   Exit Sub
End If

inp = InputBox("Type a NEW Name for [ " & lblSelected.Caption & " ]", _
               "Rename: " & lblSelected.Caption, lblSelected.Caption)

If ValidName(CStr(inp)) = False Then
   MsgBox "Invalid Category Name." & vbNewLine & _
          "---------------------------------------------------" & vbNewLine & _
          "a filename must not consist of any of the following" & vbNewLine & _
          Replace(invFileName, "x", " "), vbCritical, "No Filename Found."
   Exit Sub
End If

If CStr(inp) = Empty Then
   lblSelected.Caption = "None"
   Exit Sub
Else
   s = Me.cboLanguage.Text
   If Toggle_Category([Rename Category], s, Me.lblSelected.Caption, CStr(inp)) = False Then
      MsgBox "Sorry Unable to Rename Category." & _
             vbNewLine & _
             "This error is caused by either of the following." & vbNewLine & _
             "-----------------------------------------------------------------" & _
             vbNewLine & _
             "(1) Category Folder not Found." & _
             vbNewLine & _
             "(2) A Category having the same name is found on the library." & _
             vbNewLine & _
             "(3) Unable to Create Category due to windows file and folder authorization.", vbCritical, _
             "Category Rename Error."
    Else
       MsgBox "Done"
       'If LCase(Me.cboLanguage.Text) = LCase(mdiMain.cboLanguage.Text) Then
       '   With mdiMain
       '        .cboLanguage.Text = Me.cboLanguage.Text
       '   End With
       'End If
       Call cboLanguage_Click
       Call cboLanguage_Change
    End If
End If
Set inp = Nothing
lblSelected.Caption = "None"
End Sub

Private Sub Form_Load()
      
   CenterForm Me
   
   Call GetLanguages
   
End Sub

Private Sub cboLanguage_Change()
On Error Resume Next
Me.Caption = (Me.cboLanguage.SelectedItem.Text)
End Sub

Private Sub cboLanguage_Click()
On Error Resume Next
Call GetCategories(Me.cboLanguage.SelectedItem.Text)
End Sub

Private Sub GetLanguages()
Dim Languages() As String
Dim keyx As String
Dim i As Integer

'On Error Resume Next

Languages = Folder.GetDirectories(DataPath)
With Me.cboLanguage.ComboItems
     .Clear
     For i = 1 To UBound(Languages)
         .Add , Languages(i), Languages(i), 1
     Next
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmCategory = Nothing
End Sub

Private Sub lvButtons_H4_Click()
   Unload Me
End Sub

Private Sub lvCategory_Click()
On Error Resume Next
 
 Me.lblSelected.Caption = Me.lvCategory.SelectedItem.Text
End Sub

Public Function CheckSelectedLanguage() As Boolean
Call cboLanguage_Click
Call cboLanguage_Change
With Me.cboLanguage
     If .Text = "" Or .Text = "Language List" Or .Text = "Select Language" Then
        CheckSelectedLanguage = False
     Else
        CheckSelectedLanguage = True
     End If
End With
End Function

Private Sub lvCategory_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call lvCategory_Click

End Sub

Private Sub lvRefresh_Click()
Me.cboLanguage.ComboItems.Clear
Me.cboLanguage.Text = "Select Language"
Call GetLanguages
lblSelected.Caption = "None"
Me.lvCategory.ListItems.Clear
Me.Caption = "Category Wizard"
End Sub
