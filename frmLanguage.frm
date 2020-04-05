VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLanguage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Language Wizard"
   ClientHeight    =   5850
   ClientLeft      =   3060
   ClientTop       =   2670
   ClientWidth     =   5880
   Icon            =   "frmLanguage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   5880
   Begin VB.Frame Frame5 
      Caption         =   "Selected"
      Height          =   855
      Left            =   60
      TabIndex        =   7
      Top             =   4800
      Width           =   5745
      Begin VB.Label lblFound 
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
         Left            =   1680
         TabIndex        =   15
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categories Found:"
         Height          =   195
         Left            =   330
         TabIndex        =   14
         Top             =   480
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Language Name:"
         Height          =   195
         Left            =   330
         TabIndex        =   9
         Top             =   240
         Width           =   1230
      End
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
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   3255
      Left            =   4200
      TabIndex        =   2
      Top             =   1500
      Width           =   1605
      Begin VBSCodeLibrary.lvButtons_H cmdNew 
         Height          =   375
         Left            =   150
         TabIndex        =   3
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
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
         mPointer        =   99
      End
      Begin VBSCodeLibrary.lvButtons_H cmdRename 
         Height          =   375
         Left            =   150
         TabIndex        =   4
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
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
         Left            =   150
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
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
         Left            =   150
         TabIndex        =   6
         Top             =   2700
         Width           =   1215
         _ExtentX        =   2143
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
         Left            =   150
         TabIndex        =   16
         Top             =   2250
         Width           =   1215
         _ExtentX        =   2143
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
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -90
      TabIndex        =   1
      Top             =   1320
      Width           =   9255
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   5880
      TabIndex        =   0
      Top             =   0
      Width           =   5880
      Begin VB.Image Image2 
         Height          =   420
         Left            =   300
         Picture         =   "frmLanguage.frx":038A
         Top             =   510
         Width           =   420
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "This wizard allows you to rename, select and remove any language items found on the list"
         ForeColor       =   &H00FF0000&
         Height          =   645
         Left            =   780
         TabIndex        =   11
         Top             =   510
         Width           =   4425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LANGUAGE WIZARD"
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
         Left            =   300
         TabIndex        =   10
         Top             =   150
         Width           =   2085
      End
   End
   Begin MSComctlLib.ImageList imgLanguage 
      Left            =   3120
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLanguage.frx":0D74
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvLanguage 
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   1710
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   5318
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
      MouseIcon       =   "frmLanguage.frx":176E
      NumItems        =   0
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   4005
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   3225
      Left            =   90
      Top             =   1530
      Width           =   4065
   End
End
Attribute VB_Name = "frmLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Folder As New FolderSysObject
Dim File As New FileSysObject

Private Sub cmdNew_Click()
lblSelected.Caption = "None"
inp = InputBox("Type a New Language Name." & vbNewLine & _
               vbNewLine & _
               "Make sure that the new Language name is not yet on the Library.", _
               "Create NEW Language", "User Defined")

If ValidName(CStr(inp)) = False Then
   MsgBox "Invalid Language Name." & vbNewLine & _
          "---------------------------------------------------" & vbNewLine & _
          "a filename must not consist of any of the following" & vbNewLine & _
          Replace(invFileName, "x", " "), vbCritical, "No Filename Found."
   Exit Sub
End If

If CStr(inp) = Empty Then
   Exit Sub
Else
   If Toggle_Language([New Language], "", CStr(inp)) = False Then
      MsgBox "Sorry Unable to Add New Language." & _
             vbNewLine & _
             "This error is caused by either of the following." & vbNewLine & _
             "-----------------------------------------------------------------" & _
             vbNewLine & _
             "(1) A Language having the same name is found on the library." & _
             vbNewLine & _
             "(2) Unable to Create Language due to windows file and folder authorization.", vbCritical, _
             "Language Creation Error."
    Else
       MsgBox "Done", vbInformation
       Call GetLanguages
    End If
End If

End Sub

Private Sub cmdRemove_Click()
'If LCase(Me.lblFound.Caption) <> "none" Then
   Msg = MsgBox("Deleting this Selected Language will destroy all its category and code files." & _
           vbNewLine & "Are you sure?", vbQuestion + vbYesNo, "Confirm Delete")
   If Msg = vbYes Then
      If Toggle_Language([Remove Language], Me.lblSelected.Caption, "") = False Then
         MsgBox "Sorry Unable to DELETE Language." & _
             vbNewLine & _
             "This error is caused by either of the following." & vbNewLine & _
             "-----------------------------------------------------------------" & _
             vbNewLine & _
             "(1) language Folder not Found." & _
             vbNewLine & _
             "(2) A Language having the same name is found on the library." & _
             vbNewLine & _
             "(3) Unable to Create Language due to windows file and folder authorization.", vbCritical, _
             "Language Delete Error."
      Else
         MsgBox "Language Deleted", vbInformation, "Done"
         Call lvRefresh_Click
      End If
   End If
'End If
End Sub

Private Sub cmdRename_Click()

   If lblSelected.Caption = "" Or lblSelected.Caption = "None" Then
      MsgBox "Please Select the Language you want to rename.", vbExclamation, "No Selected Language"
      Exit Sub
   End If

   inp = InputBox("Type a NEW Name for [ " & lblSelected.Caption & " ]", _
               "Rename: " & lblSelected.Caption, lblSelected.Caption)

   If ValidName(CStr(inp)) = False Then
      MsgBox "Invalid Language Name." & vbNewLine & _
          "---------------------------------------------------" & vbNewLine & _
          "a filename must not consist of any of the following" & vbNewLine & _
          Replace(invFileName, "x", " "), vbCritical, "No Filename Found."
      Exit Sub
   End If

   If CStr(inp) = Empty Then
      lblSelected.Caption = ""
      Exit Sub
   
   Else
      If Toggle_Language([Rename Language], Me.lblSelected.Caption, CStr(inp)) = False Then
         MsgBox "Sorry Unable to Rename Language." & _
             vbNewLine & _
             "This error is caused by either of the following." & vbNewLine & _
             "-----------------------------------------------------------------" & _
             vbNewLine & _
             "(1) language Folder not Found." & _
             vbNewLine & _
             "(2) A Language having the same name is found on the library." & _
             vbNewLine & _
             "(3) Unable to Create Language due to windows file and folder authorization.", vbCritical, _
             "Language Rename Error."
    Else
       MsgBox "Rename Done", vbInformation, "Done"
       'If LCase(Me.lblSelected.Caption) = LCase(mdiMain.cboLanguage.Text) Then
       '   With mdiMain
       '        .cboLanguage.Text = CStr(inp)
       '   End With
       'End If
       Call GetLanguages
    End If
End If
Set inp = Nothing
lblSelected.Caption = ""
lblFound.Caption = ""
End Sub

Private Sub Form_Load()
   
   CenterForm Me
   
   Call GetLanguages
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmLanguage = Nothing
'On Error Resume Next
'Call mdiMain.GetLanguages
'Call mdiMain.refreshTreeview
End Sub

Private Sub lblSelected_Change()
If lblSelected.Caption = "" Then lblSelected.Caption = "None"
End Sub


Private Sub lvButtons_H4_Click()
Unload Me
End Sub

Private Sub GetLanguages()

   Dim Languages() As String
   Dim keyx As String
   Dim i As Integer

   On Error Resume Next
   Languages = Folder.GetDirectories(DataPath)
   
   With Me.lvLanguage.ListItems
      .Clear
      Set itm = .Add(, , Languages(1), 1, 1)
      For i = 2 To UBound(Languages)
         Set itm = .Add(, , Languages(i), 1, 1)
      Next
   End With
   
End Sub

Private Sub lvLanguage_Click()
Dim f() As String
On Error Resume Next
 
 Me.lblSelected.Caption = Me.lvLanguage.SelectedItem.Text
 
 f = Folder.GetDirectories(DataPath & Me.lblSelected.Caption)
 If UBound(f) = 0 Then
    Me.lblFound.Caption = "None"
 Else
    Me.lblFound.Caption = UBound(f)
 End If
End Sub

Private Sub lvRefresh_Click()
Call GetLanguages
Me.lblFound.Caption = "None"
Me.lblSelected.Caption = "None"
End Sub
