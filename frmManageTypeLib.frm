VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageTypeLib 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ActiveX Object Browser"
   ClientHeight    =   6435
   ClientLeft      =   1560
   ClientTop       =   2385
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManageTypeLib.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   11820
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6210
      Visible         =   0   'False
      Width           =   2250
   End
   Begin MSComctlLib.ListView lvwCom 
      Height          =   5595
      Left            =   30
      TabIndex        =   0
      Top             =   285
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   9869
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nº"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "GUID"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Version"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Path"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Exists"
         Object.Width           =   1764
      EndProperty
   End
   Begin jsplus.HeaderPicture hpic 
      Height          =   255
      Index           =   0
      Left            =   15
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   450
      Alignment       =   0
      Caption         =   "Double click to insert object in active document"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontSize        =   8
      FontColor       =   0
      GradientStart   =   16761024
      Picture         =   "frmManageTypeLib.frx":000C
   End
   Begin jsplus.MyButton cmdInsertar 
      Height          =   405
      Left            =   3660
      TabIndex        =   3
      Top             =   5940
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Insert"
      AccessKey       =   "I"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePos      =   3
   End
   Begin jsplus.MyButton cmdOk 
      Height          =   405
      Left            =   6600
      TabIndex        =   4
      Top             =   5940
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Cancel"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePos      =   3
   End
End
Attribute VB_Name = "frmManageTypeLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Populate()
Dim iSectCount As Long, iSect As Long, sSections() As String
Dim iVerCount As Long, iVer As Long, sVersions() As String
Dim iExeSectCount As Long, sExeSect() As String
Dim iExeSect As Long
Dim bFoundExeSect As Boolean
Dim sExists As String
Dim cTLI As cTypeLibInfo
'Dim i As IShellFolderEx_TLB.IUnknown

Dim c As Integer

   'pClearList
   'lstTypeLibs.Clear
   'lstTypeLibs.Visible = False

   Dim cR As New cRegistry
   cR.ClassKey = HKEY_CLASSES_ROOT
   cR.ValueType = REG_SZ
   cR.SectionKey = "TypeLib"
   
   c = 1
   ' Get the registered Type Libs:
   If cR.EnumerateSections(sSections(), iSectCount) Then
      For iSect = 1 To iSectCount
         ' Enumerate the versions for each typelib:
         cR.SectionKey = "TypeLib\" & sSections(iSect)
         If cR.EnumerateSections(sVersions(), iVerCount) Then
            For iVer = 1 To iVerCount
               Set cTLI = New cTypeLibInfo
               cTLI.CLSID = sSections(iSect)
               cTLI.Ver = sVersions(iVer)
               cR.SectionKey = "TypeLib\" & sSections(iSect) & "\" & sVersions(iVer)
               cTLI.Name = cR.Value
               cR.EnumerateSections sExeSect(), iExeSectCount
               If iExeSectCount > 0 Then
                  bFoundExeSect = False
                  For iExeSect = 1 To iExeSectCount
                     If IsNumeric(sExeSect(iExeSect)) Then
                        cR.SectionKey = cR.SectionKey & "\" & sExeSect(iExeSect) & "\win32"
                        bFoundExeSect = True
                        Exit For
                     End If
                  Next iExeSect
                  If bFoundExeSect Then
                     cTLI.Path = cR.Value
                     If FileExists(cTLI.Path) Then
                        sExists = "Y"
                     Else
                        sExists = "N"
                     End If
                  Else
                     sExists = "N"
                  End If
               Else
                  sExists = "N"
               End If
               cTLI.Exists = (StrComp(sExists, "Y") = 0)
               lvwCom.ListItems.Add , "k" & c, c
               If Len(cTLI.Name) > 0 Then
                  lvwCom.ListItems(c).SubItems(1) = cTLI.Name
                  'lstTypeLibs.AddItem cTLI.Name & vbTab & sExists
               Else
                  lvwCom.ListItems(c).SubItems(1) = cTLI.CLSID
                  'lstTypeLibs.AddItem cTLI.CLSID & vbTab & sExists
               End If
               lvwCom.ListItems(c).SubItems(2) = cTLI.CLSID
               lvwCom.ListItems(c).SubItems(3) = cTLI.Ver
               lvwCom.ListItems(c).SubItems(4) = cTLI.Path
               lvwCom.ListItems(c).SubItems(5) = sExists
               c = c + 1
               'lstTypeLibs.ItemData(lstTypeLibs.NewIndex) = ObjPtr(cTLI)
               'Set i = cTLI
               'i.AddRef
            Next iVer
         End If
      Next iSect
   End If
   
   'lstTypeLibs.Visible = True
   
End Sub



Private Sub cmdInsertar_Click()

    Dim Texto As String
    Dim itmx As ListItem
    
    If Not lvwCom.SelectedItem Is Nothing Then
        Set itmx = lvwCom.SelectedItem
        Texto = "<OBJECT ID=" & Chr$(34) & itmx.SubItems(1) & Chr$(34)
        Texto = Texto & " WIDTH=" & Chr$(34) & Chr$(34) & " "
        Texto = Texto & " HEIGHT=" & Chr$(34) & Chr$(34) & " "
        Texto = Texto & " CLASSID=" & Chr$(34) & itmx.SubItems(2) & Chr$(34) & ">"
        Call frmMain.ActiveForm.Insertar(Texto)
    End If
    
End Sub


Private Sub cmdOk_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Util.Hourglass hwnd, True
    
    Util.CenterForm Me
    
    Populate
        
    flat_lviews Me
    
    Util.Hourglass hwnd, False
    
    DrawXPCtl Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmManageTypeLib = Nothing
End Sub




Private Sub lvwCom_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If lvwCom.SortOrder = 0 Then
        lvwCom.SortKey = ColumnHeader.Index - 1
        lvwCom.SortOrder = 1
    Else   ' Set Sorted to True to sort the list.
        lvwCom.SortKey = ColumnHeader.Index - 1
        lvwCom.SortOrder = 0
    End If
    lvwCom.Sorted = True
    
End Sub


Private Sub lvwCom_DblClick()
    cmdInsertar_Click
End Sub

