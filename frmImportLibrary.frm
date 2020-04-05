VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImportLibrary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create New Library ..."
   ClientHeight    =   7740
   ClientLeft      =   2940
   ClientTop       =   1815
   ClientWidth     =   9030
   ControlBox      =   0   'False
   Icon            =   "frmImportLibrary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Previous"
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   3000
      TabIndex        =   47
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Next"
      Height          =   375
      Index           =   6
      Left            =   4680
      TabIndex        =   46
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Step 2 - Select Functions to Import"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7020
      Index           =   1
      Left            =   10440
      TabIndex        =   1
      Top             =   360
      Width           =   8805
      Begin VB.CommandButton cmd 
         Caption         =   "Reload"
         Height          =   375
         Index           =   17
         Left            =   7560
         TabIndex        =   57
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Remove All"
         Height          =   375
         Index           =   16
         Left            =   7560
         TabIndex        =   56
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Remove"
         Height          =   375
         Index           =   15
         Left            =   7560
         TabIndex        =   55
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Remove All"
         Height          =   375
         Index           =   14
         Left            =   7560
         TabIndex        =   54
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Remove"
         Height          =   375
         Index           =   4
         Left            =   7560
         TabIndex        =   44
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Add"
         Height          =   375
         Index           =   3
         Left            =   7560
         TabIndex        =   43
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CheckBox chkSelAllFun 
         Caption         =   "Select All"
         Height          =   195
         Left            =   6390
         TabIndex        =   34
         Top             =   270
         Width           =   990
      End
      Begin VB.ListBox lstCusFun 
         Height          =   2985
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   3780
         Width           =   7275
      End
      Begin VB.ListBox lstFun 
         Height          =   2985
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   495
         Width           =   7260
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Custom members to add :"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   3540
         Width           =   1785
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Select functions to import :"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Width           =   1875
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   2
      Left            =   6360
      TabIndex        =   42
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   41
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Step 3 - Member Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6930
      Index           =   2
      Left            =   10440
      TabIndex        =   4
      Top             =   360
      Width           =   8805
      Begin VB.CommandButton cmd 
         Caption         =   "Multi Assing"
         Height          =   375
         Index           =   22
         Left            =   7080
         TabIndex        =   62
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Reload Members"
         Height          =   375
         Index           =   19
         Left            =   5400
         TabIndex        =   59
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Update Information"
         Height          =   375
         Index           =   18
         Left            =   3720
         TabIndex        =   58
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CheckBox chkSelAllMem 
         Caption         =   "Select All"
         Height          =   195
         Left            =   2610
         TabIndex        =   38
         Top             =   285
         Width           =   990
      End
      Begin MSComctlLib.ListView lvwSelFun 
         Height          =   6150
         Left            =   120
         TabIndex        =   35
         Top             =   495
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   10848
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Member"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtHelp 
         Height          =   2265
         Left            =   3690
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   3690
         Width           =   4995
      End
      Begin VB.TextBox txtDeclaration 
         Height          =   2265
         Left            =   3690
         Locked          =   -1  'True
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   495
         Width           =   4995
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   3690
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3075
         Width           =   4995
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Method = Default"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   37
         Top             =   6675
         Width           =   1485
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Member Help"
         Height          =   195
         Index           =   6
         Left            =   3690
         TabIndex        =   12
         Top             =   3450
         Width           =   945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Member declaration:"
         Height          =   195
         Index           =   5
         Left            =   3690
         TabIndex        =   10
         Top             =   300
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Member Type"
         Height          =   195
         Index           =   4
         Left            =   3690
         TabIndex        =   8
         Top             =   2820
         Width           =   975
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Selected members to import :"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   2040
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Step 4 - Organize Library"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6780
      Index           =   3
      Left            =   10440
      TabIndex        =   14
      Top             =   360
      Width           =   8805
      Begin VB.CommandButton cmd 
         Caption         =   "Auto"
         Height          =   375
         Index           =   21
         Left            =   3960
         TabIndex        =   61
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Reload"
         Height          =   375
         Index           =   20
         Left            =   3960
         TabIndex        =   60
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Quit All"
         Height          =   375
         Index           =   12
         Left            =   3960
         TabIndex        =   52
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Add All"
         Height          =   375
         Index           =   11
         Left            =   3960
         TabIndex        =   51
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Quit"
         Height          =   375
         Index           =   10
         Left            =   3960
         TabIndex        =   50
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Add"
         Height          =   375
         Index           =   5
         Left            =   3960
         TabIndex        =   45
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox chkMembers 
         Caption         =   "Select All"
         Height          =   195
         Left            =   2880
         TabIndex        =   36
         Top             =   300
         Width           =   990
      End
      Begin MSComctlLib.ListView lvwMembers 
         Height          =   6120
         Left            =   120
         TabIndex        =   19
         Top             =   510
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   10795
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Member"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwLibrary 
         Height          =   6105
         Left            =   4935
         TabIndex        =   17
         Top             =   525
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   10769
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Final Library Design"
         Height          =   195
         Index           =   10
         Left            =   4935
         TabIndex        =   20
         Top             =   270
         Width           =   1380
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Members"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   18
         Top             =   270
         Width           =   645
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Step 5 - Library Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7050
      Index           =   4
      Left            =   10440
      TabIndex        =   21
      Top             =   360
      Width           =   8805
      Begin VB.CheckBox chkActive 
         Caption         =   "Active IDE  (You must restart to changes take effect)"
         Height          =   285
         Left            =   105
         TabIndex        =   39
         Top             =   3945
         Width           =   4050
      End
      Begin VB.TextBox txtLibName 
         Height          =   345
         Left            =   105
         MaxLength       =   255
         TabIndex        =   32
         Top             =   825
         Width           =   8565
      End
      Begin VB.TextBox txtComments 
         Height          =   2385
         Left            =   105
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   4530
         Width           =   8565
      End
      Begin VB.TextBox txtUrl 
         Height          =   345
         Left            =   105
         MaxLength       =   255
         TabIndex        =   28
         Top             =   3540
         Width           =   8565
      End
      Begin VB.TextBox txtVersion 
         Height          =   345
         Left            =   105
         MaxLength       =   255
         TabIndex        =   26
         Top             =   2880
         Width           =   8565
      End
      Begin VB.TextBox txtDescription 
         Height          =   345
         Left            =   105
         MaxLength       =   255
         TabIndex        =   24
         Top             =   2220
         Width           =   8565
      End
      Begin VB.TextBox txtAutor 
         Height          =   345
         Left            =   105
         MaxLength       =   255
         TabIndex        =   22
         Top             =   1545
         Width           =   8565
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Library Name"
         Height          =   195
         Index           =   17
         Left            =   105
         TabIndex        =   33
         Top             =   585
         Width           =   930
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Index           =   16
         Left            =   90
         TabIndex        =   31
         Top             =   4275
         Width           =   735
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Home Page"
         Height          =   195
         Index           =   15
         Left            =   105
         TabIndex        =   29
         Top             =   3300
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Version"
         Height          =   195
         Index           =   14
         Left            =   105
         TabIndex        =   27
         Top             =   2640
         Width           =   525
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Index           =   13
         Left            =   105
         TabIndex        =   25
         Top             =   1980
         Width           =   795
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Autor"
         Height          =   195
         Index           =   11
         Left            =   105
         TabIndex        =   23
         Top             =   1305
         Width           =   375
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Step 1 - Select Source Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7050
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   8880
      Begin VB.CommandButton cmd 
         Caption         =   "Remove All"
         Height          =   375
         Index           =   13
         Left            =   7680
         TabIndex        =   53
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Remove"
         Height          =   375
         Index           =   9
         Left            =   7680
         TabIndex        =   49
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Add Folder"
         Height          =   375
         Index           =   8
         Left            =   7680
         TabIndex        =   48
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Add File"
         Height          =   375
         Index           =   0
         Left            =   7680
         TabIndex        =   40
         Top             =   600
         Width           =   975
      End
      Begin VB.ListBox lstSelFiles 
         Height          =   6135
         Left            =   150
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   525
         Width           =   7335
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Selected Files:"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   15
         Top             =   285
         Width           =   1035
      End
   End
   Begin MSComctlLib.ImageList imgAyuda 
      Left            =   240
      Top             =   7095
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":000C
            Key             =   "OBJECT"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":0640
            Key             =   "PROPERTY"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":095A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":0C74
            Key             =   "EVENT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":0F8E
            Key             =   "CONSTANT"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":12A8
            Key             =   "FUNCTION"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":15C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":18DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":1A36
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":1B90
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":212A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":2284
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":23DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":2538
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportLibrary.frx":2692
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmImportLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Indice As Integer
Private LastPath As String
Private cExpFun As New cExpFunctions

Private Type eLibrary
    nombre As String
    Declaracion As String
    tipo As Integer
    ayuda As String
End Type
Private arr_member_info() As eLibrary

'indices
Private idx_object As Integer
Private idx_property As Integer
Private idx_method As Integer
Private idx_event As Integer
Private idx_constant As Integer
Private idx_collection As Integer

Public Archivo As String

Private Const member_limit = 5000

Public Sub ActualizaInfo(ByVal ListIndex As Integer, ByVal Tag As Integer, ByVal key As String)

    arr_member_info(Tag).tipo = ListIndex
                    
    If ListIndex = 0 Then
        lvwSelFun.ListItems(key).SubItems(1) = "Object"
    ElseIf ListIndex = 1 Then
        lvwSelFun.ListItems(key).SubItems(1) = "Property"
    ElseIf ListIndex = 2 Then
        lvwSelFun.ListItems(key).SubItems(1) = "Method"
    ElseIf ListIndex = 3 Then
        lvwSelFun.ListItems(key).SubItems(1) = "Event"
    ElseIf ListIndex = 4 Then
        lvwSelFun.ListItems(key).SubItems(1) = "Constant"
    ElseIf ListIndex = 5 Then
        lvwSelFun.ListItems(key).SubItems(1) = "Collection"
    End If
    
End Sub

Private Sub AgregarMiembroFinal()

    Dim k As Integer
    Dim sKey As String
    
    util.Hourglass hwnd, True
    
    If lvwMembers.ListItems.count > 0 Then
        For k = 1 To lvwMembers.ListItems.count
            If lvwMembers.ListItems(k).Checked Then
                If tvwLibrary.SelectedItem.Tag = "GF#" And lvwMembers.ListItems(k).SubItems(1) = "Method" Then
                
                    sKey = "GF#ME#" & CStr(idx_method) & "#"
                    tvwLibrary.Nodes.Add "GF#", tvwChild, sKey, lvwMembers.ListItems(k).Text, 2, 2
                    
                    tvwLibrary.Nodes(sKey).Tag = "ME"
                    tvwLibrary.Nodes("GF#").Expanded = True
                    
                    idx_method = idx_method + 1
                
                ElseIf tvwLibrary.SelectedItem.Tag = "GC#" And lvwMembers.ListItems(k).SubItems(1) = "Constant" Then
                
                    sKey = "GC#CO#" & CStr(idx_constant) & "#"
                    tvwLibrary.Nodes.Add "GC#", tvwChild, sKey, lvwMembers.ListItems(k).Text, 6, 6
                    
                    tvwLibrary.Nodes(sKey).Tag = "CO"
                    tvwLibrary.Nodes("GC#").Expanded = True
                    idx_constant = idx_constant + 1
                    
                ElseIf tvwLibrary.SelectedItem.Tag = "GO#" And lvwMembers.ListItems(k).SubItems(1) = "Object" Then
                
                    sKey = "GO#OB#" & CStr(idx_object) & "#"
                    tvwLibrary.Nodes.Add "GO#", tvwChild, sKey, lvwMembers.ListItems(k).Text, 15, 15
                    
                    tvwLibrary.Nodes(sKey).Tag = "OB"
                    tvwLibrary.Nodes("GO#").Expanded = True
                    idx_object = idx_object + 1
                    
                ElseIf tvwLibrary.SelectedItem.Tag = "GS#" And lvwMembers.ListItems(k).SubItems(1) = "Collection" Then
                    
                    sKey = "GS#CS#" & CStr(idx_collection) & "#"
                    tvwLibrary.Nodes.Add "GS#", tvwChild, sKey, lvwMembers.ListItems(k).Text, 13, 13
                    
                    tvwLibrary.Nodes(sKey).Tag = "CS"
                    tvwLibrary.Nodes("GS#").Expanded = True
                    idx_collection = idx_collection + 1
                    
                ElseIf tvwLibrary.SelectedItem.Tag = "OB" Then
                    If lvwMembers.ListItems(k).SubItems(1) = "Method" Then
                    
                        sKey = tvwLibrary.SelectedItem.key & "SME#" & CStr(idx_method)
                        tvwLibrary.Nodes.Add tvwLibrary.SelectedItem.key, tvwChild, sKey, lvwMembers.ListItems(k).Text, 2, 2
                        tvwLibrary.Nodes(sKey).Tag = "ME"
                        idx_method = idx_method + 1
                        
                    ElseIf lvwMembers.ListItems(k).SubItems(1) = "Event" Then
                    
                        sKey = tvwLibrary.SelectedItem.key & "SEV#" & CStr(idx_event)
                        tvwLibrary.Nodes.Add tvwLibrary.SelectedItem.key, tvwChild, sKey, lvwMembers.ListItems(k).Text, 5, 5
                        tvwLibrary.Nodes(sKey).Tag = "EV"
                        idx_event = idx_event + 1
                        
                    ElseIf lvwMembers.ListItems(k).SubItems(1) = "Property" Then
                    
                        sKey = tvwLibrary.SelectedItem.key & "SPR#" & CStr(idx_property)
                        tvwLibrary.Nodes.Add tvwLibrary.SelectedItem.key, tvwChild, sKey, lvwMembers.ListItems(k).Text, 3, 3
                        tvwLibrary.Nodes(sKey).Tag = "PR"
                        idx_property = idx_property + 1
                        
                    ElseIf lvwMembers.ListItems(k).SubItems(1) = "Collection" Then
                    
                        sKey = tvwLibrary.SelectedItem.key & "SCS#" & CStr(idx_collection)
                        tvwLibrary.Nodes.Add tvwLibrary.SelectedItem.key, tvwChild, sKey, lvwMembers.ListItems(k).Text, 13, 13
                        tvwLibrary.Nodes(sKey).Tag = "CS"
                        idx_collection = idx_collection + 1
                                            
                    ElseIf lvwMembers.ListItems(k).SubItems(1) = "Constant" Then
                        
                        sKey = tvwLibrary.SelectedItem.key & "SCO#" & CStr(idx_constant)
                        tvwLibrary.Nodes.Add tvwLibrary.SelectedItem.key, tvwChild, sKey, lvwMembers.ListItems(k).Text, 6, 6
                        tvwLibrary.Nodes(sKey).Tag = "CO"
                        idx_constant = idx_constant + 1
                    
                    End If
                End If
            End If
        Next k
    End If
        
    util.Hourglass hwnd, False
    
End Sub

Private Sub cargamiembros()

    Dim k As Integer
    Dim j As Integer
    
    util.Hourglass hwnd, True
    
    ReDim arr_member_info(0)
    lvwSelFun.ListItems.Clear
        
    For k = 0 To lstFun.ListCount - 1
        If lstFun.Selected(k) Then
            ReDim Preserve arr_member_info(j)
            
            arr_member_info(j).nombre = lstFun.List(k)
            arr_member_info(j).Declaracion = arr_member_info(j).nombre
            arr_member_info(j).tipo = 2
            arr_member_info(j).ayuda = "Help for " & arr_member_info(j).nombre
            j = j + 1
        End If
    Next k
    
    For k = 0 To lstCusFun.ListCount - 1
        If lstCusFun.Selected(k) Then
            ReDim Preserve arr_member_info(j)
        
            arr_member_info(j).nombre = lstFun.List(k)
            arr_member_info(j).Declaracion = arr_member_info(j).nombre
            arr_member_info(j).tipo = 2
            arr_member_info(j).ayuda = "Help for " & arr_member_info(j).nombre
            j = j + 1
        End If
    Next k
   
    If UBound(arr_member_info) > 0 Then
        For k = 0 To UBound(arr_member_info)
            lvwSelFun.ListItems.Add , "k" & k, arr_member_info(k).nombre
            lvwSelFun.ListItems("k" & k).Tag = k
            lvwSelFun.ListItems("k" & k).SubItems(1) = "Method"
        Next k
    End If
    
    util.Hourglass hwnd, False
    
End Sub

Private Sub cargar_auto()
    
    Dim k As Integer
    Dim j As Integer
    Dim i As Integer
    Dim sKey As String
    
    If Confirma("Are you sure (All info will be lost!)") = vbNo Then Exit Sub
    
    Call CargarMiembrosLibreria
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    'indices
    idx_object = 1
    idx_property = 1
    idx_method = 1
    idx_event = 1
    idx_constant = 1
    idx_collection = 1
    
    For k = 1 To lvwMembers.ListItems.count
        If lvwMembers.ListItems(k).SubItems(1) = "Method" Then
            sKey = "GF#ME#" & CStr(idx_method) & "#"
            tvwLibrary.Nodes.Add "GF#", tvwChild, sKey, lvwMembers.ListItems(k).Text, 2, 2
            
            tvwLibrary.Nodes(sKey).Tag = "ME"
            tvwLibrary.Nodes("GF#").Expanded = True
            idx_method = idx_method + 1
            
        ElseIf lvwMembers.ListItems(k).SubItems(1) = "Constant" Then
            sKey = "GC#CO#" & CStr(idx_constant) & "#"
            tvwLibrary.Nodes.Add "GC#", tvwChild, sKey, lvwMembers.ListItems(k).Text, 6, 6
            
            tvwLibrary.Nodes(sKey).Tag = "CO"
            tvwLibrary.Nodes("GC#").Expanded = True
            idx_constant = idx_constant + 1
            
        ElseIf lvwMembers.ListItems(k).SubItems(1) = "Object" Then
            sKey = "GO#OB#" & CStr(idx_object) & "#"
            tvwLibrary.Nodes.Add "GO#", tvwChild, sKey, lvwMembers.ListItems(k).Text, 15, 15
            
            tvwLibrary.Nodes(sKey).Tag = "OB"
            tvwLibrary.Nodes("GO#").Expanded = True
            idx_object = idx_object + 1
            
        ElseIf lvwMembers.ListItems(k).SubItems(1) = "Collection" Then
            sKey = "GS#CS#" & CStr(idx_collection) & "#"
            tvwLibrary.Nodes.Add "GS#", tvwChild, sKey, lvwMembers.ListItems(k).Text, 13, 13
            
            tvwLibrary.Nodes(sKey).Tag = "CS"
            tvwLibrary.Nodes("GS#").Expanded = True
            idx_collection = idx_collection + 1
           
        End If
    Next k
    
    Screen.MousePointer = vbDefault
    DoEvents
    
    If Confirma("Reply member information for every object defined (Recommended)") = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    i = 1
    For k = 1 To tvwLibrary.Nodes.count
        If Left(tvwLibrary.Nodes(k).key, 6) = "GO#OB#" Then
            For j = 1 To lvwMembers.ListItems.count
                If lvwMembers.ListItems(j).SubItems(1) = "Method" Then
                    
                    sKey = tvwLibrary.Nodes(k).key & "SME#" & CStr(idx_method)
                    tvwLibrary.Nodes.Add tvwLibrary.Nodes(k).key, tvwChild, sKey, lvwMembers.ListItems(j).Text, 2, 2
                    tvwLibrary.Nodes(sKey).Tag = "ME"
                    idx_method = idx_method + 1
                    
                ElseIf lvwMembers.ListItems(j).SubItems(1) = "Constant" Then
                
                    sKey = tvwLibrary.Nodes(k).key & "SCO#" & CStr(idx_constant)
                    tvwLibrary.Nodes.Add tvwLibrary.Nodes(k).key, tvwChild, sKey, lvwMembers.ListItems(j).Text, 6, 6
                    tvwLibrary.Nodes(sKey).Tag = "CO"
                    idx_constant = idx_constant + 1
                    
                ElseIf lvwMembers.ListItems(j).SubItems(1) = "Collection" Then
                
                    sKey = tvwLibrary.Nodes(k).key & "SCS#" & CStr(idx_collection)
                    tvwLibrary.Nodes.Add tvwLibrary.Nodes(k).key, tvwChild, sKey, lvwMembers.ListItems(j).Text, 13, 13
                    tvwLibrary.Nodes(sKey).Tag = "CS"
                    idx_collection = idx_collection + 1
                    
                ElseIf lvwMembers.ListItems(j).SubItems(1) = "Property" Then
                
                    sKey = tvwLibrary.Nodes(k).key & "SPR#" & CStr(idx_property)
                    tvwLibrary.Nodes.Add tvwLibrary.Nodes(k).key, tvwChild, sKey, lvwMembers.ListItems(j).Text, 3, 3
                    tvwLibrary.Nodes(sKey).Tag = "PR"
                    idx_property = idx_property + 1
                    
                ElseIf lvwMembers.ListItems(j).SubItems(1) = "Event" Then
                
                    sKey = tvwLibrary.Nodes(k).key & "SEV#" & CStr(idx_event)
                    tvwLibrary.Nodes.Add tvwLibrary.Nodes(k).key, tvwChild, sKey, lvwMembers.ListItems(j).Text, 5, 5
                    tvwLibrary.Nodes(sKey).Tag = "EV"
                    idx_event = idx_event + 1
                    
                End If
            Next j
            tvwLibrary.Nodes(tvwLibrary.Nodes(k).key).Expanded = True
        End If
    Next k
    
    Screen.MousePointer = vbDefault
    DoEvents
    
    MsgBox "Information loaded!", vbInformation
    
End Sub

Private Sub cargar_libreria()

    Dim k As Integer
    Dim buffer As String
    Dim sKey As String
    Dim Metodo As String
    Dim Constante As String
    Dim objeto As String
    Dim Coleccion As String
    
    Archivo = util.StripPath(App.Path) & "libraries\" & Archivo
    
    txtLibName.Text = util.LeeIni(Archivo, "information", "name")
    txtAutor.Text = util.LeeIni(Archivo, "information", "autor")
    txtDescription.Text = util.LeeIni(Archivo, "information", "description")
    txtVersion.Text = util.LeeIni(Archivo, "information", "version")
    txtUrl.Text = util.LeeIni(Archivo, "information", "homepage")
    txtComments.Text = util.LeeIni(Archivo, "information", "comments")
    
    'cargar archivos
    For k = 1 To member_limit
        buffer = util.LeeIni(Archivo, "sourcefiles", "ele" & k)
        If Len(buffer) > 0 Then
            lstSelFiles.AddItem buffer
        Else
            Exit For
        End If
    Next k
    
    'cargar los miembros originales
    For k = 0 To member_limit
        buffer = util.LeeIni(Archivo, "members", "ele" & k + 1)
        If Len(buffer) > 0 Then
            lstFun.AddItem util.Explode(buffer, 1, "#")
            
            ReDim Preserve arr_member_info(k)
            arr_member_info(k).nombre = lstFun.List(lstFun.NewIndex)
            arr_member_info(k).Declaracion = arr_member_info(k).nombre
            arr_member_info(k).tipo = util.Explode(buffer, 2, "#")
            arr_member_info(k).ayuda = util.Explode(buffer, 3, "#")
            
            lvwSelFun.ListItems.Add , "k" & k, arr_member_info(k).nombre
            lvwSelFun.ListItems("k" & k).Tag = k
            lvwSelFun.ListItems("k" & k).SubItems(1) = get_member_description(arr_member_info(k).tipo)
        Else
            Exit For
        End If
    Next k
    
    'cargar las funciones
    chkSelAllFun.Value = 1
    
    Call CargarMiembrosLibreria
    
    'indices
    idx_object = 1
    idx_property = 1
    idx_method = 1
    idx_event = 1
    idx_constant = 1
    idx_collection = 1
    
    'cargar las funciones globales
    For k = 1 To member_limit
        Metodo = util.Explode(util.LeeIni(Archivo, "functions", "ele" & k), 1, "#")
        If Len(Metodo) = 0 Then Exit For
        
        sKey = "GF#ME#" & CStr(idx_method) & "#"
        tvwLibrary.Nodes.Add "GF#", tvwChild, sKey, Metodo, 2, 2
        tvwLibrary.Nodes(sKey).Tag = "ME"
        tvwLibrary.Nodes("GF#").Expanded = True
        idx_method = idx_method + 1
    Next k
    
    'cargar las constantes
    For k = 1 To member_limit
        Constante = util.Explode(util.LeeIni(Archivo, "constants", "ele" & k), 1, "#")
        If Len(Constante) = 0 Then Exit For
        
        sKey = "GC#CO#" & CStr(idx_constant) & "#"
        tvwLibrary.Nodes.Add "GC#", tvwChild, sKey, Constante, 6, 6
        
        tvwLibrary.Nodes(sKey).Tag = "CO"
        tvwLibrary.Nodes("GC#").Expanded = True
        idx_constant = idx_constant + 1
    Next k
        
    'cargar los objetos
    For k = 1 To member_limit
        objeto = util.Explode(util.LeeIni(Archivo, "objects", "ele" & k), 1, "#")
        If Len(objeto) = 0 Then Exit For
        
        sKey = "GO#OB#" & CStr(idx_object) & "#"
        tvwLibrary.Nodes.Add "GO#", tvwChild, sKey, objeto, 15, 15
        
        tvwLibrary.Nodes(sKey).Tag = "OB"
        tvwLibrary.Nodes("GO#").Expanded = True
        idx_object = idx_object + 1
    Next k
    
    'cargar las colecciones
    For k = 1 To member_limit
        Coleccion = util.Explode(util.LeeIni(Archivo, "collections", "ele" & k), 1, "#")
        If Len(Coleccion) = 0 Then Exit For
        
        sKey = "GS#CS#" & CStr(idx_collection) & "#"
        tvwLibrary.Nodes.Add "GS#", tvwChild, sKey, Coleccion, 13, 13
        
        tvwLibrary.Nodes(sKey).Tag = "CS"
        tvwLibrary.Nodes("GS#").Expanded = True
        idx_collection = idx_collection + 1
    Next k
    
    Dim valor As String
    
    valor = Nvl(util.LeeIni(Archivo, "information", "active"), "N")
    
    If valor = "N" Then
        chkActive.Value = 0
    ElseIf valor = "Y" Then
        chkActive.Value = 1
    End If
    
End Sub

Private Sub CargarMiembrosLibreria()

    On Error Resume Next
    
    Dim k As Integer
    
    util.Hourglass hwnd, True
    
    lvwMembers.ListItems.Clear
    tvwLibrary.Nodes.Clear
        
    Set tvwLibrary.ImageList = imgAyuda
    tvwLibrary.Nodes.Add , , "root", "Library", 12, 12
    tvwLibrary.Nodes("root").Expanded = True
    tvwLibrary.Nodes("root").Tag = "Library"
    
    tvwLibrary.Nodes.Add "root", tvwChild, "GF#", "Global Methods", 2, 2
    tvwLibrary.Nodes("GF#").Tag = "GF#"
    tvwLibrary.Nodes.Add "root", tvwChild, "GC#", "Global Constants", 6, 6
    tvwLibrary.Nodes("GC#").Tag = "GC#"
    tvwLibrary.Nodes.Add "root", tvwChild, "GO#", "Global Objects", 15, 15
    tvwLibrary.Nodes("GO#").Tag = "GO#"
    tvwLibrary.Nodes.Add "root", tvwChild, "GS#", "Global Collections", 13, 13
    tvwLibrary.Nodes("GS#").Tag = "GS#"
    
    For k = 0 To UBound(arr_member_info)
        If arr_member_info(k).tipo = 0 Then         'object
            lvwMembers.ListItems.Add , "k" & k, arr_member_info(k).nombre ', 15, 15
            lvwMembers.ListItems("k" & k).SubItems(1) = "Object"
        ElseIf arr_member_info(k).tipo = 1 Then     'property
            lvwMembers.ListItems.Add , "k" & k, arr_member_info(k).nombre '3, 3
            lvwMembers.ListItems("k" & k).SubItems(1) = "Property"
        ElseIf arr_member_info(k).tipo = 2 Then     'method
            lvwMembers.ListItems.Add , "k" & k, arr_member_info(k).nombre ', 2, 2
            lvwMembers.ListItems("k" & k).SubItems(1) = "Method"
        ElseIf arr_member_info(k).tipo = 3 Then     'event
            lvwMembers.ListItems.Add , "k" & k, arr_member_info(k).nombre ', 5, 5
            lvwMembers.ListItems("k" & k).SubItems(1) = "Event"
        ElseIf arr_member_info(k).tipo = 4 Then     'constant
            lvwMembers.ListItems.Add , "k" & k, arr_member_info(k).nombre ', 6, 6
            lvwMembers.ListItems("k" & k).SubItems(1) = "Constant"
        ElseIf arr_member_info(k).tipo = 5 Then     'collection
            lvwMembers.ListItems.Add , "k" & k, arr_member_info(k).nombre ', 13, 13
            lvwMembers.ListItems("k" & k).SubItems(1) = "Collection"
        End If
    Next k
    
    util.Hourglass hwnd, False
    
    Err = 0
    
End Sub

Private Sub eliminar_miembros(ByVal todos As Boolean)

    If todos Then
        tvwLibrary.Nodes.Clear
        Set tvwLibrary.ImageList = imgAyuda
        tvwLibrary.Nodes.Add , , "root", "Library", 12, 12
        tvwLibrary.Nodes("root").Expanded = True
        tvwLibrary.Nodes("root").Tag = "Library"
        
        tvwLibrary.Nodes.Add "root", tvwChild, "GF#", "Global Methods", 2, 2
        tvwLibrary.Nodes("GF#").Tag = "GF#"
        tvwLibrary.Nodes.Add "root", tvwChild, "GC#", "Global Constants", 6, 6
        tvwLibrary.Nodes("GC#").Tag = "GC#"
        tvwLibrary.Nodes.Add "root", tvwChild, "GO#", "Global Objects", 15, 15
        tvwLibrary.Nodes("GO#").Tag = "GO#"
        tvwLibrary.Nodes.Add "root", tvwChild, "GS#", "Global Collections", 13, 13
        tvwLibrary.Nodes("GS#").Tag = "GS#"
    Else
        If Not tvwLibrary.SelectedItem Is Nothing Then
            If tvwLibrary.SelectedItem.Tag = "GF#" Then
            
            ElseIf tvwLibrary.SelectedItem.Tag = "GC#" Then
            
            ElseIf tvwLibrary.SelectedItem.Tag = "GO#" Then
            
            ElseIf tvwLibrary.SelectedItem.Tag = "GS#" Then
            
            ElseIf tvwLibrary.SelectedItem.Tag = "Library" Then
            
            Else
                tvwLibrary.Nodes.Remove tvwLibrary.SelectedItem.key
            End If
        End If
    End If
    
End Sub

Private Function GeneraLibreria() As Boolean

    On Error GoTo ErrorGeneraLibreria
    
    Dim ret As Boolean
    Dim k As Integer
    Dim j As Integer
    Dim M As Integer
    
    Dim arr_objetos() As String
    Dim sKey As String
    Dim Source As String
    
    ReDim arr_objetos(0)
    
    If Len(Archivo) = 0 Then
        Archivo = InputBox("Name (No Special Characters) :", "Create/Edit Library")
        
        Archivo = util.StripPath(App.Path) & "libraries\" & Archivo & ".lib"
        
        If ArchivoExiste2(Archivo) Then
            If Confirma("Library already exists. Replace ?") = vbNo Then
                Archivo = vbNullString
                GeneraLibreria = False
                Exit Function
            End If
        End If
    End If
    
    If Len(Archivo) > 0 Then
        util.Hourglass hwnd, True
        
        'informacion de la libreria
        util.BorrarArchivo Archivo
        
        util.GrabaIni Archivo, "information", "name", txtLibName.Text
        util.GrabaIni Archivo, "information", "autor", txtAutor.Text
        util.GrabaIni Archivo, "information", "description", txtDescription.Text
        util.GrabaIni Archivo, "information", "version", txtVersion.Text
        util.GrabaIni Archivo, "information", "homepage", txtUrl.Text
        util.GrabaIni Archivo, "information", "comments", txtComments.Text
            
        'grabar los archivos desde donde vino la informacion
        j = 1
        For k = 0 To lstSelFiles.ListCount - 1
            util.GrabaIni Archivo, "sourcefiles", "ele" & j, lstSelFiles.List(k)
            j = j + 1
        Next k
            
        'grabar la informacion base de los miembros finales que se registraron en la libreria
        For k = 1 To lvwMembers.ListItems.count
            util.GrabaIni Archivo, "members", "ele" & k, get_member_information(lvwMembers.ListItems(k).Text)
        Next k
        
        'grabar los objetos base de la libreria
        j = 1
        For k = 1 To tvwLibrary.Nodes.count
            If tvwLibrary.Nodes(k).Tag = "OB" Then
                util.GrabaIni Archivo, "objects", "ele" & j, tvwLibrary.Nodes(k).Text
                ReDim Preserve arr_objetos(j)
                arr_objetos(j) = tvwLibrary.Nodes(k).Text & "|" & tvwLibrary.Nodes(k).key
                j = j + 1
            End If
        Next k
            
        'grabar los metodos globales que no pertenecen a algun objeto
        j = 1
        For k = 1 To tvwLibrary.Nodes.count
            If tvwLibrary.Nodes(k).Tag = "ME" And Left$(tvwLibrary.Nodes(k).key, 6) = "GF#ME#" Then
                util.GrabaIni Archivo, "functions", "ele" & j, get_member_information(tvwLibrary.Nodes(k).Text)
                j = j + 1
            End If
        Next k
            
        'grabar las constantes globales
        j = 1
        For k = 1 To tvwLibrary.Nodes.count
            If tvwLibrary.Nodes(k).Tag = "CO" And Left$(tvwLibrary.Nodes(k).key, 6) = "GC#CO#" Then
                util.GrabaIni Archivo, "constants", "ele" & j, get_member_information(tvwLibrary.Nodes(k).Text)
                j = j + 1
            End If
        Next k
        
        'grabar las colecciones globales
        j = 1
        For k = 1 To tvwLibrary.Nodes.count
            If tvwLibrary.Nodes(k).Tag = "CS" And Left$(tvwLibrary.Nodes(k).key, 6) = "GS#CS#" Then
                util.GrabaIni Archivo, "collections", "ele" & j, get_member_information(tvwLibrary.Nodes(k).Text)
                j = j + 1
            End If
        Next k
            
        'por cada objeto cargar los miembros relacionados
        For k = 1 To UBound(arr_objetos)
        
            Source = util.Explode(arr_objetos(k), 1, "|")
            sKey = util.Explode(arr_objetos(k), 2, "|")
            M = 1
            For j = 1 To tvwLibrary.Nodes.count
                If Len(tvwLibrary.Nodes(j).key) > 0 Then
                    If Len(tvwLibrary.Nodes(j).key) >= Len(sKey) Then
                        If Left$(tvwLibrary.Nodes(j).key, Len(sKey)) = sKey Then
                            If TagValido(tvwLibrary.Nodes(j).Tag) Then
                                util.GrabaIni Archivo, Source, "ele" & M, get_member_information(tvwLibrary.Nodes(j).Text)
                                M = M + 1
                            End If
                        End If
                    End If
                End If
            Next j
        Next k
        
        If chkActive.Value = 0 Then
            util.GrabaIni Archivo, "information", "active", "N"
        Else
            util.GrabaIni Archivo, "information", "active", "Y"
        End If
        
        util.Hourglass hwnd, False
        
        ret = True
    End If
    
    GeneraLibreria = ret
    
    Exit Function
    
ErrorGeneraLibreria:
    ret = False
    MsgBox "GeneraLibreria : " & Err & " " & Error$, vbCritical
    GeneraLibreria = False
    
End Function

Private Function get_member_description(ByVal Indice As Integer) As String

    Dim ret As String
    
    If Indice = 0 Then
        ret = "Object"
    ElseIf Indice = 1 Then
        ret = "Property"
    ElseIf Indice = 2 Then
        ret = "Method"
    ElseIf Indice = 3 Then
        ret = "Event"
    ElseIf Indice = 4 Then
        ret = "Constant"
    ElseIf Indice = 5 Then
        ret = "Collection"
    End If
    
    get_member_description = ret
    
End Function

Private Function get_member_information(ByVal texto As String) As String

    Dim k As Integer
    Dim ret As String
    
    For k = 0 To UBound(arr_member_info)
        If arr_member_info(k).nombre = texto Then
            ret = arr_member_info(k).nombre & "#" & arr_member_info(k).tipo & "#" & arr_member_info(k).ayuda
            Exit For
        End If
    Next k
    
    get_member_information = ret
    
End Function


Private Sub LoadFunctions()

    On Error GoTo ErrorLoadCodeFunctions
        
    Dim funcion As New CFuncion
    Dim k As Integer
    Dim j As Integer
    
    util.Hourglass hwnd, True
    
    lstFun.Clear
    
    For j = 0 To lstSelFiles.ListCount - 1
        cExpFun.Clear
        cExpFun.filename = lstSelFiles.List(j)
        If Not cExpFun.Explore Then
            Exit Sub
        End If

        With cExpFun
            For k = 1 To .Funciones.count
                Set funcion = New CFuncion
                Set funcion = .Funciones.ITem(k)
                lstFun.AddItem Replace(funcion.FullName, " ", "")
                Set funcion = Nothing
            Next k
        End With
    Next j
    
    chkSelAllFun.Value = 1
    
    util.Hourglass hwnd, False
            
    Exit Sub
    
ErrorLoadCodeFunctions:
    MsgBox "LoadFunctions : " & Err & " " & Error$, vbCritical
    Exit Sub
    
End Sub

Private Sub agregar_archivo()

    Dim Archivo As String
    
    If Cdlg.VBGetOpenFileName(Archivo, , , , , , strGlosa(), , LastPath, "Select a file ...", , hwnd) Then
        lstSelFiles.AddItem Archivo
        LastPath = util.PathArchivo(Archivo)
    End If
    
End Sub

Private Sub agregar_carpeta()

    Dim Path As String
    Dim arr_files() As String
    Dim k As Integer
    
    Path = util.BrowseFolder(hwnd)
    
    If Len(Path) > 0 Then
        util.Hourglass hwnd, True
        get_files_from_folder Path, arr_files()
        For k = 1 To UBound(arr_files)
            lstSelFiles.AddItem arr_files(k)
        Next k
        util.Hourglass hwnd, False
    End If
    
End Sub

Private Sub multi_assign()

    Dim k As Integer
    Dim j As Integer
    Dim fload As Boolean
    
    If lvwSelFun.ListItems.count > 0 Then
        If Not lvwSelFun.ListItems Is Nothing Then
            j = 1
            For k = 1 To lvwSelFun.ListItems.count
                If lvwSelFun.ListItems(k).Checked Then
                    If Not fload Then
                        Load frmMultiAssign
                        fload = True
                    End If
                    With frmMultiAssign.lvwSelFun
                        .ListItems.Add , "k" & j, lvwSelFun.ListItems(k).Text
                        .ListItems("k" & j).SubItems(1) = lvwSelFun.ListItems(k).SubItems(1)
                        .ListItems("k" & j).Tag = lvwSelFun.ListItems(k).Tag & "|" & lvwSelFun.ListItems(k).key
                        j = j + 1
                    End With
                End If
            Next k
            
            If fload Then
                frmMultiAssign.Show vbModal
            End If
        End If
    End If
    
End Sub

Private Function TagValido(ByVal Tag As String) As Boolean

    Dim ret As Boolean
    
    Select Case Tag
        Case "ME", "EV", "PR", "CS", "CO"
            ret = True
    End Select
    
    TagValido = ret
    
End Function

Private Function ValidaIngreso() As Boolean

    Dim ret As Boolean
    Dim k As Integer
    
    'step 1
    If lstSelFiles.ListCount = 0 Then
        MsgBox "You must select source files", vbCritical
        cmd(7).Enabled = True
        Indice = 0
        cmd(6).Enabled = True
        fra(0).ZOrder 0
        Exit Function
    End If
    
    'step 2
    If lstFun.ListCount = 0 Then
        MsgBox "You must select functions to import", vbCritical
        cmd(7).Enabled = True
        Indice = 1
        cmd(6).Enabled = True
        fra(1).ZOrder 0
        Exit Function
    End If
    
    'step 3
    If lvwSelFun.ListItems.count = 0 Then
        MsgBox "You must select functions to import", vbCritical
        cmd(7).Enabled = True
        Indice = 2
        cmd(6).Enabled = True
        fra(2).ZOrder 0
        Exit Function
    End If
    
    'step 4
    If tvwLibrary.Nodes.count = 0 Then
        MsgBox "You must select members (Final Design Library)", vbCritical
        cmd(7).Enabled = True
        Indice = 3
        cmd(6).Enabled = True
        fra(3).ZOrder 0
        Exit Function
    End If
    
    'verificar que al menos uno de los nodos principales tenga al menos un nodo es su interior ...
    For k = 1 To tvwLibrary.Nodes.count
        If tvwLibrary.Nodes(k).key = "root" Then
        ElseIf tvwLibrary.Nodes(k).key = "GF" Then
        ElseIf tvwLibrary.Nodes(k).key = "GC" Then
        ElseIf tvwLibrary.Nodes(k).key = "GO" Then
        ElseIf tvwLibrary.Nodes(k).key = "GS" Then
        Else
            ret = True
            Exit For
        End If
    Next k
    
    If Not ret Then
        MsgBox "You must design the library first.", vbCritical
        cmd(7).Enabled = True
        Indice = 3
        cmd(6).Enabled = True
        fra(3).ZOrder 0
        Exit Function
    End If
    
    If txtLibName.Text = "" Then
        MsgBox "Please input the library name", vbCritical
        cmd(7).Enabled = True
        Indice = 4
        cmd(6).Enabled = False
        fra(4).ZOrder 0
        txtLibName.SetFocus
        Exit Function
    End If
    
    ValidaIngreso = True
    
End Function

Private Sub chkMembers_Click()

    Dim k As Integer
    Dim ret As Boolean
    
    util.Hourglass hwnd, True
    
    If chkMembers.Value = 1 Then ret = True
    
    For k = 1 To lvwMembers.ListItems.count
        lvwMembers.ListItems(k).Checked = ret
    Next k
    
    util.Hourglass hwnd, False
    
End Sub

Private Sub chkSelAllFun_Click()
    
    Dim k As Integer
    Dim ret As Boolean
    
    If chkSelAllFun.Value = 1 Then ret = True
    
    For k = 0 To lstFun.ListCount - 1
        lstFun.Selected(k) = ret
    Next k
    
End Sub

Private Sub chkSelAllMem_Click()

    Dim ret As Boolean
    Dim k As Integer
    
    If chkSelAllMem.Value Then ret = True
    
    util.Hourglass hwnd, True
    
    For k = 1 To lvwSelFun.ListItems.count
        lvwSelFun.ListItems(k).Checked = ret
    Next k
    
    util.Hourglass hwnd, False
End Sub

Private Sub cmd_Click(Index As Integer)

    Dim k As Integer
    
    Select Case Index
        Case 22 'multi-assign
            Call multi_assign
        Case 21 'auto. Intentar cargar en forma optima segun tipos
            Call cargar_auto
        Case 20 'recargar miembros finales
            Call CargarMiembrosLibreria
        Case 5  'agregar miembro final
            Call AgregarMiembroFinal
        Case 10 'eliminar miembro final
            eliminar_miembros False
        Case 11 'agregar todos los miembros finales
            Call cargar_auto
        Case 12 'eliminar todos los miembros finales
            eliminar_miembros True
        Case 18 'actualizar informacion de miembro
            If lvwSelFun.ListItems.count > 0 Then
                If Not lvwSelFun.SelectedItem Is Nothing Then
                    arr_member_info(lvwSelFun.SelectedItem.Tag).Declaracion = txtDeclaration.Text
                    arr_member_info(lvwSelFun.SelectedItem.Tag).tipo = cboType.ListIndex
                    arr_member_info(lvwSelFun.SelectedItem.Tag).ayuda = txtHelp.Text
                    
                    If cboType.ListIndex = 0 Then
                        lvwSelFun.ListItems(lvwSelFun.SelectedItem.key).SubItems(1) = "Object"
                    ElseIf cboType.ListIndex = 1 Then
                        lvwSelFun.ListItems(lvwSelFun.SelectedItem.key).SubItems(1) = "Property"
                    ElseIf cboType.ListIndex = 2 Then
                        lvwSelFun.ListItems(lvwSelFun.SelectedItem.key).SubItems(1) = "Method"
                    ElseIf cboType.ListIndex = 3 Then
                        lvwSelFun.ListItems(lvwSelFun.SelectedItem.key).SubItems(1) = "Event"
                    ElseIf cboType.ListIndex = 4 Then
                        lvwSelFun.ListItems(lvwSelFun.SelectedItem.key).SubItems(1) = "Constant"
                    ElseIf cboType.ListIndex = 5 Then
                        lvwSelFun.ListItems(lvwSelFun.SelectedItem.key).SubItems(1) = "Collection"
                    End If
                End If
            End If
        Case 19 'recargar miembros
            If lvwSelFun.ListItems.count > 0 Then
                If Confirma("Are you sure to reload members. All information will be lost!") = vbYes Then
                    Call cargamiembros
                End If
            End If
        Case 15 'eliminar funcion no customizada
            If lstFun.ListCount > -1 Then
                For k = lstFun.ListCount - 1 To 0 Step -1
                    If lstFun.Selected(k) Then
                        lstFun.RemoveItem k
                    End If
                Next k
            End If
        Case 16 'eliminar todas las funciones no customizadas
            For k = lstFun.ListCount - 1 To 0 Step -1
                lstFun.RemoveItem k
            Next k
        Case 17 'recargar las funciones
            Call LoadFunctions
        Case 0  'agregar archivo
            Call agregar_archivo
        Case 8  'agregar carpeta
            Call agregar_carpeta
        Case 9  'eliminar archivo
            If lstSelFiles.ListCount > -1 Then
                For k = lstSelFiles.ListCount - 1 To 0 Step -1
                    If lstSelFiles.Selected(k) Then
                        lstSelFiles.RemoveItem k
                    End If
                Next k
            End If
        Case 13 'eliminar todos
            If lstSelFiles.ListCount > -1 Then
                For k = lstSelFiles.ListCount - 1 To 0 Step -1
                    lstSelFiles.RemoveItem k
                Next k
            End If
        Case 1  'generar la libreria ....
            If ValidaIngreso() Then
                If GeneraLibreria() Then
                    MsgBox "The library was generated successful.", vbInformation
                    frmLibraryManager.cargar_librerias
                    Unload Me
                End If
            End If
        Case 2
            Unload Me
        Case 3  'agregar miembro customizado
            Dim nombre As String
            nombre = InputBox("Name", "Add Custom Member")
            If Len(nombre) > 0 Then
                lstCusFun.AddItem nombre
            End If
        Case 4  'eliminar miembro customizado
            If lstCusFun.ListCount > -1 Then
                For k = lstCusFun.ListCount - 1 To 0 Step -1
                    If lstCusFun.Selected(k) Then
                        lstCusFun.RemoveItem k
                    End If
                Next k
            End If
        Case 14 'eliminar todos los miembros customizados
            For k = lstCusFun.ListCount - 1 To 0 Step -1
                lstCusFun.RemoveItem k
            Next k
        Case 6  'next
            Indice = Indice + 1
            cmd(7).Enabled = True
            If Indice < 4 Then
                If Indice = 1 Then      'step 2
                    If lstFun.ListCount = 0 Then
                        Call LoadFunctions
                    End If
                ElseIf Indice = 2 Then  'step 3
                    If lvwSelFun.ListItems.count = 0 Then
                        Call cargamiembros
                    End If
                ElseIf Indice = 3 Then  'step 4
                    If lvwSelFun.ListItems.count > 0 Then
                        If lvwMembers.ListItems.count = 0 Then
                            Call CargarMiembrosLibreria
                        End If
                    End If
                End If
                cmd(6).Enabled = True
            Else
                Indice = 4
                cmd(6).Enabled = False
            End If
            fra(Indice).ZOrder 0
        Case 7  'previo
            Indice = Indice - 1
            cmd(6).Enabled = True
            If Indice < 0 Then
                cmd(7).Enabled = False
                Indice = 0
            Else
                cmd(6).Enabled = True
            End If
            fra(Indice).ZOrder 0
    End Select
        
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
        
    ReDim arr_member_info(0)
    
    'indices
    idx_object = 1
    idx_property = 1
    idx_method = 1
    idx_event = 1
    idx_constant = 1
    idx_collection = 1
    
    cboType.AddItem "Object"
    cboType.AddItem "Property"
    cboType.AddItem "Method"
    cboType.AddItem "Event"
    cboType.AddItem "Constant"
    cboType.AddItem "Collection"
    
    fra(1).Move fra(0).Left, fra(0).Top, fra(0).Width, fra(0).Height
    fra(2).Move fra(0).Left, fra(0).Top, fra(0).Width, fra(0).Height
    fra(3).Move fra(0).Left, fra(0).Top, fra(0).Width, fra(0).Height
    fra(4).Move fra(0).Left, fra(0).Top, fra(0).Width, fra(0).Height
    
    fra(0).ZOrder 0
    
    If Len(Archivo) > 0 Then
        Call cargar_libreria
    End If
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmImportLibrary = Nothing
End Sub


Private Sub lvwSelFun_ItemClick(ByVal ITem As MSComctlLib.ListItem)

    txtDeclaration.Text = arr_member_info(ITem.Tag).Declaracion
    cboType.ListIndex = arr_member_info(ITem.Tag).tipo
    txtHelp.Text = arr_member_info(ITem.Tag).ayuda
    
End Sub


