VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{971CBA62-C40B-4E49-9602-35CDA8C00036}#1.0#0"; "vbsExpBar6.ocx"
Object = "{2128BF45-F895-4206-84CD-F4DE2DD8D6B1}#2.0#0"; "vbsTbar6.ocx"
Object = "{E861E505-03C0-49EC-8FC6-8AB54B4361FE}#2.0#0"; "vbsDTab6.ocx"
Object = "{98F993CC-3598-405A-9E9A-0D2CF198B250}#2.0#0"; "vbsDkTb6.ocx"
Object = "{A9700EB9-4073-41EA-AF2D-6410341636CD}#1.0#0"; "vbsmditabs.ocx"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "JavaScript Plus!"
   ClientHeight    =   8505
   ClientLeft      =   3345
   ClientTop       =   3030
   ClientWidth     =   8685
   Icon            =   "mfrmMain.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSizeRight 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7740
      Left            =   5610
      MousePointer    =   9  'Size W E
      ScaleHeight     =   7740
      ScaleWidth      =   45
      TabIndex        =   23
      Top             =   405
      Visible         =   0   'False
      Width           =   40
   End
   Begin VB.PictureBox picSizeLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7740
      Left            =   3000
      MousePointer    =   9  'Size W E
      ScaleHeight     =   7740
      ScaleWidth      =   45
      TabIndex        =   22
      Top             =   405
      Visible         =   0   'False
      Width           =   40
   End
   Begin RevMDITabs.RevMDITabsCtl mdiforms 
      Left            =   3555
      Top             =   3690
      _ExtentX        =   847
      _ExtentY        =   847
      Style           =   1
   End
   Begin vbalDkTb6.vbalDockContainer vbalDockContainer4 
      Align           =   4  'Align Right
      Height          =   7740
      Left            =   8655
      TabIndex        =   3
      Top             =   405
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   13653
      AllowUndock     =   0   'False
   End
   Begin vbalDkTb6.vbalDockContainer vbalDockContainer3 
      Align           =   3  'Align Left
      Height          =   7740
      Left            =   3045
      TabIndex        =   2
      Top             =   405
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   13653
      AllowUndock     =   0   'False
   End
   Begin vbalDkTb6.vbalDockContainer vbalDockContainer1 
      Align           =   1  'Align Top
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   53
      AllowUndock     =   0   'False
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
      ScaleWidth      =   8625
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8685
      Begin vbalTBar6.cToolbar tbrFile 
         Height          =   255
         Left            =   1290
         Top             =   15
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
      End
      Begin vbalTBar6.cToolbar tbrMenu 
         Height          =   210
         Left            =   150
         Top             =   45
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   370
      End
      Begin vbalTBar6.cToolbar tbrEdit 
         Height          =   255
         Left            =   1845
         Top             =   45
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
      End
      Begin vbalTBar6.cToolbar tbrFormat 
         Height          =   270
         Left            =   2400
         Top             =   15
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   476
      End
      Begin vbalTBar6.cToolbar tbrForms 
         Height          =   270
         Left            =   3045
         ToolTipText     =   "Forms Toolbar"
         Top             =   30
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   476
         DrawStyle       =   2
      End
      Begin vbalTBar6.cToolbar tbrPlus 
         Height          =   270
         Left            =   3720
         Top             =   45
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   476
      End
      Begin vbalTBar6.cToolbar tbrJs 
         Height          =   270
         Left            =   4350
         Top             =   45
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   476
      End
      Begin vbalTBar6.cToolbar tbrHtm 
         Height          =   270
         Left            =   5025
         Top             =   45
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   476
      End
      Begin vbalTBar6.cToolbar tbrTools 
         Height          =   270
         Left            =   5685
         Top             =   30
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   476
      End
   End
   Begin vbalDTab6.vbalDTabControl tabLeft 
      Align           =   3  'Align Left
      Height          =   7740
      Left            =   0
      TabIndex        =   4
      Top             =   405
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   13653
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Pinnable        =   -1  'True
      Begin VB.PictureBox picTab 
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         Height          =   1050
         Index           =   6
         Left            =   105
         ScaleHeight     =   1050
         ScaleWidth      =   2370
         TabIndex        =   28
         Top             =   6240
         Visible         =   0   'False
         Width           =   2370
         Begin VB.PictureBox picTbFiles 
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
            TabIndex        =   32
            Top             =   0
            Width           =   4800
            Begin VB.CommandButton cmdFiles 
               Caption         =   "Clear Unused"
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   40
               Width           =   1215
            End
            Begin VB.Label lblTot 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Label1"
               Height          =   195
               Left            =   2040
               TabIndex        =   34
               Top             =   45
               Width           =   480
            End
         End
         Begin MSComctlLib.ListView lvwOpeFiles 
            Height          =   375
            Left            =   240
            TabIndex        =   29
            ToolTipText     =   "Clic to open or activate selected file"
            Top             =   600
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   661
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "FileName"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Path"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "FTP Site"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Remote Path"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox picTab 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   870
         Index           =   3
         Left            =   285
         ScaleHeight     =   870
         ScaleWidth      =   2385
         TabIndex        =   26
         Top             =   4920
         Visible         =   0   'False
         Width           =   2385
         Begin jsplus.vbsXHTML vbsXHTML1 
            Height          =   570
            Left            =   330
            TabIndex        =   27
            Top             =   180
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   1005
         End
      End
      Begin VB.PictureBox picTab 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   870
         Index           =   5
         Left            =   195
         ScaleHeight     =   870
         ScaleWidth      =   2385
         TabIndex        =   20
         Top             =   4050
         Visible         =   0   'False
         Width           =   2385
         Begin jsplus.vbsCSS vbsCSS1 
            Height          =   480
            Left            =   285
            TabIndex        =   21
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   847
         End
      End
      Begin VB.PictureBox picTab 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   870
         Index           =   4
         Left            =   180
         ScaleHeight     =   870
         ScaleWidth      =   2385
         TabIndex        =   18
         Top             =   3120
         Visible         =   0   'False
         Width           =   2385
         Begin jsplus.vbSDhtml vbSDhtml1 
            Height          =   555
            Left            =   300
            TabIndex        =   19
            Top             =   180
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   979
         End
      End
      Begin VB.PictureBox picTab 
         BackColor       =   &H00404080&
         BorderStyle     =   0  'None
         Height          =   870
         Index           =   2
         Left            =   75
         ScaleHeight     =   870
         ScaleWidth      =   2385
         TabIndex        =   7
         Top             =   2280
         Visible         =   0   'False
         Width           =   2385
         Begin jsplus.vbSMarkup MarkHlp 
            Height          =   405
            Left            =   195
            TabIndex        =   12
            Top             =   210
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   714
         End
      End
      Begin VB.PictureBox picTab 
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         Height          =   930
         Index           =   1
         Left            =   30
         ScaleHeight     =   930
         ScaleWidth      =   2370
         TabIndex        =   6
         Top             =   1185
         Visible         =   0   'False
         Width           =   2370
         Begin jsplus.vbSJava jsHlp 
            Height          =   1605
            Left            =   210
            TabIndex        =   13
            Top             =   90
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   2831
         End
      End
      Begin VB.PictureBox picTab 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   0
         Left            =   120
         ScaleHeight     =   720
         ScaleWidth      =   1905
         TabIndex        =   5
         Top             =   405
         Visible         =   0   'False
         Width           =   1905
         Begin jsplus.vbsFileExp filExp 
            Height          =   1665
            Left            =   165
            TabIndex        =   14
            Top             =   150
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   2937
         End
      End
   End
   Begin vbalDkTb6.vbalDockContainer vbalDockContainer2 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   8145
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   635
      NonDockingArea  =   -1  'True
      AllowUndock     =   0   'False
      Begin MSComctlLib.StatusBar picStatus 
         Height          =   360
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   13380
         _ExtentX        =   23601
         _ExtentY        =   635
         SimpleText      =   "Editor de Javascript"
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   7
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   14111
               MinWidth        =   14111
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   1411
               MinWidth        =   1411
               Text            =   "Line"
               TextSave        =   "Line"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   882
               MinWidth        =   882
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   1235
               MinWidth        =   1235
               Text            =   "Column"
               TextSave        =   "Column"
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   882
               MinWidth        =   882
            EndProperty
            BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   5292
               MinWidth        =   5292
            EndProperty
            BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   3528
               MinWidth        =   3528
            EndProperty
         EndProperty
      End
   End
   Begin vbalDTab6.vbalDTabControl tabRight 
      Align           =   4  'Align Right
      Height          =   7740
      Left            =   5655
      TabIndex        =   9
      Top             =   405
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   13653
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Pinnable        =   -1  'True
      Begin VB.PictureBox picRight 
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Height          =   795
         Index           =   4
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   2340
         TabIndex        =   30
         Top             =   6000
         Visible         =   0   'False
         Width           =   2340
         Begin jsplus.CodeLibrary CodeLibrary1 
            Height          =   375
            Left            =   720
            TabIndex        =   31
            Top             =   360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
         End
      End
      Begin VB.PictureBox picRight 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   0  'None
         Height          =   540
         Index           =   2
         Left            =   120
         ScaleHeight     =   540
         ScaleWidth      =   2340
         TabIndex        =   24
         Top             =   3105
         Visible         =   0   'False
         Width           =   2340
         Begin jsplus.vbsClipboard tboClp 
            Height          =   420
            Left            =   675
            TabIndex        =   25
            Top             =   135
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   741
         End
      End
      Begin VB.PictureBox picRight 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   540
         Index           =   1
         Left            =   90
         ScaleHeight     =   540
         ScaleWidth      =   2340
         TabIndex        =   15
         Top             =   2145
         Visible         =   0   'False
         Width           =   2340
         Begin jsplus.ColPicker ColPicker1 
            Height          =   345
            Left            =   585
            TabIndex        =   16
            Top             =   135
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   609
         End
      End
      Begin VB.PictureBox picRight 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   540
         Index           =   0
         Left            =   105
         ScaleHeight     =   540
         ScaleWidth      =   2340
         TabIndex        =   10
         Top             =   1230
         Visible         =   0   'False
         Width           =   2340
         Begin vbalExplorerBarLib6.vbalExplorerBarCtl HlpExp 
            Height          =   360
            Left            =   555
            TabIndex        =   11
            Top             =   315
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   635
            BackColorEnd    =   0
            BackColorStart  =   0
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents m_cMenu As cPopupMenu
Attribute m_cMenu.VB_VarHelpID = -1
Private WithEvents m_cMenuPop As cPopupMenu
Attribute m_cMenuPop.VB_VarHelpID = -1
Public m_MainImg As cVBALImageList
Public m_cMRU As New cMRUFileList

Public fLoading As Boolean
Public SaveMru As Boolean
Public fexpired As Boolean
Public paulina As Boolean
Public palic As Integer

Private Const c_file = 296
Private tabrightsel As Integer
Private floading_prop As Boolean
Private fdisplaywelcome As Boolean
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Private attrib2 As String

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Sub action_menu(mMenu As cPopupMenu, ByVal ItemNumber As Integer)
    
    Dim Archivo As String
    Dim str As New cStringBuilder
    Dim CodeLib As New cCodeLibrary
    
    Select Case mMenu.ItemKey(ItemNumber)
        Case "mnuPdf1"
            ShowPdfHelp (1)
        Case "mnuPdf2"
            ShowPdfHelp (2)
        Case "mnuPdf3"
            ShowPdfHelp (3)
        Case "mnuPdf4"
            ShowPdfHelp (4)
        Case "mnuJavascript(28)"
            frmOfuscator.Show vbModal
        Case "mnuJavascript(27)"
            frmHelp.url = StripPath(App.Path) & "howto\howtojs.htm"
            Load frmHelp
            frmHelp.Show
        Case "mnuJavascript(24)"    'tutorial
            frmHelp.url = StripPath(App.Path) & "tutorial\tutorial.htm"
            Load frmHelp
            frmHelp.Show
        Case "mnuAddIn(5)"
         frmNewPlug.Show vbModal
        Case "mnuWindow_Top(1)"  'top window
           mMenu.Checked(ItemNumber) = Not mMenu.Checked(ItemNumber)
           If mMenu.Checked(ItemNumber) Then
              windowontop hwnd
           Else
              windownontop hwnd
           End If
        Case "mnuWindow_Top(2)"  'cascade
           frmMain.Arrange vbCascade
        Case "mnuWindow_Top(3)"  'horizontal
           frmMain.Arrange vbTileHorizontal
        Case "mnuWindow_Top(4)"  'vertical
           frmMain.Arrange vbTileVertical
        Case "mnuWindow_Top(5)"  'icons
           frmMain.Arrange vbArrangeIcons
        Case "mnuLibrary(1)"    'category
            CodeLib.DataPath = util.StripPath(App.Path)
            CodeLib.CategoryWizard
        Case "mnuLibrary(2)"    'language
            CodeLib.DataPath = util.StripPath(App.Path)
            CodeLib.LanguageWizard
        Case "mnuLibrary(3)"    'browse
            CodeLib.DataPath = util.StripPath(App.Path)
            CodeLib.LanguagePath = StripPath(App.Path) & "languages\"
            CodeLib.BrowseLibrary
        Case "mnuLibrary(5)"    'save
            If Not ActiveForm Is Nothing Then
                If Trim$(ActiveForm.txtCode.Text) <> "" Then
                    Dim nFreeFile As Long
                    
                    CodeLib.DataPath = util.StripPath(App.Path)
                    
                    On Error Resume Next
                    nFreeFile = FreeFile
                    Archivo = util.ArchivoTemporal()
                    
                    Open Archivo For Output As #nFreeFile
                        Print #nFreeFile, ActiveForm.txtCode.Text
                    Close #nFreeFile
                    
                    If Err = 0 Then
                        CodeLib.SaveCode (Archivo)
                    Else
                        MsgBox "Failed to save temp file : " & Err & " " & Err.description, vbCritical
                    End If
                Else
                    MsgBox "Nothing to save", vbCritical
                End If
            End If
        Case "mnuConfigBrowsers(1)"
            frmConfBrow.Show vbModal
        Case "mnuFormat(20)"
            If Not ActiveForm Is Nothing Then
                frmFmtSpecial.Show vbModal
            End If
        Case "mnInsert(50)"
            frmAnsiExp.inifile = util.StripPath(App.Path) & "config\ansihelp.ini"
            frmAnsiExp.Show vbModal
        Case "mnuItemHelp(400)" 'insertar texto al inicio
            If Not frmMain.ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.InsertarInicioFin(1)
            End If
        Case "mnuItemHelp(401)" 'insertar texto al final
            If Not frmMain.ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.InsertarInicioFin(2)
            End If
        Case "mnuItemHelp(315)" 'asp
            If Not frmMain.ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.InsertaCodigoHtml("Response.Write(" & Chr$(34), Chr$(34) & ")", False)
            End If
        Case "mnuItemHelp(316)" 'php
            If Not frmMain.ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.InsertaCodigoHtml("echo " & Chr$(34), Chr$(34) & ";", False)
            End If
        Case "mnuItemHelp(317)" 'jsp
            If Not frmMain.ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.InsertaCodigoHtml("out.print (" & Chr$(34), Chr$(34) & ");", False)
            End If
        Case "mnuItemHelp(314)" 'encode url
            If Not frmMain.ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.EncodeUrl
            End If
        Case "mnuColorNames(0)" 'color names
            frmColorNames.Show vbModal
        Case "mnuTools(14)"      'icon extractor
            ExecuteTool ("IconExtractor.exe")
        Case "mnuTools(100)"     'bitmap extractor
            ExecuteTool ("BitmapExtractor.exe")
        Case "mnuTools(140)"     'icon editor
            ExecuteTool ("IconEditor.exe")
        Case "mnuItemHelp(200)" 'texto a parrafo html
            If Not frmMain.ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.InsertaCodigoHtml("<P>", "</P>", False)
            End If
        Case "mnuItemHelp(201)" 'texto a italica html
            If Not frmMain.ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.InsertaCodigoHtml("<I>", "</I>", False)
            End If
        Case "mnuItemHelp(202)" 'texto a bold html
            If Not frmMain.ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.InsertaCodigoHtml("<B>", "</B>", False)
            End If
        Case "mnuItemHelp(203)" 'texto a underline html
            If Not frmMain.ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.InsertaCodigoHtml("<U>", "</U>", False)
            End If
        Case "mnuItemHelp(204)" 'texto a commentario html
            If Not frmMain.ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.InsertaCodigoHtml("<!--", "-->", False)
            End If
        Case "mnuItemHelp(205)" 'html entities
            If Not frmMain.ActiveForm Is Nothing Then
                If Len(frmMain.ActiveForm.txtCode.SelText) > 0 Then
                    frmMain.ActiveForm.txtCode.SelText = frmMain.ActiveForm.ConvertHTMLEntityToCharacter(frmMain.ActiveForm.txtCode.SelText)
                Else
                    frmMain.ActiveForm.txtCode.Text = frmMain.ActiveForm.ConvertHTMLEntityToCharacter(frmMain.ActiveForm.txtCode.Text)
                End If
            End If
        Case "mnuItemHelp(206)" 'character to entity ..
            If Not frmMain.ActiveForm Is Nothing Then
                If Len(frmMain.ActiveForm.txtCode.SelText) > 0 Then
                    frmMain.ActiveForm.txtCode.SelText = frmMain.ActiveForm.ConvertHTMLCharacterToEntity(frmMain.ActiveForm.txtCode.SelText)
                Else
                    frmMain.ActiveForm.txtCode.Text = frmMain.ActiveForm.ConvertHTMLCharacterToEntity(frmMain.ActiveForm.txtCode.Text)
                End If
            End If
        Case "mnuItemHelp(300)" 'document.write ...
            Call DocumentWrite
        Case "mnuItemHelp(301)" 'array ...
            If Not frmMain.ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.CreaArray
            End If
        Case "mnuItemHelp(303)" 'comment text ...
            Call SingleComent
        Case "mnuItemHelp(304)" 'line end ...
            Call CreateEndLine
        Case "mnuItemHelp(305)" 'crear string
            If Not ActiveForm Is Nothing Then
                frmString.Show vbModal
            End If
        Case "mnuItemHelp(306)" 'elimina espacios en blanco
            If Not ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.EliminaEspaciosEnBlanco
            End If
        Case "mnuItemHelp(307)" 'elimina lineas
            If Not ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.EliminaLineasEnBlanco
            End If
        Case "mnuItemHelp(308)" 'elimina texto seleccionado
            If Not ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.txtCode.DeleteSel
            End If
        Case "mnuItemHelp(309)"
            If Not ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.CerrarCon(Chr$(34))
            End If
        Case "mnuItemHelp(310)"
            If Not ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.CerrarCon("'")
            End If
        Case "mnuItemHelp(311)"
            If Not ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.InsertaCodigoHtml("(", ")", False)
            End If
        Case "mnuItemHelp(312)"
            If Not ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.InsertaCodigoHtml("[", "]", False)
            End If
        Case "mnuItemHelp(318)"
            If Not ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.InsertaCodigoHtml("{", "}", False)
            End If
        Case "mnuItemHelp(319)"
            If Not ActiveForm Is Nothing Then
                frmFmtSpecial.Show vbModal
            End If
        Case "mnuItemHelp(313)"
            If Not ActiveForm Is Nothing Then
                Call frmMain.ActiveForm.InsertaCodigoHtml("<%", "%>", False)
            End If
        Case "mnuHelp(10)"
            util.ShellFunc "http://www.vbsoftware.cl/tutorial.htm", vbNormalFocus
        Case "mnuProjectTOP(0)"
            'new project
            Call ProjectMan.NewProject
        Case "mnuProjectTOP(1)"
            'manage
            Call ProjectMan.ManageProject
        Case "mnuProjectTOP(2)"
            Call ProjectMan.OpenProject
        Case "mnuProjectTOP(3)"
            'save
            Call ProjectMan.SaveProject
        Case "mnuProjectTOP(4)"
            'quit
            Call ProjectMan.CloseProject
        Case "mnuDOM(1)", "mnuTBDOM(1)"
            'dom
            Call dom_help
        Case "mnuJavascript(1)", "mnuTBJavascript(1)"
            'js 1.3
            Call jshelp("1.3")
        Case "mnuJavascript(2)", "mnuTBJavascript(2)"
            'js 1.4
            Call jshelp("1.4")
        Case "mnuJavascript(3)", "mnuTBJavascript(3)"
            'js 1.5
            Call jshelp("1.5")
        Case "mnuJavascript(5)", "mnuTBJavascript(5)"
            Call jsguide("1.3")
        Case "mnuJavascript(6)", "mnuTBJavascript(6)"
            Call jsguide("1.4")
        Case "mnuJavascript(7)", "mnuTBJavascript(7)"
            Call jsguide("1.5")
        Case "mnuJavascript(9)"
            'navigator version
            frmJsVersion.Show vbModal
        Case "mnuJavascript(10)"
            'object browser
            frmObjExa.Show vbModal
        Case "mnuJavascript(11)"
            frmResWord.Show vbModal
        Case "mnuJavascript(13)"
            'array
            If Not ActiveForm Is Nothing Then
                frmArray.Show vbModal
            End If
        Case "mnuJavascript(14)"
            'block code
            If Not ActiveForm Is Nothing Then
                Call InsertBlock
            End If
        Case "mnuJavascript(15)"
            'end of line
            If Not ActiveForm Is Nothing Then
                CreateEndLine
            End If
        Case "mnuJavascript(16)"
            'escape character
            If Not ActiveForm Is Nothing Then
                frmEscChar.Show vbModal
            End If
        Case "mnuJavascript(17)"
            'multiline
            If Not ActiveForm Is Nothing Then
                Call BlockComment
            End If
        Case "mnuJavascript(18)"
            'regular expresion
            If Not ActiveForm Is Nothing Then
                frmRegExp.Show vbModal
            End If
        Case "mnuJavascript(19)"
            'single comment
            If Not ActiveForm Is Nothing Then
                Call SingleComent
            End If
        Case "mnuJavascript(20)"
            'variable
            If Not ActiveForm Is Nothing Then
                frmNewVar.Show vbModal
            End If
        Case "mnuJavascript(22)"
            'statements
            If Not ActiveForm Is Nothing Then
                frmStatements.Show vbModal
            End If
        Case "mnuJavascript(23)"
            'windows
            If Not ActiveForm Is Nothing Then
                frmNewWindow.Show vbModal
            End If
        Case "mnuJavascript(25)"
            frmjslitopt.Show vbModal
        Case "mnuJScript(1)", "mnuTBJScript(1)"
            'jscript
            util.ShellFunc "http://msdn.microsoft.com/library/default.asp?url=/library/en-us/script56/html/js56jslrfjscriptlanguagereference.asp", vbNormalFocus
        Case "mnuJScript(2)"
            'runtime
            'If Not ActiveForm Is Nothing Then
            '    Call ejecuta_visor_errores("R")
            'End If
        Case "mnuJScript(3)"
            'sintax
            'If Not ActiveForm Is Nothing Then
            '    Call ejecuta_visor_errores("S")
            'End If
        Case "mnuHelp(9)"
            frmTip.Show vbModal
        'Case "mnuOptions(0)"
        '    frmLanMan.Show vbModal
        Case "mnuOptions(1)"
            frmPreferences.Show vbModal
        
        Case "mnuTools(5)"
            'html/xml
            util.ShellFunc "http://validator.w3.org/", vbNormalFocus
        Case "mnuTools(6)"
            'css
            util.ShellFunc "http://jigsaw.w3.org/css-validator/", vbNormalFocus
        Case "mnuTools(7)"
            'hiperlinks
            util.ShellFunc "http://validator.w3.org/checklink", vbNormalFocus
        Case "mnuTools(8)"
            'xml
            util.ShellFunc "http://www.w3.org/2001/03/webdata/xsv", vbNormalFocus
        Case "mnuTools(9)"
            'xml explorer
            Dim glosaxml As String
                        
            glosaxml = "Xml Files (*.xml)|*.xml|"
            glosaxml = glosaxml & "All Files (*.*)|*.*"
        
            If Cdlg.VBGetOpenFileName(Archivo, , , , , , glosaxml, , LastPath, , "XML", Me.hwnd) Then
                
                Dim xmlexp As New cXmlExplorer
                xmlexp.filename = Archivo
                xmlexp.LangPath = util.StripPath(App.Path) & "languages\"
                xmlexp.StartExplorer
            End If
        Case "mnuOptions(0)"
            'frmConfEditor.Show vbModal
        Case "mnuTools(4)"
            If Not ActiveForm Is Nothing Then
                frmPlugMan.Show vbModal
            End If
        Case "mnuItemHelp(1)"
            If Not ActiveForm Is Nothing Then
                get_word_from_cursor
            End If
        Case "mnuTidy(1)"
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    HTidy.ExecuteDefault
                End If
            End If
        Case "mnuJslint(1)"
            If Not ActiveForm Is Nothing Then
                Call ejecuta_jslint
            End If
        Case "mnuEdit(8)"
            If Not ActiveForm Is Nothing Then
                MsgBox "Keep pressed the Ctrl Key and select column with left mouse button.", vbInformation
            End If
        Case "mnuEdit(10)"
            If Not ActiveForm Is Nothing Then
                Call word_count
            End If
        Case "FILE:PREVIEW:1"
            'iexplorer
            Call file_preview(1)
        Case "FILE:PREVIEW:2"
            'firefox
            Call file_preview(2)
        Case "FILE:PREVIEW:3"
            'netscape
            Call file_preview(3)
        Case "FILE:PREVIEW:4"
            'opera
            Call file_preview(4)
        Case "FILE:PREVIEW:5"
            'chrome
            Call file_preview(5)
        Case "mnuFile(0)"
            'Call newEdit
            frmNewDoc.Show vbModal
        Case "mnuFile(2)"
            Call opeEdit
        Case "mnuFile(3)", "FILE:OPEN:FTP"
            frmFtpFiles.Show vbModal
        Case "mnuFile(31)"
            Call UploadFilesToFtp
        Case "mnuFile(5)"
            Call savEdit
        Case "mnuFile(6)"
            'save as
            Call savEdit(True)
        Case "mnuFile(7)"
            'save all
            Call savEdit(False, True)
        Case "mnuFile(9)"
            'close
            Call cloEdit
        Case "mnuFile(10)"
            'close all
            Call cloEdit(True)
        Case "mnuFile(12)"
            'print
            Call prnEdit
        Case "mnuFile(13)", "FILE:PREVIOUS"
            'print preview
            Call PrintPreview
        Case "mnuFile(14)"
            'print setup
            Call prnSetup
        Case "mnuFile(16)"
            Unload Me
        Case "FILE:PREVIEW:5"
            frmConfBrow.Show vbModal
        Case "mnuFile(17)", "FILE:OPEN:WEB"
            'open from web
            frmOpeWeb.Show vbModal
        Case "mnuFile(20)", "FILE:OPEN:FOLDER"
            'open folder
            'Call open_folder
            frmOpenFolder.Show vbModal
        Case "mnuFile(22)"
            'MRU LNUNEZ
            frmOpenMRU.Show vbModal
        Case "mnuFile(21)"
            'save to folder
            Call save_to_folder
        Case "mnuFile(18)"
            'grabar como template
            Call save_template
        Case "mnuFile(19)"
            'grabar a ftp
            If Not frmMain.ActiveForm Is Nothing Then
                Call GuardaArchivoFTP
                'frmSelFilesToUpload.Show vbModal
                'frmSites.tipo_conexion = 1
                'frmSites.Show vbModal
            End If
        Case "mnuFile(23)"
            frmSelFilesToUpload.Show vbModal
        Case "mnuEdit(0)"
            'undo
            Call edtOpe(1)
        Case "mnuEdit(1)"
            'redo
            Call edtOpe(2)
        Case "mnuEdit(3)"
            'cut
            Call edtOpe(3)
        Case "mnuEdit(4)"
            'copy
            Call edtOpe(4)
        Case "mnuEdit(5)"
            'paste
            Call edtOpe(5)
        Case "mnuEdit(22)"
            'paste html
            If Not frmMain.ActiveForm Is Nothing Then
                If CanPasteHTML() Then
                    frmMain.ActiveForm.Insertar GetHTMLClipboard()
                End If
            End If
        Case "mnuEdit(6)"
            'delete
            Call edtOpe(6)
        Case "mnuEdit(8)"
            'column select
        Case "mnuEdit(9)"
            'select all
            Call edtOpe(7)
        Case "mnuEdit(10)"
            'word count
        Case "mnuEdit(12)"
            'increase indent
            Call edtOpe(8)
        Case "mnuEdit(13)"
            'decrease indent
            Call edtOpe(9)
        Case "mnuEdit(16)"
            'uppercase
            Call edtOpe(10)
        Case "mnuEdit(17)"
            'lowercase
            Call edtOpe(11)
        Case "mnuEdit(18)"
            'capitalize
            Call edtOpe(12)
        Case "mnuSearch(0)"
            'find
            Call edtOpe(13)
        Case "mnuSearch(1)"
            'replace
            Call edtOpe(16)
        Case "mnuSearch(3)"
            'find previous
            Call edtOpe(15)
        Case "mnuSearch(4)"
            'find next
            Call edtOpe(14)
        Case "mnuSearch(6)"
            'match bracket
            Call edtOpe(17)
        Case "mnuSearch(8)"
            'goto line
            Call edtOpe(18)
        Case "mnuSearch(9)"
            'find in files
            frmFindFiles.Show vbModal
        Case "mnInsert(0)"
            'activex
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    frmActiveX.Show vbModal
                End If
            End If
        Case "mnInsert(1)"
            'applet
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar LoadSnipet("applet.txt")
                End If
            End If
        Case "mnInsert(2)"
            'style
            If Not ActiveForm Is Nothing Then
                frmStyleSheet.Show vbModal
            End If
        Case "mnInsert(23)"
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar LoadSnipet("css.txt")
                End If
            End If
        Case "mnInsert(5)"
            'hyperlink
            If Not ActiveForm Is Nothing Then
                frmHyperlink.Show vbModal
                'Call Hiperlink
            End If
        Case "mnInsert(6)"
            'horizontal line/ruler
            If Not ActiveForm Is Nothing Then
                frmRuler.Show vbModal
            End If
        Case "mnInsert(7)"
            'image
            If Not ActiveForm Is Nothing Then
                frmImage.Show vbModal
            End If
        Case "mnInsert(8)"
            'space
            If Not ActiveForm Is Nothing Then
                Call BreakSpace
            End If
        Case "mnInsert(9)"
            'noscript
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar(LoadSnipet("noscript.txt"))
                End If
            End If
        Case "mnInsert(12)"
            'script
            If Not ActiveForm Is Nothing Then
                frmScript.Show vbModal
            End If
        Case "mnInsert(13)"
            'symbol
            If Not ActiveForm Is Nothing Then
                frmCharExp.inifile = util.StripPath(App.Path) & "config\htmlmap.ini"
                
                frmCharExp.Show vbModal
            End If
        Case "mnInsert(16)"
            'document.write
            If Not ActiveForm Is Nothing Then
                Call DocumentWrite
            End If
        Case "mnInsert(17)"
            'predefined
            'If Not ActiveForm Is Nothing Then
            '    If frmMain.ActiveForm.Name = "frmEdit" Then
            '        frmPreTemplates.Show vbModal
            '    End If
            'End If
        Case "mnInsert(18)"
            'statusbar
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar(LoadSnipet("statusbar.txt"))
                End If
            End If
        Case "mnInsert(19)"
            'user template
            'If Not ActiveForm Is Nothing Then
            '    frmUserTemplate.Show vbModal
            'End If
        Case "mnInsert(21)"
            'date/time
            If Not ActiveForm Is Nothing Then
                frmDateTime.Show vbModal
            End If
        Case "mnInsert(22)"
            'file/contents
            If Not ActiveForm Is Nothing Then
                Call InsertarArchivo
            End If
        Case "mnInsert_Page(1)"
            'body
            If Not ActiveForm Is Nothing Then
                Call HtmlBody
            End If
        Case "mnInsert_Page(2)"
            'content language
            If Not ActiveForm Is Nothing Then
                frmConLan.Show vbModal
            End If
        Case "mnInsert_Page(3)"
            'doctype
            If Not ActiveForm Is Nothing Then
                frmDocType.Show vbModal
            End If
        Case "mnInsert_Page(4)"
            'encoding
            If Not ActiveForm Is Nothing Then
                frmEncoding.Show vbModal
            End If
        Case "mnInsert_Page(5)"
            'icon page
            If Not ActiveForm Is Nothing Then
                frmHomePage.Show vbModal
            End If
        Case "mnInsert_Page(6)"
            'page title
            If Not ActiveForm Is Nothing Then
                Call HomePageTitle
            End If
        Case "mnInsert_Page(7)"
            'character set
            If Not ActiveForm Is Nothing Then
                frmCharSet.Show vbModal
            End If
        Case "mnInsert_Table(1)"
            'tabla
            If Not ActiveForm Is Nothing Then
                frmTabla.Show vbModal
            End If
        Case "mnInsert_Table(2)"
            'row
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar(LoadSnipet("rows.txt"))
                End If
            End If
        Case "mnInsert_Table(3)"
            'cell
            If Not ActiveForm Is Nothing Then
                frmTableCell.Show vbModal
                'Call ActiveForm.Insertar("<td></td>")
            End If
        Case "mnuFormat(0)"
            'font
            If Not ActiveForm Is Nothing Then
                frmFont.Show vbModal
            End If
        Case "mnuFormat(1)"
            'format paragraph
            If Not ActiveForm Is Nothing Then
                frmFParagraph.Show vbModal
            End If
        Case "mnuFormat(3)"
            'numbered
            If Not ActiveForm Is Nothing Then
                frmListType.mycaption = "Ordered List"
                frmListType.mytype = 1
                frmListType.Show vbModal
            End If
        Case "mnuFormat(4)"
            'bulleted
            If Not ActiveForm Is Nothing Then
                frmListType.mycaption = "Unordered List"
                frmListType.mytype = 1
                frmListType.Show vbModal
            End If
        Case "mnuFormat(6)"
            'big
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<big></big>")
                End If
            End If
        Case "mnuFormat(7)"
            'small
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<small></small>")
                End If
            End If
        Case "mnuFormat_Hea(1)", "FORMAT:HEADING:1"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<h1></h1>")
                End If
            End If
        Case "mnuFormat_Hea(2)", "FORMAT:HEADING:2"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<h2></h2>")
                End If
            End If
        Case "mnuFormat_Hea(3)", "FORMAT:HEADING:3"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<h3></h3>")
                End If
            End If
        Case "mnuFormat_Hea(4)", "FORMAT:HEADING:4"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<h4></h4>")
                End If
            End If
        Case "mnuFormat_Hea(5)", "FORMAT:HEADING:5"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<h5></h5>")
                End If
            End If
        Case "mnuFormat_Hea(6)", "FORMAT:HEADING:6"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<h6></h6>")
                End If
            End If
        Case "mnuFormat(11)"
            'bold
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<b></b>")
                End If
            End If
        Case "mnuFormat(12)"
            'italic
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<i></i>")
                End If
            End If
        Case "mnuFormat(13)"
            'underline
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<u></u>")
                End If
            End If
        Case "mnuFormat(14)"
            'parrafo
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<p></p>")
                End If
            End If
        Case "mnuFormat(15)"
            'pre
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<pre></pre>")
                End If
            End If
        Case "mnuFormat_Align(0)"
            'left
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<p align=""left""></p>")
                End If
            End If
        Case "mnuFormat_Align(1)"
            'center
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<center></center>")
                End If
            End If
        Case "mnuFormat_Align(2)"
            'right
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<p align=""right""></p>")
                End If
            End If
        Case "mnuFormat_Align(3)"
            'justify
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<p align=""justify""></p>")
                End If
            End If
        Case "mnuForms(0)", "FORMS:BUTTON:1"
            'button
            If Not ActiveForm Is Nothing Then
                frmHtmlBoton.tipo_boton = 0
                frmHtmlBoton.Show vbModal
            End If
        Case "mnuForms(1)"
            'checkbox
            If Not ActiveForm Is Nothing Then
                frmHtmlCheck.tipo_control = "checkbox"
                frmHtmlCheck.Show vbModal
            End If
        Case "mnuForms(2)"
            'combobox
            If Not ActiveForm Is Nothing Then
                frmHtmlCombo.Show vbModal
            End If
        Case "mnuForms(3)"
            'file attach
            If Not ActiveForm Is Nothing Then
                frmHtmlFileAttach.Show vbModal
            End If
        Case "mnuForms(4)"
            'forms
            If Not ActiveForm Is Nothing Then
                frmInsForm.Show vbModal
            End If
        Case "mnuForms(5)"
            'hidden entry
            If Not ActiveForm Is Nothing Then
                frmHtmlHidden.Show vbModal
            End If
        Case "mnuForms(6)"
            'listbox
            If Not ActiveForm Is Nothing Then
                frmHtmlListbox.Show vbModal
            End If
        Case "mnuForms(8)"
            'radio button
            If Not ActiveForm Is Nothing Then
                frmHtmlCheck.tipo_control = "radio"
                frmHtmlCheck.Show vbModal
            End If
        Case "mnuForms(9)", "FORMS:BUTTON:2"
            'reset button
            If Not ActiveForm Is Nothing Then
                frmHtmlBoton.tipo_boton = 1
                frmHtmlBoton.Show vbModal
            End If
        Case "mnuForms(10)", "FORMS:BUTTON:3"
            'submit button
            If Not ActiveForm Is Nothing Then
                frmHtmlBoton.tipo_boton = 2
                frmHtmlBoton.Show vbModal
            End If
        Case "mnuForms(11)"
            'textbox
            If Not ActiveForm Is Nothing Then
                frmHtmlText.tipo_texto = "text"
                frmHtmlText.Show vbModal
            End If
        Case "mnuForms(12)"
            'textarea
            If Not ActiveForm Is Nothing Then
                frmHtmlTextArea.Show vbModal
            End If
        Case "mnuForms(13)"
            'password
            If Not ActiveForm Is Nothing Then
                frmHtmlText.tipo_texto = "password"
                frmHtmlText.Show vbModal
            End If
        Case "mnuMacro(0)"
            'record
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdRecordMacro
                    On Error Resume Next
                    Dim x As Integer
                    Dim Path As String
                    
                    Path = util.StripPath(App.Path) & "macros\"
                    For x = 0 To 9
                        SaveMacros Path & x & ".dem", x
                    Next
                End If
            End If
        Case "mnuMacro(1)"
            If Not ActiveForm Is Nothing Then
                frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro1
            End If
        Case "mnuMacro(2)"
            If Not ActiveForm Is Nothing Then
                frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro2
            End If
        Case "mnuMacro(3)"
            If Not ActiveForm Is Nothing Then
                frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro3
            End If
        Case "mnuMacro(4)"
            If Not ActiveForm Is Nothing Then
                frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro4
            End If
        Case "mnuMacro(5)"
            If Not ActiveForm Is Nothing Then
                frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro5
            End If
        Case "mnuMacro(6)"
            If Not ActiveForm Is Nothing Then
                frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro6
            End If
        Case "mnuMacro(7)"
            If Not ActiveForm Is Nothing Then
                frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro7
            End If
        Case "mnuMacro(8)"
            If Not ActiveForm Is Nothing Then
                frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro8
            End If
        Case "mnuMacro(9)"
            If Not ActiveForm Is Nothing Then
                frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro9
            End If
        Case "mnuMacro(10)"
            If Not ActiveForm Is Nothing Then
                frmMain.ActiveForm.txtCode.ExecuteCmd cmCmdPlayMacro10
            End If
        Case "mnuFunctions(0)"
            If Not ActiveForm Is Nothing Then
                Dim funcion As String
                funcion = InputBox("Function Name:", "New Function")
                If Len(Trim$(funcion)) > 0 Then
                    frmMain.ActiveForm.Insertar "function " & funcion & "()" & vbNewLine & "{" & vbNewLine & vbNewLine & "}"
                End If
            End If
        Case "mnuJavascript(26)"
            frmLibraryManager.Show vbModal
        Case "mnuFunctions(3)"
            'refresh
        Case "mnuPlus(0)"
            'add to favorites
            If Not ActiveForm Is Nothing Then
                frmAddFavorites.Show vbModal
            End If
        Case "mnuPlus(1)"
            'countries menu
            If Not ActiveForm Is Nothing Then
                frmCountryMenus.Show vbModal
            End If
        Case "mnuPlus(2)"
            'drop down menu
            If Not ActiveForm Is Nothing Then
                frmDropDownMenu.Show vbModal
            End If
        Case "mnuPlus(3)"
            'email link
            If Not ActiveForm Is Nothing Then
                frmCreateEmail.Show vbModal
            End If
        Case "mnuPlus(4)"
            'iframe
            If Not ActiveForm Is Nothing Then
                frmIframe.Show vbModal
            End If
        Case "mnuPlus(5)"
            'image rollover
            If Not ActiveForm Is Nothing Then
                frmRollover.Show vbModal
            End If
        Case "mnuPlus(6)"
            'last date
            If Not ActiveForm Is Nothing Then
                frmLastModDate.Show vbModal
            End If
        Case "mnuPlus(7)"
            'left menu
            If Not ActiveForm Is Nothing Then
                frmLeftMenu.Show vbModal
            End If
        Case "mnuPlus(8)"
            'metatag
            If Not ActiveForm Is Nothing Then
                frmMetaTag.Show vbModal
            End If
        Case "mnuPlus(9)"
            'page tran
            If Not ActiveForm Is Nothing Then
                frmPageTran.Show vbModal
            End If
        Case "mnuPlus(10)"
            'popup
            If Not ActiveForm Is Nothing Then
                frmPopup.Show vbModal
            End If
        Case "mnuPlus(11)"
            'tree menu
            If Not ActiveForm Is Nothing Then
                frmTreeMenu.Show vbModal
            End If
        Case "mnuPlus(12)"
            'tab menu
            If Not ActiveForm Is Nothing Then
                frmTabMenu.Show vbModal
            End If
        Case "mnuPlus(13)"
            'tab menu
            If Not ActiveForm Is Nothing Then
                
                frmPopupMenu.Show vbModal
            End If
        Case "mnuPlus(14)"
            'calendar
            If Not ActiveForm Is Nothing Then
                frmCalendar.Show vbModal
            End If
        Case "mnuPlus(15)"
            'slideshow
            If Not ActiveForm Is Nothing Then
                frmSlideShow.Show vbModal
            End If
        Case "mnuPlus_CSS(0)"
            'colored scrollbar
            If Not ActiveForm Is Nothing Then
                frmCreateColScrollbar.Show vbModal
            End If
        Case "mnuPlus_CSS(1)"
            'colored scrollbar
            If Not ActiveForm Is Nothing Then
                frmMouseOverLinks.Show vbModal
            End If
        Case "mnuConsole(0)"
            Dim cons As New cDOS
            cons.StartConsole
            Set cons = Nothing
        Case "mnuDataBase_Query(1)"
            Call ejecuta_query_studio
        Case "mnuView_Toolbars(1)"
            'file
            m_cMenu.Checked(ItemNumber) = Not m_cMenu.Checked(ItemNumber)
            activa_toolbars_dock vbalDockContainer1, "FILE", m_cMenu.Checked(ItemNumber)
        Case "mnuView_Toolbars(2)"
            'edit
            m_cMenu.Checked(ItemNumber) = Not m_cMenu.Checked(ItemNumber)
            activa_toolbars_dock vbalDockContainer1, "EDIT", m_cMenu.Checked(ItemNumber)
        Case "mnuView_Toolbars(3)"
            'format
            m_cMenu.Checked(ItemNumber) = Not m_cMenu.Checked(ItemNumber)
            'activa_toolbars_dock vbalDockContainer3, "FORMAT", m_cMenu.Checked(ItemNumber)
            activa_toolbars_dock vbalDockContainer1, "FORMAT", m_cMenu.Checked(ItemNumber)
        Case "mnuView_Toolbars(4)"
            'javascript
            m_cMenu.Checked(ItemNumber) = Not m_cMenu.Checked(ItemNumber)
            activa_toolbars_dock vbalDockContainer1, "JS", m_cMenu.Checked(ItemNumber)
        Case "mnuView_Toolbars(5)"
            'forms
            m_cMenu.Checked(ItemNumber) = Not m_cMenu.Checked(ItemNumber)
            'activa_toolbars_dock vbalDockContainer2, "FORMS", m_cMenu.Checked(ItemNumber)
            activa_toolbars_dock vbalDockContainer1, "FORMS", m_cMenu.Checked(ItemNumber)
        Case "mnuView_Toolbars(6)"
            'plus
            m_cMenu.Checked(ItemNumber) = Not m_cMenu.Checked(ItemNumber)
            'activa_toolbars_dock vbalDockContainer4, "PLUS", m_cMenu.Checked(ItemNumber)
            activa_toolbars_dock vbalDockContainer1, "PLUS", m_cMenu.Checked(ItemNumber)
        Case "mnuView_Toolbars(7)"
            'html
            m_cMenu.Checked(ItemNumber) = Not m_cMenu.Checked(ItemNumber)
            'activa_toolbars_dock vbalDockContainer2, "HTML", m_cMenu.Checked(ItemNumber)
            activa_toolbars_dock vbalDockContainer1, "HTML", m_cMenu.Checked(ItemNumber)
        Case "mnuView_Toolbars(8)"
            'tools
            m_cMenu.Checked(ItemNumber) = Not m_cMenu.Checked(ItemNumber)
            activa_toolbars_dock vbalDockContainer1, "TOOLS", m_cMenu.Checked(ItemNumber)
            
            
        Case "mnuView(0)"
            'file explorer
            selecttab "FILE:EXPLORER", 1
        Case "mnuView(1)"
            'javascript explorer
            selecttab "JAVASCRIPT:BROWSER", 1
        Case "mnuView(2)"
            'markup explorer
            selecttab "MARKUP:BROWSER", 1
        Case "mnuView(3)"
            'clipboard
            selecttab "CLIPBOARD", 1
        Case "mnuView(4)"
            'clipboard
            selecttab "COLOR:BROWSER", 2
        Case "mnuView(5)"
            'clipboard
            selecttab "QUICK:HELP", 1
        Case "mnuView(6)"
            'clipboard
            selecttab "ONLINE:BROWSER", 2
        Case "mnuView(7)"
            'ansi
            selecttab "ANSI:BROWSER", 2
        Case "mnuView(8)"
            'properties
            selecttab "PROPERTY:BROWSER", 2
        'Case "mnuHelp(0)"
            'ayuda app
        Case "mnuHelp(1)"
            'html
            Archivo = util.StripPath(App.Path) & "help\html40.chm"
            If Not ArchivoExiste2(Archivo) Then
                MsgBox "File " & Archivo & " not found!", vbCritical
                Exit Sub
            End If
            
            util.ShellFunc Archivo, vbNormalFocus
        Case "mnuHelp(2)"
            'css
            Archivo = util.StripPath(App.Path) & "help\css.hlp"
            If Not ArchivoExiste2(Archivo) Then
                MsgBox "File " & Archivo & " not found!", vbCritical
                Exit Sub
            End If
            
            util.ShellFunc Archivo, vbNormalFocus
        Case "mnuHelp(6)"
            #If LITE = 1 Then
                frmInfoAbout.Show vbModal
            #Else
                frmAbout.Show vbModal
            #End If
        Case "JS:ESCAPE:1"
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar ("\b")
                End If
            End If
        Case "mnuHelp(7)"
            'tidy
            util.ShellFunc "http://tidy.sourceforge.net", vbNormalFocus
        Case "mnuHelp(8)"
            'jslint
            util.ShellFunc "http://www.crockford.com", vbNormalFocus
        Case "JS:ESCAPE:2"
            'backslage
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar ("\\")
                End If
            End If
        Case "JS:ESCAPE:3"
            'retorno de carro
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar ("\r")
                End If
            End If
        Case "JS:ESCAPE:4"
            'doble comillas
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar ("\""")
                End If
            End If
        Case "JS:ESCAPE:5"
            'form feed
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar ("\f")
                End If
            End If
        Case "JS:ESCAPE:6"
            'tab horizontal
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar ("\t")
                End If
            End If
        Case "JS:ESCAPE:7"
            'line feed
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar ("\n")
                End If
            End If
        Case "JS:ESCAPE:8"
            'single quotation mark
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar ("\'")
                End If
            End If
        Case "JS:STATEMENTS:1"  'do
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar LoadSnipet("jsdowhile.txt")
                End If
            End If
        Case "JS:STATEMENTS:2"  'for
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar LoadSnipet("jsfor.txt")
                End If
            End If
        Case "JS:STATEMENTS:3"  'for in
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar LoadSnipet("jsforin.txt")
                End If
            End If
        Case "JS:STATEMENTS:4"  'function
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar LoadSnipet("jsfunction.txt")
                End If
            End If
        Case "JS:STATEMENTS:5"  'if..else
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar LoadSnipet("jsifthenelse.txt")
                End If
            End If
        Case "JS:STATEMENTS:6"  'switch
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar LoadSnipet("jsswitch.txt")
                End If
            End If
        Case "JS:STATEMENTS:7"  'try...catch
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar LoadSnipet("jstrycatch.txt")
                End If
            End If
        Case "JS:STATEMENTS:8"  'var
            If Not ActiveForm Is Nothing Then
                frmNewVar.Show vbModal
            End If
        Case "JS:STATEMENTS:9"  'while
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar LoadSnipet("jswhile.txt")
                End If
            End If
        Case "JS:STATEMENTS:10" 'with
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    ActiveForm.Insertar LoadSnipet("jswith.txt")
                End If
            End If
        Case "mnInsert_Frames(1)"
            'frameset
            If Not ActiveForm Is Nothing Then
                frmFramesWiz.Show vbModal
            End If
        Case "mnInsert_Frames(2)"
            'frame
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    If frmMain.ActiveForm.Name = "frmEdit" Then
                        ActiveForm.Insertar LoadSnipet("frame.txt")
                    End If
                End If
            End If
        Case "mnInsert_Frames(3)"
        
        Case "mnInsert_Frames(4)"
            'noframes
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    If frmMain.ActiveForm.Name = "frmEdit" Then
                        ActiveForm.Insertar LoadSnipet("noframe.txt")
                    End If
                End If
            End If
        Case "mnuTools(0)"
            'color browser
            If Not ActiveForm Is Nothing Then
                frmColorBrowser.Show vbModal
            End If
        Case "mnuTools(1)"
            'image browser
            Dim ib As New cImageBrowser
            ib.StartBrowse
            Set ib = Nothing
        Case "mnuTools(2)"
            'tidy
            frmTidyConfig.Show vbModal
        Case "mnuTools(3)"
            'jslint
            If Not ActiveForm Is Nothing Then
                Call ejecuta_jslint
            End If
        Case "mnuAnalize(2)"
            'batch analysis
            
            frmBatch.Show vbModal
        Case "mnuTools(10)"
            'dictionary.com
            If Not ActiveForm Is Nothing Then
                Call busca_palabra(1)
            End If
        Case "mnuTools(11)"
            'thesaurus.com
            If Not ActiveForm Is Nothing Then
                Call busca_palabra(2)
            End If
        Case "mnuTools(12)"
            'window list
            If Not ActiveForm Is Nothing Then
                frmWinList.Show vbModal
            End If
        Case "mnuTools(13)"
            'image editor
            Call ejecuta_image_editor
        Case "TOOLS:TIDYCONFIG"
            frmTidyConfig.Show vbModal
        Case "TOOLS:TIDYCLEAN"
            'clear html
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    HTidy.Run "clean_html.tidy"
                End If
            End If
        Case "TOOLS:TIDYXHTML"
            'convert to xhtml
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    HTidy.Run "convert_to_xhtml.tidy"
                End If
            End If
        Case "TOOLS:TIDYCONXML"
            'convert to xml
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    HTidy.Run "convert_to_xml.tidy"
                End If
            End If
        Case "TOOLS:TIDYINDHTMTAG"
            'Indent HTML Tags
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    HTidy.Run "indent_html.tidy"
                End If
            End If
        Case "TOOLS:TIDYUPDFONSTY"
            'Upgrade FONT tags to Styles
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    HTidy.Run "upgrade_font_to_css.tidy"
                End If
            End If
        Case "TOOLS:TIDYVALFIX"
            'Validate and Fix HTML
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    HTidy.Run "validate_and_fix_html.tidy"
                End If
            End If
        Case "TOOLS:TIDYVALHTML"
            'Validate HTML
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    HTidy.Run "validate_html.tidy"
                End If
            End If
        Case "mnuHelp(90)"
            ShowHelpContents
        Case "mnuHelp(91)"
            frmSymbols.Show vbModal
        Case "mnuHelp(92)"
            frmQuickTip.Show
            glbquickon = True
        Case "mnuHelp(3)"
            'home page
            util.ShellFunc "http://www.vbsoftware.cl", vbNormalFocus
        Case "mnuHelp(4)"
            'forum
            util.ShellFunc "http://www.vbsoftware.cl/forum", vbNormalFocus
        Case "mnuHelp(5)"
            'order now
            util.ShellFunc "http://www.vbsoftware.cl/register.htm", vbNormalFocus
        Case "mnInsert_ssi(2)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssiecho.txt"))
                End If
            End If
        Case "mnInsert_ssi(3)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssiexec.txt"))
                End If
            End If
        Case "mnInsert_ssi(4)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssiexeccmd.txt"))
                End If
            End If
        Case "mnInsert_ssi(5)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssilastmod.txt"))
                End If
            End If
        Case "mnInsert_ssi(6)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssilastmodv.txt"))
                End If
            End If
        Case "mnInsert_ssi(7)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssifsize.txt"))
                End If
            End If
        Case "mnInsert_ssi(8)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssifsizev.txt"))
                End If
            End If
        Case "mnInsert_ssi(9)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssigoto.txt"))
                End If
            End If
        Case "mnInsert_ssi(10)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssilabel.txt"))
                End If
            End If
        Case "mnInsert_ssi(11)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssibreak.txt"))
                End If
            End If
        Case "mnInsert_ssi(12)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssisetvar.txt"))
                End If
            End If
        Case "mnInsert_ssi(13)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssiemsg.txt"))
                End If
            End If
        Case "mnInsert_ssi(14)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssiif.txt"))
                End If
            End If
        Case "mnInsert_ssi(15)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssiifelse.txt"))
                End If
            End If
        Case "mnInsert_ssi(17)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    frmTagVar.lang = "SSI"
                    frmTagVar.titulo = "SSI Server Variables"
                    frmTagVar.File = "ssivar.ini"
                    frmTagVar.prefijo = "#echo var="
                    frmTagVar.Show vbModal
                End If
            End If
        Case "mnInsert_ssi(19)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssiincludefile.txt"))
                End If
            End If
        Case "mnInsert_ssi(20)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("ssiincludevirtual.txt"))
                End If
            End If
        Case "mnInsert_php(2)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("phpif.txt"))
                End If
            End If
        Case "mnInsert_php(18)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("phpifelse.txt"))
                End If
            End If
        Case "mnInsert_php(3)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("phpifelseif.txt"))
                End If
            End If
        Case "mnInsert_php(4)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("phpswitch.txt"))
                End If
            End If
        Case "mnInsert_php(5)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("phpfor.txt"))
                End If
            End If
        Case "mnInsert_php(6)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("phpwhile.txt"))
                End If
            End If
        Case "mnInsert_php(8)"
            'asp variable
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    frmTagVar.lang = "PHP"
                    frmTagVar.titulo = "PHP Server Variables"
                    frmTagVar.File = "phpvar.ini"
                    frmTagVar.prefijo = "$_SERVER"
                    frmTagVar.Show vbModal
                End If
            End If
        Case "mnInsert_php(10)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("phptag.txt"))
                End If
            End If
        Case "mnInsert_php(11)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("phpblock.txt"))
                End If
            End If
        Case "mnInsert_php(12)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("phpoutputtag.txt"))
                End If
            End If
        Case "mnInsert_php(13)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("phpincludefile.txt"))
                End If
            End If
        Case "mnInsert_php(14)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("phprequirefile.txt"))
                End If
            End If
        Case "mnInsert_php(16)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("phpcblock.txt"))
                End If
            End If
        Case "mnInsert_php(17)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("phplcomm.txt"))
                End If
            End If
        Case "mnInsert_php(19)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("mysql_connect.txt"))
                End If
            End If
        Case "mnInsert_php(20)"
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("mysql_query.txt"))
                End If
            End If
        Case "mnInsert_asp(2)"
            'if then
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("aspifthen.txt"))
                End If
            End If
        Case "mnInsert_asp(3)"
            'if then else
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("aspifthenelse.txt"))
                End If
            End If
        Case "mnInsert_asp(4)"
            'switch
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("aspselcase.txt"))
                End If
            End If
        Case "mnInsert_asp(5)"
            'for
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("aspfor.txt"))
                End If
            End If
        Case "mnInsert_asp(6)"
            'do while
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("aspdowhile.txt"))
                End If
            End If
        Case "mnInsert_asp(7)"
            'loop until
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("aspdountil.txt"))
                End If
            End If
        Case "mnInsert_asp(8)"
            '<if then>
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("aspifthen2.txt"))
                End If
            End If
        Case "mnInsert_asp(9)"
            '<else>
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("aspelse.txt"))
                End If
            End If
        Case "mnInsert_asp(10)"
            '<end if>
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("aspendif.txt"))
                End If
            End If
        Case "mnInsert_asp(12)"
            'asp variable
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    frmTagVar.lang = "ASP"
                    frmTagVar.titulo = "ASP Server Variables"
                    frmTagVar.File = "aspvar.ini"
                    frmTagVar.prefijo = "Request.ServerVariables"
                    frmTagVar.Show vbModal
                End If
            End If
        Case "mnInsert_asp(14)"
            'asp tag
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("asptag.txt"))
                End If
            End If
        Case "mnInsert_asp(15)"
            'asp block
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("aspblock.txt"))
                End If
            End If
        Case "mnInsert_asp(16)"
            'asp output tag
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("aspoutputtag.txt"))
                End If
            End If
        Case "mnInsert_asp(17)"
            'include file
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("aspincludefile.txt"))
                End If
            End If
        Case "mnInsert_asp(18)"
            'include virtual
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("aspincludevirtual.txt"))
                End If
            End If
        Case "mnInsert_asp(20)"
            'include virtual
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("asp_connect_db.txt"))
                End If
            End If
        Case "mnInsert_asp(21)"
            'include virtual
            If Not ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    Call frmMain.ActiveForm.Insertar(LoadSnipet("asp_fetch.txt"))
                End If
            End If
        Case Else
            'If InStr(m_cMenu.ItemKey(ItemNumber), "mnuPlugins") Then
            '    Call Plugins.Run(m_cMenu.ItemKey(ItemNumber))
            'End If
                        
            If Left$(m_cMenu.ItemKey(ItemNumber), 10) = "mnuFileMRU" Then
                opeEdit m_cMenu.HelpText(ItemNumber)
            End If
            
            If Left$(m_cMenu.ItemKey(ItemNumber), 17) = "mnuFunctions_Date" Then
                If Not frmMain.ActiveForm Is Nothing Then
                    If frmMain.ActiveForm.Name = "frmEdit" Then
                        frmMain.ActiveForm.Insertar m_cMenu.Caption(ItemNumber)
                    End If
                End If
            End If
            
            If Left$(m_cMenu.ItemKey(ItemNumber), 21) = "mnuFunctions_Document" Then
                If Not frmMain.ActiveForm Is Nothing Then
                    If frmMain.ActiveForm.Name = "frmEdit" Then
                        frmMain.ActiveForm.Insertar m_cMenu.Caption(ItemNumber)
                    End If
                End If
            End If
            
            If Left$(m_cMenu.ItemKey(ItemNumber), 18) = "mnuFunctions_Event" Then
                If Not frmMain.ActiveForm Is Nothing Then
                    If frmMain.ActiveForm.Name = "frmEdit" Then
                        frmMain.ActiveForm.Insertar m_cMenu.Caption(ItemNumber)
                    End If
                End If
            End If
            
            If Left$(m_cMenu.ItemKey(ItemNumber), 17) = "mnuFunctions_Math" Then
                If Not frmMain.ActiveForm Is Nothing Then
                    If frmMain.ActiveForm.Name = "frmEdit" Then
                        frmMain.ActiveForm.Insertar m_cMenu.Caption(ItemNumber)
                    End If
                End If
            End If
            
            If Left$(m_cMenu.ItemKey(ItemNumber), 19) = "mnuFunctions_String" Then
                If Not frmMain.ActiveForm Is Nothing Then
                    If frmMain.ActiveForm.Name = "frmEdit" Then
                        frmMain.ActiveForm.Insertar m_cMenu.Caption(ItemNumber)
                    End If
                End If
            End If
            
            If Left$(m_cMenu.ItemKey(ItemNumber), 19) = "mnuFunctions_Window" Then
                If Not frmMain.ActiveForm Is Nothing Then
                    If frmMain.ActiveForm.Name = "frmEdit" Then
                        frmMain.ActiveForm.Insertar m_cMenu.Caption(ItemNumber)
                    End If
                End If
            End If
    End Select
    
    Set CodeLib = Nothing
    Set str = Nothing
    
End Sub

Private Sub activa_toolbar_edit()

    With tbrEdit
        .ImageSource = CTBExternalImageList
        .SetImageList m_MainImg, CTBImageListNormal
        .DrawStyle = CTBDrawOfficeXPStyle
        .CreateToolbar 16, True, True ', True
        ' Now we create the toolbar:
        .CreateFromMenu2 m_cMenu, CTBToolbarStyle, "EDTTOOLBAR"
    End With
    
End Sub

Private Sub activa_toolbar_file()

    With tbrFile
        .ImageSource = CTBExternalImageList
        .SetImageList m_MainImg, CTBImageListNormal
        .DrawStyle = CTBDrawOfficeXPStyle
        .CreateToolbar 16, True, True ', True
        ' Now we create the toolbar:
        .CreateFromMenu2 m_cMenu, CTBToolbarStyle, "FILETOOLBAR"
        .Visible = True
    End With
                
End Sub

Private Sub activa_toolbar_format()

    With tbrFormat
        .ImageSource = CTBExternalImageList
        .SetImageList m_MainImg, CTBImageListNormal
        .DrawStyle = CTBDrawOfficeXPStyle
        .CreateToolbar 16, True, True ', True
        ' Now we create the toolbar:
        .CreateFromMenu2 m_cMenu, CTBToolbarStyle, "FMTTOOLBAR"
        .ButtonStyle("FORMAT:HEADING") = CTBDropDownArrow
    End With
    
End Sub

Private Sub activa_toolbar_forms()
    
    With tbrForms
        .ImageSource = CTBExternalImageList
        .SetImageList m_MainImg, CTBImageListNormal
        .DrawStyle = CTBDrawOfficeXPStyle
        .CreateToolbar 16, True, True ', True
        ' Now we create the toolbar:
        .CreateFromMenu2 m_cMenu, CTBToolbarStyle, "FORMSTOOLBAR"
        .ButtonStyle("FORMS:BUTTON") = CTBDropDownArrow
    End With
    
End Sub

Private Sub activa_toolbar_html()

    With tbrHtm
        .ImageSource = CTBExternalImageList
        .SetImageList m_MainImg, CTBImageListNormal
        .DrawStyle = CTBDrawOfficeXPStyle
        .CreateToolbar 16, True, True ', True
        ' Now we create the toolbar:
        .CreateFromMenu2 m_cMenu, CTBToolbarStyle, "HTMTOOLBAR"
    End With
    
End Sub

Private Sub activa_toolbar_js()

    With tbrJs
        .ImageSource = CTBExternalImageList
        .SetImageList m_MainImg, CTBImageListNormal
        .DrawStyle = CTBDrawOfficeXPStyle
        .CreateToolbar 16, True, True ', True
        ' Now we create the toolbar:
        .CreateFromMenu2 m_cMenu, CTBToolbarStyle, "JSTOOLBAR"
    End With
    
End Sub

Private Sub activa_toolbar_plus()

    With tbrPlus
        .ImageSource = CTBExternalImageList
        .SetImageList m_MainImg, CTBImageListNormal
        .DrawStyle = CTBDrawOfficeXPStyle
        .CreateToolbar 16, True, True ', True
        ' Now we create the toolbar:
        .CreateFromMenu2 m_cMenu, CTBToolbarStyle, "PLUSTOOLBAR"
    End With
    
End Sub

Private Sub activa_toolbar_tools()

    With tbrTools
        .ImageSource = CTBExternalImageList
        .SetImageList m_MainImg, CTBImageListNormal
        .DrawStyle = CTBDrawOfficeXPStyle
        .CreateToolbar 16, True, True ', True
        ' Now we create the toolbar:
        .CreateFromMenu2 m_cMenu, CTBToolbarStyle, "TOOLSTOOLBAR"
    End With
    
End Sub

Private Sub activa_toolbars()

    Dim Accion As Variant
    'Dim docked As Variant
    Dim ini As String
    
    ini = IniPath
    
    With vbalDockContainer1
        .Add _
         "MENU", _
         tbrMenu.ToolbarWidth, tbrMenu.ToolbarHeight, getVerticalHeight(tbrMenu), tbrMenu.MaxButtonWidth, _
         "Menu Bar", , , , , , False, False
        .Capture "MENU", tbrMenu.hwnd
    End With
    
    If InIDE() Then Exit Sub
    
    Call dock_toolbar(vbalDockContainer1, "FILE", Accion, tbrFile, "File")
    Call dock_toolbar(vbalDockContainer1, "EDIT", Accion, tbrEdit, "Edit")
    Call dock_toolbar(vbalDockContainer1, "JS", Accion, tbrJs, "Javascript")
    Call dock_toolbar(vbalDockContainer1, "TOOLS", Accion, tbrTools, "Tools")
    Call dock_toolbar(vbalDockContainer1, "FORMS", Accion, tbrForms, "Forms")
    Call dock_toolbar(vbalDockContainer1, "HTM", Accion, tbrHtm, "HTML")
    Call dock_toolbar(vbalDockContainer1, "FORMAT", Accion, tbrFormat, "Format")
    Call dock_toolbar(vbalDockContainer1, "PLUS", Accion, tbrPlus, "Plus")
        
End Sub
Private Sub activa_toolbars_dock(DockContainer As vbalDockContainer, ByVal sKey As String, ByVal Estado As Boolean)
            
    On Error Resume Next
    
    Dim ini As String
    
    ini = IniPath
    Select Case sKey
        Case "FILE"
            If Estado Then
                Call activa_toolbar_file
                Call dock_toolbar(DockContainer, "FILE", 1, tbrFile, "File")
                Call util.GrabaIni(ini, "toolbars", "FILE_visible", "1")
            Else
                tbrFile.DestroyToolBar
                Call dock_toolbar(DockContainer, "FILE", 2, tbrFile, "File")
                Call util.GrabaIni(ini, "toolbars", "FILE_visible", "0")
            End If
        Case "EDIT"
            If Estado Then
                Call activa_toolbar_edit
                Call dock_toolbar(DockContainer, "EDIT", 1, tbrEdit, "Edit")
                Call util.GrabaIni(ini, "toolbars", "EDIT_visible", "1")
            Else
                tbrEdit.DestroyToolBar
                Call dock_toolbar(DockContainer, "EDIT", 2, tbrEdit, "Edit")
                Call util.GrabaIni(ini, "toolbars", "EDIT_visible", "0")
            End If
        Case "FORMAT"
            If Estado Then
                Call activa_toolbar_format
                Call dock_toolbar(DockContainer, "FORMAT", 1, tbrFormat, "Format")
                Call util.GrabaIni(ini, "toolbars", "FORMAT_visible", "1")
            Else
                tbrFormat.DestroyToolBar
                Call dock_toolbar(DockContainer, "FORMAT", 2, tbrFormat, "Format")
                Call util.GrabaIni(ini, "toolbars", "FORMAT_visible", "0")
            End If
        Case "FORMS"
            If Estado Then
                Call activa_toolbar_forms
                Call dock_toolbar(DockContainer, "FORMS", 1, tbrForms, "Forms")
                Call util.GrabaIni(ini, "toolbars", "FORMS_visible", "1")
            Else
                tbrForms.DestroyToolBar
                Call dock_toolbar(DockContainer, "FORMS", 2, tbrForms, "Forms")
                Call util.GrabaIni(ini, "toolbars", "FORMS_visible", "0")
            End If
        Case "JS"
            If Estado Then
                Call activa_toolbar_js
                Call dock_toolbar(DockContainer, "JS", 1, tbrJs, "Javascript")
                Call util.GrabaIni(ini, "toolbars", "JS_visible", "1")
            Else
                tbrJs.DestroyToolBar
                Call dock_toolbar(DockContainer, "JS", 2, tbrJs, "Javascript")
                Call util.GrabaIni(ini, "toolbars", "JS_visible", "0")
            End If
        Case "PLUS"
            If Estado Then
                Call activa_toolbar_plus
                Call dock_toolbar(DockContainer, "PLUS", 1, tbrPlus, "Plus")
                Call util.GrabaIni(ini, "toolbars", "PLUS_visible", "1")
            Else
                tbrPlus.DestroyToolBar
                Call dock_toolbar(DockContainer, "PLUS", 2, tbrPlus, "Plus")
                Call util.GrabaIni(ini, "toolbars", "PLUS_visible", "0")
            End If
        Case "HTML"
            If Estado Then
                Call activa_toolbar_html
                Call dock_toolbar(DockContainer, "HTM", 1, tbrHtm, "Html")
                Call util.GrabaIni(ini, "toolbars", "HTM_visible", "1")
            Else
                tbrHtm.DestroyToolBar
                Call dock_toolbar(DockContainer, "HTM", 2, tbrPlus, "Html")
                Call util.GrabaIni(ini, "toolbars", "HTM_visible", "0")
            End If
        Case "TOOLS"
            If Estado Then
                Call activa_toolbar_tools
                Call dock_toolbar(DockContainer, "TOOLS", 1, tbrTools, "Tools")
                Call util.GrabaIni(ini, "toolbars", "TOOLS_visible", "1")
            Else
                tbrTools.DestroyToolBar
                Call dock_toolbar(DockContainer, "TOOLS", 2, tbrPlus, "Tools")
                Call util.GrabaIni(ini, "toolbars", "TOOLS_visible", "0")
            End If
    End Select
    
    Err = 0
    
End Sub

Public Sub activar_editor(ByVal hwnd As Long)
    mdiforms_TabClick hwnd, vbLeftButton, 168, -10
End Sub

Private Sub busca_palabra(ByVal Indice As Integer)

    Dim texto As String
    Dim sh As String
    
    texto = frmMain.ActiveForm.txtCode.SelText
    
    If Indice = 1 Then
        sh = "http://dictionary.reference.com/search?q=" & texto
    Else
        sh = "http://thesaurus.reference.com/search?q=" & texto
    End If
    
    Call util.ShellFunc(sh, vbNormalFocus)
    
End Sub

Private Sub buscar_hwnd()

    Dim frm As Form
    Dim File As New cFile
    Dim k As Integer
    'Dim titulo As String
    
    For Each frm In Forms
        If TypeName(frm) = "frmEdit" Then
            If frm.hwnd = ActiveForm.hwnd Then
                For k = 1 To Files.Files.count
                    Set File = New cFile
                    Set File = Files.Files.ITem(k)
                    If File.Caption = Replace(ActiveForm.Caption, "*", "") Then
                       If CStr(File.IdDoc) = CStr(ActiveForm.Tag) Then
                           If Len(File.filename) > 0 Then
                              picStatus.Panels(1).Text = util.PathArchivo(File.filename)
                              picStatus.Panels(1).Tag = File.filename
                           Else
                              picStatus.Panels(1).Text = vbNullString
                              picStatus.Panels(1).Tag = vbNullString
                           End If
                        End If
                    End If
                    Set File = Nothing
                Next k
            End If
        End If
    Next
    
End Sub

Private Sub cargar_explorer_bars()

    'help explorer
    Dim ceBar As cExplorerBar
    Dim ceItem As cExplorerBarItem
    Dim arr_help_sec() As String
    Dim arr_help_ite() As String
    Dim ini As String
    Dim k As Integer
    Dim C As Integer
    Dim j As Integer
    Dim i As Integer
    
    ini = util.StripPath(App.Path) & "config\help.ini"
    
    With HlpExp
        .ImageList = frmMain.m_MainImg.hIml
        Set ceBar = .Bars.Add(, "HELP", "Help")

        ceBar.ToolTipText = "Online help"
        Set ceItem = ceBar.Items.Add(, "HELP:1", "Online Help", 85)
        ceItem.Tag = util.StripPath(App.Path) & "help\jsplus.hlp"
        ceItem.ToolTipText = "The Javascript Plus! Online Help"

        'reference
        Set ceItem = ceBar.Items.Add(, "JSHELP:1", "JavaScript Reference 1.3", 85)
        ceItem.Tag = util.StripPath(App.Path) & "reference\referenceJS13\index.htm"
        ceItem.ToolTipText = "The JavaScript Reference 1.3"

        Set ceItem = ceBar.Items.Add(, "JSHELP:2", "JavaScript Reference 1.4", 85)
        ceItem.Tag = util.StripPath(App.Path) & "reference\ReferenceJS14\index.htm"
        ceItem.ToolTipText = "The JavaScript Reference 1.4"
        
        Set ceItem = ceBar.Items.Add(, "JSHELP:3", "JavaScript Reference 1.5", 85)
        ceItem.Tag = util.StripPath(App.Path) & "reference\ReferenceJS15\contents.html"
        ceItem.ToolTipText = "The JavaScript Reference 1.5"
        
        'core guide
        Set ceItem = ceBar.Items.Add(, "JSHELP:4", "JavaScript Guide 1.3", 85)
        ceItem.Tag = util.StripPath(App.Path) & "reference\GuideJS13\index.htm"
        ceItem.ToolTipText = "The JavaScript Guide 1.3"

        Set ceItem = ceBar.Items.Add(, "JSHELP:5", "JavaScript Guide 1.4", 85)
        ceItem.Tag = util.StripPath(App.Path) & "reference\GuideJS14\index.htm"
        ceItem.ToolTipText = "The JavaScript Guide 1.4"
        
        Set ceItem = ceBar.Items.Add(, "JSHELP:6", "JavaScript Guide 1.5", 85)
        ceItem.Tag = util.StripPath(App.Path) & "reference\GuideJS15\contents.html"
        ceItem.ToolTipText = "The JavaScript Guide 1.5"
        
        Set ceItem = ceBar.Items.Add(, "HELP:3", "HTML Help", 85)
        ceItem.Tag = util.StripPath(App.Path) & "help\html40.chm"
        ceItem.ToolTipText = "The Html Online Help"

        Set ceItem = ceBar.Items.Add(, "HELP:4", "CSS Help", 85)
        ceItem.Tag = util.StripPath(App.Path) & "help\css.hlp"
        ceItem.ToolTipText = "The CSS Online Help"
        
        Set ceItem = ceBar.Items.Add(, "HELP:5", "Tidy Reference", 85)
        ceItem.Tag = "http://tidy.sourceforge.net"
        ceItem.ToolTipText = "The Tidy Online Help"
        
        Set ceItem = ceBar.Items.Add(, "HELP:6", "JSLint Reference", 85)
        ceItem.Tag = "http://www.crockford.com"
        ceItem.ToolTipText = "The JSLint Online Help"
        
        'cargar la ayuda dinamica
        get_info_section "help", arr_help_sec(), ini
        C = 1: i = 1
        For k = 1 To UBound(arr_help_sec)
            'cabezera
            Set ceBar = .Bars.Add(, "HELP" & C, Explode(arr_help_sec(k), 1, "|"))
            ceBar.ToolTipText = Explode(arr_help_sec(k), 2, "|")
            
            'leer los itemes
            get_info_section Explode(arr_help_sec(k), 1, "|"), arr_help_ite(), ini
            For j = 1 To UBound(arr_help_ite)
                Set ceItem = ceBar.Items.Add(, "HELPITE:" & i, Explode(arr_help_ite(j), 1, "|"), 96)
                ceItem.Tag = Explode(arr_help_ite(j), 3, "|")
                ceItem.ToolTipText = Explode(arr_help_ite(j), 2, "|")
                i = i + 1
            Next j
            
            C = C + 1
        Next k
        
        
    End With

    Exit Sub
    
    ini = util.StripPath(App.Path) & "config\httpcodes.ini"
    
'    With HlpHttpCodes
'        'cargar la ayuda dinamica
'        get_info_section "groups", arr_help_sec(), ini
'        C = 1: i = 1
'        For k = 2 To UBound(arr_help_sec)
'            'cabezera
'            Set ceBar = .Bars.Add(, "HELP" & C, Explode(arr_help_sec(k), 1, "|"))
'            ceBar.ToolTipText = Explode(arr_help_sec(k), 2, "|")
'
'            'leer los itemes
'            get_info_section Explode(arr_help_sec(k), 1, "|"), arr_help_ite(), ini
'            For j = 2 To UBound(arr_help_ite)
'                Set ceItem = ceBar.Items.Add(, "HELPITE:" & i, Explode(arr_help_ite(j), 1, "|"), 96)
'                ceItem.Tag = Explode(arr_help_ite(j), 3, "|")
'                ceItem.ToolTipText = Explode(arr_help_ite(j), 2, "|")
'                i = i + 1
'            Next j
'
'            C = C + 1
'        Next k
'    End With
End Sub



Private Sub dock_toolbar(DockContainer As vbalDockContainer, ByVal sKey As String, ByVal Accion As Integer, TBR As cToolbar, ByVal Title As String)

    If Accion = 1 Then
        With DockContainer
            .Add sKey, TBR.ToolbarWidth, TBR.ToolbarHeight, getVerticalHeight(TBR), getVerticalWidth(TBR), Title
            .Capture sKey, TBR.hwnd
        End With
     Else
        With DockContainer
            .Remove sKey
        End With
    End If
    
 End Sub

Private Sub dom_help()
    
    Dim Archivo As String
    Dim Path As String
    
    Path = App.Path & "\reference\domref\domref\"
    Archivo = "index.html"
        
    Archivo = Path & Archivo
    
    If ArchivoExiste2(Archivo) Then
        frmHelp.url = Archivo
        Load frmHelp
        frmHelp.Show
    Else
        MsgBox "File not found : " & Archivo, vbCritical
        
        If Confirma("Do you want to install JavaScript Reference and Core Guide") = vbYes Then
            frmInstallHelp.Show vbModal
        End If
    End If
    
End Sub

Private Sub ejecuta_image_editor()

    Dim Archivo As String
    Dim cImEffect As New cImageEffect
    
    '#If LITE = 1 Then
    '    archivo = util.StripPath(App.Path) & "tools\jsimage.exe"
    '#Else
    '    archivo = util.StripPath(App.Path) & "tools\jsimage.exe"
    '#End If
    
    'If ArchivoExiste2(archivo) Then
    '    #If LITE = 1 Then
    '        archivo = util.StripPath(App.Path) & "tools\jsimage.exe /start_unregistered"
    '    #Else
    '        archivo = util.StripPath(App.Path) & "tools\jsimage.exe /start_registered"
    '    #End If
    
    '    Shell archivo, vbNormalFocus
    'Else
    '    MsgBox "File not found", vbCritical
    'End If
    
    cImEffect.Start
    
    Set cImEffect = Nothing
    
End Sub

Private Sub ejecuta_jslint()

    Dim Archivo As String
        
    Archivo = util.StripPath(App.Path) & "jslint\jslint.js"
    
    If Not ArchivoExiste2(Archivo) Then
        MsgBox "File doesn't found : " & Archivo, vbAbortRetryIgnore
        Exit Sub
    End If
        
    Archivo = util.StripPath(App.Path) & "jslint.ini"
    
    If Not ArchivoExiste2(Archivo) Then
        MsgBox "You must first configure the javascript code analizer.", vbInformation
        frmjslitopt.Show vbModal
        Exit Sub
    End If
    
    On Error GoTo Error_Jslint
    
    Dim k As Integer
    Dim tmpfile As String
    Dim nFreeFile As Long
    
    tmpfile = util.StripPath(App.Path) & "jslint\tmpfile.js"
    
    util.BorrarArchivo tmpfile
    
    DoEvents
    
    With frmMain.ActiveForm
        If .txtCode.SelLength > 0 Then
            nFreeFile = FreeFile
            Open tmpfile For Output As #nFreeFile
                Print #nFreeFile, .txtCode.SelText
            Close #nFreeFile
        End If
    End With
                
    Dim frwait As Boolean
    Dim frjslint As Boolean
    
    frwait = True
    Load frmWait
    frmWait.Show
    frmWait.lblaccion.Caption = "Analyzing selected file ..."
    
    frjslint = True
    Load frmJsLint
    Unload frmJsLint
    
    frmMain.ActiveForm.vbsMsg1.LoadJsLintFile
    
    Unload frmWait
    
    Exit Sub:
    
Error_Jslint:
    If frjslint Then Unload frmJsLint
    If frwait Then Unload frmWait
    
End Sub

Private Sub ejecuta_query_studio()

    Dim Archivo As String
    
    #If LITE = 1 Then
        Archivo = util.StripPath(App.Path) & "tools\jsquery.exe"
    #Else
        Archivo = util.StripPath(App.Path) & "tools\jsquery.exe"
    #End If
    
    If ArchivoExiste2(Archivo) Then
        #If LITE = 1 Then
            Archivo = util.StripPath(App.Path) & "tools\jsquery.exe /start_unregistered"
        #Else
            Archivo = util.StripPath(App.Path) & "tools\jsquery.exe /start_registered"
        #End If
    
        Shell Archivo, vbNormalFocus
    Else
        MsgBox "File not found : " & Archivo, vbCritical
    End If
    
End Sub
Private Sub ExecuteTool(ByVal tool As String)

    Dim File As String
    
    File = util.StripPath(App.Path) & "tools\" & tool
    
    If Not ArchivoExiste2(File) Then
        MsgBox "File : " & File & " doesn't exists", vbCritical
        Exit Sub
    End If
    
    Shell File, vbNormalFocus
    
End Sub

Private Sub file_preview(ByVal opt As Integer)

    Dim File As String
    Dim tmpfile As String
    Dim tmpdir As String
    Dim k As Long
    Dim nFreeFile As Long
    Dim Archivo As String
    
    util.Hourglass hwnd, True
    
    If Not frmMain.ActiveForm Is Nothing Then
        If opt = 1 Then
            File = util.LeeIni(IniPath, "browsers", "iexplorer")
            Archivo = "iexplorer.exe"
        ElseIf opt = 2 Then
            File = util.LeeIni(IniPath, "browsers", "firefox")
            Archivo = "firefox.exe"
        ElseIf opt = 3 Then
            File = util.LeeIni(IniPath, "browsers", "netscape")
            Archivo = "netscape.exe"
        ElseIf opt = 4 Then
            File = util.LeeIni(IniPath, "browsers", "opera")
            Archivo = "opera.exe"
        ElseIf opt = 5 Then
            File = util.LeeIni(IniPath, "browsers", "chrome")
            Archivo = "chrome.exe"
        End If
    
        If Len(File) > 0 Then
            If ArchivoExiste2(File) Then
                File = util.GetShortPath(File)
                
                Call buscar_hwnd
                
                If InStr(ActiveForm.Caption, "*") > 0 Then
                    If Confirma("Do you want to save the changes") = vbYes Then
                        Call savEdit
                    End If
                End If
                
                'frmConfBrow.Show vbModal
                
                If Len(picStatus.Panels(1).Tag) > 0 Then
                    'tmpfile = util.GetShortPath(picStatus.Panels(1).tag)
                    tmpfile = picStatus.Panels(1).Tag
                Else
                    tmpfile = util.ArchivoTemporal()
                    tmpfile = Replace(tmpfile, ".tmp", ".htm")
                    tmpdir = util.StripPath(App.Path) & "temp"
                    util.CrearDirectorio tmpdir
                    tmpfile = tmpdir & "\" & util.StripFile(tmpfile)
                    nFreeFile = FreeFile
                
                    Open tmpfile For Output As #nFreeFile
                        With frmMain.ActiveForm
                            For k = 0 To .txtCode.LineCount
                                Print #nFreeFile, .txtCode.GetLine(k)
                            Next k
                        End With
                    Close #nFreeFile
                End If
                
                Dim localhost As String
                
                localhost = util.LeeIni(IniPath, "server", "path")
                If Len(localhost) > 0 Then
                    If Left$(tmpfile, Len(localhost)) = localhost Then
                        tmpfile = "http://localhost" & Mid$(tmpfile, Len(localhost) + 1)
                        tmpfile = Replace(tmpfile, "\", "/")
                        ShellExecute Me.hwnd, "open", File, tmpfile, util.PathArchivo(tmpfile), SW_SHOWNORMAL
                    Else
                        If opt = 2 Or opt = 3 Or opt = 5 Then
                            tmpfile = "file:///" & Replace(tmpfile, " ", "%20")
                            tmpfile = Replace(tmpfile, "\", "/")
                        ElseIf opt = 4 Then
                            tmpfile = "file://localhost/" & Replace(tmpfile, " ", "%20")
                            tmpfile = Replace(tmpfile, "\", "/")
                        End If
                        ShellExecute Me.hwnd, "open", File, tmpfile, util.PathArchivo(tmpfile), SW_SHOWNORMAL
                    End If
                Else
                    'ShellExecute Me.hwnd, "open", File, tmpfile, util.PathArchivo(tmpfile), SW_SHOWNORMAL
                    If opt = 2 Or opt = 3 Or opt = 5 Then
                         tmpfile = "file:///" & Replace(tmpfile, " ", "%20")
                         tmpfile = Replace(tmpfile, "\", "/")
                     ElseIf opt = 4 Then
                         tmpfile = "file://localhost/" & Replace(tmpfile, " ", "%20")
                         tmpfile = Replace(tmpfile, "\", "/")
                     End If
                     ShellExecute Me.hwnd, "open", File, tmpfile, util.PathArchivo(tmpfile), SW_SHOWNORMAL
                End If
            Else
                MsgBox "File :" & File & " doesn't exists", vbCritical
            End If
        Else
            If Confirma("File not found : " & Archivo & vbNewLine & vbNewLine & "Do you want to configure executable path for " & Archivo & "?") = vbYes Then
                frmConfBrow.Show vbModal
            End If
        End If
    End If
    
    util.Hourglass hwnd, False
    
End Sub

Private Sub get_word_from_cursor()

    Dim w As String
    Dim k As Integer
    Dim j As Integer
    
    util.Hourglass hwnd, True
    With ActiveForm
        w = LCase$(.txtCode.CurrentWord)
        If Len(w) > 0 Then
            For k = 1 To UBound(arr_html)
                If arr_html(k).Tag = w Then
                    util.Hourglass hwnd, False
                    frmHtmlHelp.elem = k + 2
                    frmHtmlHelp.Show vbModal
                    Exit Sub
                End If
            Next k
        End If
    End With
    
    'buscar la ayuda por atributos de html
    For k = 1 To UBound(arr_html)
        For j = 1 To UBound(arr_html(k).elems)
            If arr_html(k).elems(j).attribute = w Then
                util.Hourglass hwnd, False
                frmHtmlHelp.elem = k + 2
                frmHtmlHelp.Show vbModal
                Exit Sub
            End If
        Next j
    Next k
            
    'buscar por ayuda de js
    'Dim j As Integer
    Dim ini As String
    'Dim num As String
    Dim tipo As String
    Dim miembro As String
    Dim glosa As String
    Dim objeto As String
    Dim sSections() As String
    
    ini = util.StripPath(App.Path) & "config\jshelp.ini"
        
    If Not ArchivoExiste2(ini) Then
        MsgBox "File not found : " & ini, vbCritical
        Exit Sub
    End If
    
    For j = 1 To UBound(udtObjetos)
        objeto = udtObjetos(j)
        get_info_section objeto, sSections, ini
        
        For k = 2 To UBound(sSections)
            glosa = sSections(k)
            If Len(glosa) > 0 Then
                miembro = util.Explode(glosa, 1, "#")
                tipo = util.Explode(glosa, 2, "#")
                glosa = util.Explode(glosa, 3, "#")
                
                If Len(miembro) > 0 Then
                    If InStr(miembro, "(") Then
                        miembro = Left$(miembro, InStr(miembro, "(") - 1)
                    End If
                    If LCase$(miembro) = LCase$(w) Then
                        util.Hourglass hwnd, False
                        frmItemHelp.member = miembro
                        frmItemHelp.mtype = tipo
                        frmItemHelp.mhelp = glosa
                        frmItemHelp.Show vbModal
                        Exit Sub
                    End If
                End If
            End If
        Next k
    Next j
    
    util.Hourglass hwnd, False
    
End Sub

Private Sub GuardaArchivoFTP()

    Dim j As Integer
    Dim File As New cFile
    
    util.Hourglass hwnd, True
    
    For j = 1 To Files.Files.count
        Set File = New cFile
        Set File = Files.Files.ITem(j)
        
        If File.IdDoc = CInt(ActiveForm.Tag) Then
            If File.Ftp Then
                Call File.SaveFile(ActiveForm, False)
                Exit For
            Else
                If Len(File.filename) > 0 Then
                    frmFtpFiles.tipo_conexion = 1
                    frmFtpFiles.localfilename = File.filename
                    frmFtpFiles.Show vbModal
                Else
                    MsgBox "Please save the document first", vbCritical
                End If
                Exit For
            End If
        End If
        
        Set File = Nothing
    Next j
    
    util.Hourglass hwnd, False
    
End Sub

Private Sub jsguide(ByVal Version As String)

    Dim Archivo As String
    Dim Path As String
    
    Path = App.Path & "\reference\guide"
    If Version = "1.3" Then
        Path = Path & "JS13\"
        Archivo = "index.htm"
    ElseIf Version = "1.4" Then
        Path = Path & "JS14\"
        Archivo = "index.htm"
    ElseIf Version = "1.5" Then
        Path = Path & "JS15\"
        Archivo = "intro.html"
    End If
    
    Archivo = Path & Archivo
    
    If ArchivoExiste2(Archivo) Then
        'util.ShellFunc Archivo, vbNormalFocus
        frmHelp.url = Archivo
        Load frmHelp
        frmHelp.Show
    Else
        MsgBox "File not found : " & Archivo, vbCritical
        If Confirma("Do you want to install JavaScript Reference and Core Guide") = vbYes Then
            frmInstallHelp.Show vbModal
        End If
    End If
    
End Sub

Private Sub jshelp(ByVal Version As String)

    Dim Archivo As String
    Dim Path As String
    
    Path = App.Path & "\reference\reference"
    If Version = "1.3" Then
        Path = Path & "JS13\"
        Archivo = "contents.htm"
    ElseIf Version = "1.4" Then
        Path = Path & "JS14\"
        Archivo = "contents.htm"
    ElseIf Version = "1.5" Then
        Path = Path & "JS15\"
        Archivo = "contents.html"
    End If
    
    Archivo = Path & Archivo
    
    If ArchivoExiste2(Archivo) Then
        'util.ShellFunc Archivo, vbNormalFocus
        
        frmHelp.url = Archivo
        Load frmHelp
        frmHelp.Show
    Else
        MsgBox "File not found : " & Archivo, vbCritical
        If Confirma("Do you want to install JavaScript Reference and Core Guide") = vbYes Then
            frmInstallHelp.Show vbModal
        End If
    End If
    
End Sub




Public Function LoadSnipet(ByVal File As String) As String
    
    Dim nFreeFile As Long
    Dim Archivo As String
    Dim ret As New cStringBuilder
    
    Archivo = util.StripPath(App.Path) & "snipets\" & File
    
    If ArchivoExiste2(Archivo) Then
        nFreeFile = FreeFile
        
        Open Archivo For Input As #nFreeFile
            ret.Append Input(LOF(nFreeFile), nFreeFile)
        Close #nFreeFile
    End If
    
    LoadSnipet = ret.ToString
    
End Function

Public Sub open_folder(ByVal Path As String)

    'Dim Path As String
    Dim File As cFile
    Dim afiles() As String
    Dim k As Integer
    Dim sFileName As String
    
    'Path = util.BrowseFolder(hwnd)
    
    If Len(Path) > 0 Then
        'abrir el archivo
        util.Hourglass hwnd, True
            
        Call get_files_from_folder(Path, afiles)
        
        Load frmOpenFiles
        frmOpenFiles.pgb.Max = UBound(afiles)
        frmOpenFiles.Show
        
        For k = 1 To UBound(afiles)
            sFileName = afiles(k)
            frmOpenFiles.lblFile.Caption = sFileName
            Files.filename = sFileName
            
            If frmOpenFiles.Cancelo Then
                Unload frmOpenFiles
                Exit For
            End If
            
            If Files.IsOpen() Then
                Unload frmOpenFiles
                MsgBox "File : " & sFileName & " already open", vbCritical
                Exit Sub
            End If
            
            'verificar que el archivo a abrir exista
            If Not ArchivoExiste2(sFileName) Then
                Unload frmOpenFiles
                MsgBox "File : " & sFileName & " doesn't exist.", vbCritical
                Exit Sub
            End If
                                
            If ListaLangs.IsValidExt(sFileName) Then
                Set File = New cFile
                
                File.IdDoc = Files.GetId
                File.filename = sFileName
                File.Caption = util.StripFile(sFileName)
                    
                'agregar archivo a la lista de archivos
                Files.Add File
                Call lvwOpeFiles.ListItems.Add(, "k" & File.IdDoc, File.Caption)
                lvwOpeFiles.ListItems("k" & File.IdDoc).SubItems(1) = PathArchivo(File.filename)
                lbltot.Caption = CStr(lvwOpeFiles.ListItems.count) & " documents"
                frmOpenFiles.pgb.Value = k
                Set File = Nothing
            End If
            'DoEvents
        Next k
        Unload frmOpenFiles
        util.Hourglass hwnd, False
    End If
    
End Sub

Public Sub opeweb(ByVal webfile As String)

    Dim tmpfile As String
    
    If FTPManager.open_from_web(webfile, tmpfile) Then
        Call newEdit
        
        ActiveForm.txtCode.OpenFile tmpfile
        
        ActiveForm.LoadFunctions tmpfile
        ActiveForm.CodeExp1.UrlSource = webfile
        ActiveForm.CodeExp1.LoadCode tmpfile
        ListaLangs.SetLang tmpfile, ActiveForm.txtCode
        util.BorrarArchivo tmpfile
        
        'modificado
        ActiveForm.txtCode.Modified = True
    End If
    
End Sub

Private Sub cloEdit(Optional ByVal bCloseAll As Boolean = False)

    If Not ActiveForm Is Nothing Then
        If Not bCloseAll Then
            Unload ActiveForm
        Else
            frmClose.Show vbModal
        End If
    End If
    
End Sub

Private Sub edtOpe(ByVal Index As Integer)

    If Not ActiveForm Is Nothing Then
        If Index = 1 Then
            'undo
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    ActiveForm.txtCode.ExecuteCmd cmCmdUndo
                End If
            End If
        ElseIf Index = 2 Then
            'redo
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    ActiveForm.txtCode.ExecuteCmd cmCmdRedo
                End If
            End If
        ElseIf Index = 3 Then
            'cut
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    If ActiveForm.txtCode.CanCut Then
                        ActiveForm.txtCode.Cut
                    End If
                End If
            End If
        ElseIf Index = 4 Then
            'copy
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    If ActiveForm.txtCode.CanCopy Then
                        ActiveForm.txtCode.Copy
                    End If
                End If
            End If
        ElseIf Index = 5 Then
            'paste
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    If ActiveForm.txtCode.CanPaste Then
                        ActiveForm.txtCode.Paste
                    End If
                End If
            End If
        ElseIf Index = 6 Then
            'delete
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    ActiveForm.txtCode.ExecuteCmd cmCmdDelete
                End If
            End If
        ElseIf Index = 7 Then
            'select all
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    ActiveForm.txtCode.ExecuteCmd cmCmdSelectAll
                End If
            End If
        ElseIf Index = 8 Then
            'increase indent
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    ActiveForm.txtCode.ExecuteCmd cmCmdIndentSelection
                End If
            End If
        ElseIf Index = 9 Then
            'increase indent
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    ActiveForm.txtCode.ExecuteCmd cmCmdUnindentSelection
                End If
            End If
        ElseIf Index = 10 Then
            'uppercase
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    ActiveForm.txtCode.ExecuteCmd cmCmdUppercaseSelection
                End If
            End If
        ElseIf Index = 11 Then
            'lowercase
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    ActiveForm.txtCode.ExecuteCmd cmCmdLowercaseSelection
                End If
            End If
        ElseIf Index = 12 Then
            'capitalize
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    ActiveForm.txtCode.ExecuteCmd cmCmdWordCapitalize
                End If
            End If
        ElseIf Index = 13 Then
            'find
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    ActiveForm.txtCode.ExecuteCmd cmCmdFind
                End If
            End If
        ElseIf Index = 14 Then
            'find next
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    ActiveForm.txtCode.ExecuteCmd cmCmdFindNext
                End If
            End If
        ElseIf Index = 15 Then
            'find prev
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    ActiveForm.txtCode.ExecuteCmd cmCmdFindPrev
                End If
            End If
        ElseIf Index = 16 Then
            'replace
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    ActiveForm.txtCode.ExecuteCmd cmCmdFindReplace
                End If
            End If
        ElseIf Index = 17 Then
            'match bracket
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    ActiveForm.txtCode.ExecuteCmd cmCmdGotoMatchBrace
                End If
            End If
        ElseIf Index = 18 Then
            'goto line
            If Not frmMain.ActiveForm Is Nothing Then
                If frmMain.ActiveForm.Name = "frmEdit" Then
                    ActiveForm.txtCode.ExecuteCmd cmCmdGotoLine, -1
                End If
            End If
        End If
    End If
    
End Sub

Private Sub crea_toolbars()
    
    Call activa_toolbar_file
    
    If Not InIDE() Then
        Call activa_toolbar_edit
        Call activa_toolbar_format
        Call activa_toolbar_forms
        Call activa_toolbar_js
        Call activa_toolbar_plus
        Call activa_toolbar_html
        Call activa_toolbar_tools
    End If
    
End Sub


Private Sub crea_tabs()

    Dim tabx As cTab

    tabLeft.Pinned = False
    tabLeft.ImageList = m_MainImg.hIml

    Set tabx = tabLeft.Tabs.Add("FILE:EXPLORER", , "Files", 88)
    tabx.Panel = picTab(0)
    tabx.ToolTipText = "Shows or hides the File explorer"
    tabx.CanClose = False
    tabx.Selected = True
    
    Set tabx = tabLeft.Tabs.Add("FILE:OPENFILES", , "File Manager", 194)
    tabx.Panel = picTab(6)
    tabx.ToolTipText = "Shows or hides the File Manager"
    tabx.CanClose = False
    tabx.Selected = True
    
    Set tabx = tabLeft.Tabs.Add("JAVASCRIPT:BROWSER", , "JavaScript", 98)
    Set tabx.Panel = picTab(1)
    tabx.CanClose = False
    tabx.ToolTipText = "Shows or hides the Javascript Browser"
    
    Set tabx = tabLeft.Tabs.Add("MARKUP:BROWSER", , "Markup", 99)
    Set tabx.Panel = picTab(2)
    tabx.CanClose = False
    tabx.ToolTipText = "Shows or hides the Markup Browser window"
            
    Set tabx = tabLeft.Tabs.Add("XHTML:HELP", , "XHTML Help", 143)
    Set tabx.Panel = picTab(3)
    tabx.CanClose = False
    tabx.ToolTipText = "Shows or hides the XHTML Help window"
    
    Set tabx = tabLeft.Tabs.Add("DHTML:HELP", , "DHTML Help", 149)
    Set tabx.Panel = picTab(4)
    tabx.CanClose = False
    tabx.ToolTipText = "Shows or hides the DHTML Help window"
    
    Set tabx = tabLeft.Tabs.Add("CSS:HELP", , "CSS Help", 125)
    Set tabx.Panel = picTab(5)
    tabx.CanClose = False
    tabx.ToolTipText = "Shows or hides the CSS Help window"
        
    tabRight.Pinned = False
    tabRight.ImageList = m_MainImg.hIml
    
    Set tabx = tabRight.Tabs.Add("ONLINE:BROWSER", , "Online Help", 85)
    Set tabx.Panel = picRight(0)
    tabx.CanClose = False
    tabx.ToolTipText = "Shows or hides the Online Help window"

    Set tabx = tabRight.Tabs.Add("COLOR:BROWSER", , "Color Picker", 100)
    Set tabx.Panel = picRight(1)
    tabx.CanClose = False
    tabx.ToolTipText = "Shows or hides the Color picker window"

    Set tabx = tabRight.Tabs.Add("CLIPBOARD", , "Clipboard", 2)
    Set tabx.Panel = picRight(2)
    tabx.CanClose = False
    tabx.ToolTipText = "Shows or hides the Clipboard window"
    
    'Set tabx = tabRight.Tabs.Add("HTTPERROR:BROWSER", , "HTTP Error Codes", 92)
    'Set tabx.Panel = picRight(3)
    'tabx.CanClose = False
    'tabx.ToolTipText = "Shows or hides the HTTP Error Codes window"
    
    Set tabx = tabRight.Tabs.Add("CODELIB:EXPLORER", , "Active Library", 185)
    tabx.Panel = picRight(4)
    tabx.ToolTipText = "Shows or hides the Active Library"
    tabx.CanClose = False
    tabx.Selected = True
    
    Call cargar_explorer_bars
        
End Sub

Public Sub opeEdit(Optional sFileName As String = vbNullString)

    Dim File As New cFile
    Dim kblimit As String
    Dim chkana As String
    
    kblimit = util.LeeIni(IniPath, "maxkb", "value")
    If kblimit = "" Then
        kblimit = "1000"
    End If
    
    If Len(sFileName) > 0 Then
        sFileName = ObtenerNombreLargoArchivo(sFileName)
        'verificar que el archivo no este abierto por menu archivo y ahora abriendolo desde lista de archivos usados
        Files.filename = sFileName
        If Files.IsOpen() Then
            MsgBox "File " & sFileName & " already open", vbCritical
            Exit Sub
        End If
        
        'verificar que el archivo a abrir exista
        If Not ArchivoExiste2(sFileName) Then
            MsgBox "File " & sFileName & " doesn't exist.", vbCritical
            Exit Sub
        End If
        
        'tamaño del archivo
        If lngGetFileSize(sFileName) > CDbl(kblimit) Then
            If Confirma("Do you want to open this file. File size is greater than KB limit allowed to open a file - " & kblimit & " KB.") = vbNo Then
                Exit Sub
            End If
        End If
        
        If Not ListaLangs.IsValidExt(sFileName) Then
            If Confirma("Unknown file extension. (" & GetFileExtension(sFileName) & ")" & vbNewLine & "Do you want to open the selected file and associate a file extension ?") = vbNo Then
                Exit Sub
            Else
                frmOpenAs.Show vbModal
            End If
        End If
        
        'abrir el archivo
        util.Hourglass hwnd, True
                
        File.IdDoc = Files.GetId
        File.filename = sFileName
        File.Caption = util.StripFile(sFileName)
        
        If File.OpenFile Then
            
            'analizar archivo despues de abrir ?
            chkana = util.LeeIni(IniPath, "startup", "checkjs")
            If chkana = "1" Then
               If GetFileExtension(sFileName) = "js" Then
                  Call ejecuta_jslint
               End If
            End If
    
            'agregar archivo a la lista de archivos
            Files.Add File
            
            Call lvwOpeFiles.ListItems.Add(, "k" & File.IdDoc, File.Caption)
            lvwOpeFiles.ListItems("k" & File.IdDoc).SubItems(1) = PathArchivo(File.filename)
        Else
            MsgBox "Failed to open selected file.", vbCritical
        End If
                
        util.Hourglass hwnd, False
    Else
        'pedir abrir el archivo
        If Files.GetOpenFile() Then
            'verificar que el archivo no este abierto
            If Not Files.IsOpen() Then
            
                'verificar que el archivo a abrir exista
                If Not ArchivoExiste2(Files.filename) Then
                    MsgBox "File doesn't exist.", vbCritical
                    Exit Sub
                End If
        
                If Not ListaLangs.IsValidExt(Files.filename) Then
                    If Confirma("Unknown file extension. Do you want to open the selected file and associate a file extension ?") = vbNo Then
                        Exit Sub
                    Else
                        frmOpenAs.Show vbModal
                    End If
                End If
                
                'tamaño del archivo
                If lngGetFileSize(Files.filename) > CDbl(kblimit) Then
                    If Confirma("Do you want to open this file. File size is greater than KB limit allowed to open a file - " & kblimit & " KB.") = vbNo Then
                        Exit Sub
                    End If
                End If
                
                util.Hourglass hwnd, True
                
                File.IdDoc = Files.GetId
                File.filename = Files.filename
                File.Caption = util.StripFile(File.filename)
                
                If File.OpenFile Then
                    
                    'analizar archivo despues de abrir ?
                     chkana = util.LeeIni(IniPath, "startup", "checkjs")
                     If chkana = "1" Then
                        If GetFileExtension(Files.filename) = "js" Then
                           ejecuta_jslint
                        End If
                     End If
            
                    'agregar archivo a la lista de archivos
                    Files.Add File
                    
                    Call lvwOpeFiles.ListItems.Add(, "k" & File.IdDoc, File.Caption)
                    lvwOpeFiles.ListItems("k" & File.IdDoc).SubItems(1) = PathArchivo(File.filename)
                Else
                    MsgBox "Failed to open selected file.", vbCritical
                End If
                
                util.Hourglass hwnd, False
            Else
                MsgBox "File already open", vbCritical
            End If
        End If
    End If
    
    lbltot.Caption = CStr(lvwOpeFiles.ListItems.count) & " documents"
    
    ListaLangs.TempExtension = vbNullString
    
    Set File = Nothing
    
    util.Hourglass hwnd, False
    
End Sub
Private Sub prnEdit()

    If Not ActiveForm Is Nothing Then
        If ActiveForm.Name = "frmEdit" Then
            On Error Resume Next
            Call ActiveForm.txtCode.PrintContents(Printer.hdc, 0)
            
            If Err <> 0 Then
                MsgBox Error$, vbCritical
            End If
            Err = 0
        End If
    End If
                
End Sub

Sub prnSetup()

    On Error GoTo error_setup
    
    Dim pp As New cPrintPreview
    
    pp.PageSetup
    
    Exit Sub
    
error_setup:
    Err = 0
    
End Sub


Public Sub RemoverArchivoExplorador(ByVal IdDoc As Integer)

    Dim k As Integer
    
    For k = 1 To lvwOpeFiles.ListItems.count
        If CInt(Mid$(lvwOpeFiles.ListItems(k).key, 2)) = IdDoc Then
            lvwOpeFiles.ListItems.Remove k
            Exit For
        End If
    Next k
    
    lbltot.Caption = CStr(lvwOpeFiles.ListItems.count) & " documents"
    
End Sub

Private Sub save_template()

    Dim Archivo As String
    Dim nFreeFile As Long
    
    If Not frmMain.ActiveForm Is Nothing Then
        Archivo = InputStr("Input FileName", "Save As Template")
        
        If Len(Archivo) > 0 Then
            On Error Resume Next
            nFreeFile = FreeFile
            
            If Not ArchivoExiste2(Archivo) Then
                Archivo = util.StripPath(App.Path) & "templates\" & Archivo & ".txt"
                Open Archivo For Output As #nFreeFile
                    Print #nFreeFile, "1"
                Close #nFreeFile
                
                If Err = 0 Then
                    frmMain.ActiveForm.txtCode.SaveFile Archivo, False
                Else
                    MsgBox "Invalid template name", vbCritical
                End If
            Else
                MsgBox "File already exists", vbCritical
            End If
        End If
    End If
    
End Sub

Private Sub save_to_folder()
    
    Dim Path As String
    
    If Not ActiveForm Is Nothing Then
        Path = util.BrowseFolder(hwnd)
        If Len(Path) > 0 Then
            Path = util.StripPath(Path)
            Dim frm As Form
        
            For Each frm In Forms
                If TypeName(frm) = "frmEdit" Then
                    Files.SaveAll frm, Path
                End If
            Next
            
            Set frm = Nothing
        End If
    End If
    
End Sub

Public Sub savEdit(Optional ByVal fSaveAs As Boolean = False, Optional ByVal fSaveAll As Boolean = False)
        
    If Not ActiveForm Is Nothing Then
        If Not fSaveAll Then
            Files.Save ActiveForm, fSaveAs
        Else
            Dim frm As Form
        
            For Each frm In Forms
                If TypeName(frm) = "frmEdit" Then
                    Files.Save frm, fSaveAs
                End If
            Next
            
            Set frm = Nothing
        End If
    End If
    
End Sub





Private Sub createMenu()

    Dim iMain As Long
    Dim ip As Long
    Dim ip2 As Long
    Dim iP3 As Long
    Dim k As Integer
    Dim ini  As String
    
    ini = util.StripPath(App.Path) & "config\jshelp.ini"
    
    With m_cMenu
        ' Build bars for MENU & TOOLBAR:
        iMain = .AddItem("MENU", "Menu Bar", , , , , , "MENUBAR")
        
        'menu archivo
        ip = .AddItem("&File", , , iMain, , , , "mnuFileTOP")
        .AddItem "New ..." & vbTab & "Ctrl+N", "Create a new file", , ip, 11, , , "mnuFile(0)"
        .AddItem "-", , , ip, , , , "mnuFile(1)"
        .AddItem "Open ..." & vbTab & "Ctrl+O", "Open a file", , ip, 12, , , "mnuFile(2)"
        .AddItem "Open from Folder ...", "Opens all files from specified folder", , ip, 178, , , "mnuFile(20)"
        .AddItem "-", , , ip, , , , "mnuFile(40)"
        ip2 = .AddItem("FTP", , , ip, , , , "mnuFTP")
        .AddItem "Download files from FTP ...", "Opens files from FTP Server", , ip2, 75, , , "mnuFile(3)"
        .AddItem "Upload files to FTP ...", "Upload files to FTP Server", , ip2, , , , "mnuFile(31)"
        .AddItem "-", , , ip, , , , "mnuFile(41)"
        .AddItem "Open from Web ...", "Opens the specified URL", , ip, 180, , , "mnuFile(17)"
        .AddItem "Open from MRU ...", "Opens all most recent used files", , ip, 179, , , "mnuFile(22)"
        .AddItem "-", , , ip, , , , "mnuFile(4)"
        ip2 = .AddItem("Project", , , ip, , , , "mnuProjectTOP")
        iP3 = .AddItem("New Project", "Creates a new project", , ip2, 159, , , "mnuProjectTOP(0)")
        iP3 = .AddItem("Manage files ...", "Add/Removes files to current project", , ip2, , , , "mnuProjectTOP(1)")
        iP3 = .AddItem("Open Project", "Open a project", , ip2, 194, , , "mnuProjectTOP(2)")
        iP3 = .AddItem("Save Project", "Saves current files inside project", , ip2, , , , "mnuProjectTOP(3)")
        iP3 = .AddItem("Close Project", "Close current project", , ip2, , , , "mnuProjectTOP(4)")
        
        .AddItem "-", , , ip, , , , "mnuFile(4)"
        .AddItem "Save File ..." & vbTab & "Ctrl+S", "Saves the active document", , ip, 13, , , "mnuFile(5)"
        .AddItem "Save File As ...", "Saves the active document with a new name", , ip, , , , "mnuFile(6)"
        .AddItem "Save All Files" & vbTab & "Shift+Ctrl+S", "Saves all open documents", , ip, 14, , , "mnuFile(7)"
        .AddItem "-", , , ip, , , , "mnuFile(144)"
        .AddItem "Save File to FTP ...", "Saves files to FTP Server", , ip, 181, , , "mnuFile(19)"
        .AddItem "Save Files to FTP ...", "Saves files to FTP Server", , ip, 181, , , "mnuFile(23)"
        .AddItem "Save to Folder ...", "Saves file on FTP Server", , ip, , , , "mnuFile(21)"
        '.AddItem "Save As Template ...", "Makes a template from the active document", , ip, , , , "mnuFile(18)"
        .AddItem "-", , , ip, , , , "mnuFile(11)"
        .AddItem "Print" & vbTab & "Ctrl+P", "Prints the active document", , ip, 15, , , "mnuFile(12)"
        .AddItem "Print Preview", "Previous the document before print", , ip, 16, , , "mnuFile(13)"
        .AddItem "Print Setup", "Changes printer settings", , ip, 76, , , "mnuFile(14)"
        .AddItem "-", , , ip, , , , "mnuFile(15)"
        .AddItem "Exit", "Quits the application", , ip, , , , "mnuFile(16)"
                                
        'menu edit
        ip = .AddItem("&Edit", , , iMain, , , , "mnuEditTOP")
        .AddItem "Undo" & vbTab & "Ctrl+Z", "Reverses previous action", , ip, 7, , , "mnuEdit(0)"
        .AddItem "Redo" & vbTab & "Ctrl+Shift+Z", "Repeats previous actions", , ip, 8, , , "mnuEdit(1)"
        .AddItem "-", , , ip, , , , "mnuEdit(2)"
        .AddItem "Cut" & vbTab & "Ctrl+X", "Cuts the selection and puts it on the clipboard", , ip, 0, , , "mnuEdit(3)"
        .AddItem "Copy" & vbTab & "Ctrl+C", "Copies the selection and puts it on the clipboard", , ip, 1, , , "mnuEdit(4)"
        .AddItem "Paste" & vbTab & "Ctrl+V", "Inserts clipboard contents", , ip, 2, , , "mnuEdit(5)"
        .AddItem "Paste HTML", "Inserts html clipboard contents", , ip, 2, , , "mnuEdit(22)"
        .AddItem "Delete", "Clears the selection", , ip, 77, , , "mnuEdit(6)"
        .AddItem "-", , , ip, , , , "mnuEdit(7)"
        .AddItem "Vertical Select", "Selects a column text", , ip, , , , "mnuEdit(8)"
        .AddItem "Select All" & vbTab & "Ctrl+A", "Selects the entire document", , ip, , , , "mnuEdit(9)"
        .AddItem "Word Count", "Shows a document report", , ip, , , , "mnuEdit(10)"
        .AddItem "-", , , ip, , , , "mnuEdit(11)"
        .AddItem "Increase Indent" & vbTab & "Tab", "Increases text indent", , ip, 190, , , "mnuEdit(12)"
        .AddItem "Decrease Indent" & vbTab & "Shift+Tab", "Decreases text indent", , ip, 189, , , "mnuEdit(13)"
        .AddItem "-", , , ip, , , , "mnuEdit(14)"
        ip2 = .AddItem("Change Character Case", , , ip, , , , "mnuEdit(15)")
        iP3 = .AddItem("Convert to Uppercase" & vbTab & "Ctrl+Shift+U", "Converts all characters in the selected text to uppercase", , ip2, 78, , , "mnuEdit(16)")
        iP3 = .AddItem("Convert to Lowercase" & vbTab & "Ctrl+Shift+L", "Converts all characters in the selected text to lowercase", , ip2, , , , "mnuEdit(17)")
        iP3 = .AddItem("Capitalize", "Converts first character of each word to uppercase", , ip2, 79, , , "mnuEdit(18)")
        
        'menu search
        
        .AddItem "-", , , ip, , , , "mnuEdit(19)"
        .AddItem "Find" & vbTab & "Ctrl+F", "Searches for the specified text", , ip, 3, , , "mnuSearch(0)"
        .AddItem "Replace" & vbTab & "Ctrl+R", "Replaces specified text with different text", , ip, 6, , , "mnuSearch(1)"
        .AddItem "-", , , ip, , , , "mnuSearch(2)"
        .AddItem "Find Previous" & vbTab & "Shift+F3", "Finds previous occurence of string", , ip, 4, , , "mnuSearch(3)"
        .AddItem "Find Next" & vbTab & "F3", "Repeats last search", , ip, 5, , , "mnuSearch(4)"
        .AddItem "-", , , ip, , , , "mnuSearch(5)"
        .AddItem "Find in Files ...", "Searches for a string in multiple files", , ip, 139, , , "mnuSearch(9)"
        .AddItem "Match Bracket", "Finds matching bracket", , ip, 87, , , "mnuSearch(6)"
        .AddItem "-", , , ip, , , , "mnuSearch(7)"
        .AddItem "Go to ..." & vbTab & "Ctrl+G", "Moves the insertion point to the specified row", , ip, , , , "mnuSearch(8)"
        .AddItem "-", , , ip
        .AddItem "Lookup at Dictionary.com", "Searches the selected word on dictionary.com", , ip, , , , "mnuTools(10)"
        .AddItem "Lookup at Thesaurus.com", "Searches the selected word on thesaurus.com", , ip, , , , "mnuTools(11)"
        
        'insert
        ip = .AddItem("&Insert", , , iMain, , , , "mnuInsertTOP")
        .AddItem "ActiveX Object", "Creates a object tag", , ip, 110, , , "mnInsert(0)"
        .AddItem "Applet", "Creates a applet tag", , ip, , , , "mnInsert(1)"
        .AddItem "Cascading Style Sheet", "Creates a style sheet tag", , ip, , , , "mnInsert(2)"
        .AddItem "Cascading Style Block", "Creates a style sheet block", , ip, , , , "mnInsert(23)"
        .AddItem "-", , , ip, , , , "mnInsert(3)"
        'forms
        ip2 = .AddItem("Fo&rms", , , ip, , , , "mnuFormsTOP")
        .AddItem "Button ...", "Insert a button tag", , ip2, 39, , , "mnuForms(0)"
        .AddItem "Check Box ...", "Insert a check box tag", , ip2, 32, , , "mnuForms(1)"
        .AddItem "Combo Box ...", "Insert a combo box tag", , ip2, 35, , , "mnuForms(2)"
        .AddItem "File Attach", "Insert a file attach tag", , ip2, 41, , , "mnuForms(3)"
        .AddItem "Form ...", "Insert a form tag", , ip2, 31, , , "mnuForms(4)"
        .AddItem "Hidden Entry ...", "Insert a hidden entry tag", , ip2, 40, , , "mnuForms(5)"
        .AddItem "List Box ...", "Insert a list box tag", , ip2, 34, , , "mnuForms(6)"
        .AddItem "Radio Button ...", "Insert a radio button tag", , ip2, 33, , , "mnuForms(8)"
        .AddItem "Reset Button ...", "Insert a reset button tag", , ip2, , , , "mnuForms(9)"
        .AddItem "Submit Button ...", "Insert a submit button tag", , ip2, , , , "mnuForms(10)"
        .AddItem "Text Box ...", "Insert a  text box tag", , ip2, 36, , , "mnuForms(11)"
        .AddItem "Text Area ...", "Insert a text area tag", , ip2, 38, , , "mnuForms(12)"
        .AddItem "Text Password ...", "Insert a text password tag", , ip2, 37, , , "mnuForms(13)"
        .AddItem "-", , , ip, , , , "mnuForms(14)"
        'frame
        ip2 = .AddItem("Frames", , , ip, , , , "mnInsert_Frames(0)")
        .AddItem "Frameset ...", "Creates frameset tags", , ip2, 116, , , "mnInsert_Frames(1)"
        .AddItem "Frame ...", "Creates frame tag", , ip2, 112, , , "mnInsert_Frames(2)"
        .AddItem "NoFrames ...", "Inserts noframes tags", , ip2, 115, , , "mnInsert_Frames(4)"
        .AddItem "-", , , ip, , , , "mnInsert(4)"
        .AddItem "Hyperlink", "Creates a hiperlink tag", , ip, 96, , , "mnInsert(5)"
        .AddItem "Horizontal Line", "Creates a line tag", , ip, 109, , , "mnInsert(6)"
        .AddItem "Image", "Creates an image tag", , ip, 97, , , "mnInsert(7)"
        .AddItem "Non-breaking space", "Creates a break tag", , ip, , , , "mnInsert(8)"
        .AddItem "Noscript", "Creates a noscript tag", , ip, , , , "mnInsert(9)"
        .AddItem "-", , , ip, , , , "mnInsert(10)"
        ip2 = .AddItem("Page", , , ip, , , , "mnInsert_Page(0)")
        .AddItem "Body", "Inserts a BODY tag", , ip2, , , , "mnInsert_Page(1)"
        .AddItem "Character Set", "Inserts a character set", , ip2, , , , "mnInsert_Page(7)"
        .AddItem "Content Language", "Inserts a Content language tag", , ip2, , , , "mnInsert_Page(2)"
        .AddItem "DOCTYPE", "Inserts a DOCTYPE tag", , ip2, , , , "mnInsert_Page(3)"
        .AddItem "Encoding", "Inserts a encoding tag", , ip2, , , , "mnInsert_Page(4)"
        .AddItem "Icon", "Inserts a SHORTCUT ICON tag", , ip2, , , , "mnInsert_Page(5)"
        .AddItem "Title", "Inserts a TITLE tag", , ip2, , , , "mnInsert_Page(6)"
        .AddItem "-", , , ip, , , , "mnInsert(11)"
        .AddItem "Script", "Creates a script tag", , ip, 103, , , "mnInsert(12)"
        .AddItem "Ansi Character", "View the ansi table characters", , ip, , , , "mnInsert(50)"
        .AddItem "Symbol", "View the html tag characters", , ip, 108, , , "mnInsert(13)"
        .AddItem "-", , , ip, , , , "mnInsert(14)"
        'table
        ip2 = .AddItem("Tables", , , ip, , , , "mnInsert_Table(0)")
        .AddItem "Table ...", "Creates table tags <TABLE>", , ip2, 80, , , "mnInsert_Table(1)"
        .AddItem "Row ...", "Creates table row tags <TR>", , ip2, 82, , , "mnInsert_Table(2)"
        .AddItem "Cell ...", "Creates table cell tags <TD>", , ip2, 81, , , "mnInsert_Table(3)"
        .AddItem "-", , , ip, , , , "mnInsert(15)"
        ip2 = .AddItem("ASP", , , ip, , , , "mnInsert_asp(0)")
        iP3 = .AddItem("Statements", , , ip2, , , , "mnInsert_asp(1)")
        .AddItem "If .. Then", "", , iP3, , , , "mnInsert_asp(2)"
        .AddItem "If .. Then .. Else", "", , iP3, , , , "mnInsert_asp(3)"
        .AddItem "Switch", "", , iP3, , , , "mnInsert_asp(4)"
        .AddItem "For", "", , iP3, , , , "mnInsert_asp(5)"
        .AddItem "Do While", "", , iP3, , , , "mnInsert_asp(6)"
        .AddItem "Loop Until", "", , iP3, , , , "mnInsert_asp(7)"
        .AddItem "<% If Then %>", "", , iP3, , , , "mnInsert_asp(8)"
        .AddItem "<% Else %>", "", , iP3, , , , "mnInsert_asp(9)"
        .AddItem "<% End if %>", "", , iP3, , , , "mnInsert_asp(10)"
        .AddItem "-", , , ip2, , , , "mnInsert_asp(11)"
        .AddItem "ASP Server Variable ...", "", , ip2, , , , "mnInsert_asp(12)"
        .AddItem "-", , , ip2, , , , "mnInsert_asp(13)"
        .AddItem "ASP Tag", "", , ip2, , , , "mnInsert_asp(14)"
        .AddItem "ASP Block", "", , ip2, , , , "mnInsert_asp(15)"
        .AddItem "ASP Output Tag", "", , ip2, , , , "mnInsert_asp(16)"
        .AddItem "Include File", "", , ip2, , , , "mnInsert_asp(17)"
        .AddItem "Include Virtual Path", "", , ip2, , , , "mnInsert_asp(18)"
        .AddItem "-", , , ip2, , , , "mnInsert_asp(19)"
        .AddItem "Connect Database", "", , ip2, , , , "mnInsert_asp(20)"
        .AddItem "Fetch Database", "", , ip2, , , , "mnInsert_asp(21)"
        
        .AddItem "-", , , ip, , , , "mnInsert(23)"
        ip2 = .AddItem("PHP", , , ip, , , , "mnInsert_php(0)")
        iP3 = .AddItem("Statements", , , ip2, , , , "mnInsert_php(1)")
        .AddItem "if", "", , iP3, , , , "mnInsert_php(2)"
        .AddItem "if .. else", "", , iP3, , , , "mnInsert_php(18)"
        .AddItem "if .. else if .. else", "", , iP3, , , , "mnInsert_php(3)"
        .AddItem "switch", "", , iP3, , , , "mnInsert_php(4)"
        .AddItem "for", "", , iP3, , , , "mnInsert_php(5)"
        .AddItem "while", "", , iP3, , , , "mnInsert_php(6)"
        .AddItem "-", , , ip2, , , , "mnInsert_php(7)"
        .AddItem "PHP Server Variable ...", , , ip2, , , , "mnInsert_php(8)"
        .AddItem "-", , , ip2, , , , "mnInsert_php(9)"
        .AddItem "PHP Tag", , , ip2, , , , "mnInsert_php(10)"
        .AddItem "PHP Block", , , ip2, , , , "mnInsert_php(11)"
        .AddItem "PHP Output Tag", , , ip2, , , , "mnInsert_php(12)"
        .AddItem "Include file", , , ip2, , , , "mnInsert_php(13)"
        .AddItem "Require file", , , ip2, , , , "mnInsert_php(14)"
        .AddItem "-", , , ip2, , , , "mnInsert_php(15)"
        .AddItem "PHP Comment Block", , , ip2, , , , "mnInsert_php(16)"
        .AddItem "PHP Line Comment", , , ip2, , , , "mnInsert_php(17)"
        .AddItem "-", , , ip2, , , , "mnInsert_php(18)"
        .AddItem "MySQL Connect", , , ip2, , , , "mnInsert_php(19)"
        .AddItem "MySQL Fetch Array", , , ip2, , , , "mnInsert_php(20)"
        .AddItem "-", , , ip, , , , "mnInsert(23)"
        ip2 = .AddItem("SSI", , , ip, , , , "mnInsert_ssi(0)")
        iP3 = .AddItem("Statements", , , ip2, , , , "mnInsert_ssi(1)")
        .AddItem "Echo", , , iP3, , , , "mnInsert_ssi(2)"
        .AddItem "Exec CGI", , , iP3, , , , "mnInsert_ssi(3)"
        .AddItem "Exec Command", , , iP3, , , , "mnInsert_ssi(4)"
        .AddItem "File Last Modification", , , iP3, , , , "mnInsert_ssi(5)"
        .AddItem "File Last Modification (Virtual)", , , iP3, , , , "mnInsert_ssi(6)"
        .AddItem "File Size", , , iP3, , , , "mnInsert_ssi(7)"
        .AddItem "File Size (Virtual)", , , iP3, , , , "mnInsert_ssi(8)"
        .AddItem "Goto", , , iP3, , , , "mnInsert_ssi(9)"
        .AddItem "Goto Label", , , iP3, , , , "mnInsert_ssi(10)"
        .AddItem "Break HTML Label", , , iP3, , , , "mnInsert_ssi(11)"
        .AddItem "Set Variable", , , iP3, , , , "mnInsert_ssi(12)"
        .AddItem "Change Error Message", , , iP3, , , , "mnInsert_ssi(13)"
        .AddItem "if", , , iP3, , , , "mnInsert_ssi(14)"
        .AddItem "if .. else", , , iP3, , , , "mnInsert_ssi(15)"
        .AddItem "-", , , ip2, , , , "mnInsert_ssi(16)"
        .AddItem "SSI Server Variable ...", , , ip2, , , , "mnInsert_ssi(17)"
        .AddItem "-", , , ip2, , , , "mnInsert_ssi(18)"
        .AddItem "Include File", , , ip2, , , , "mnInsert_ssi(19)"
        .AddItem "Include Virtual Path", , , ip2, , , , "mnInsert_ssi(20)"
        .AddItem "-", , , ip, , , , "mnInsert(23)"
        .AddItem "Document/Write", "Inserts a document.write", , ip, , , , "mnInsert(16)"
        '.AddItem "Predefined Template", "Inserts a predefined template", , ip, 63, , , "mnInsert(17)"
        .AddItem "Statusbar ...", "Inserts a status message", , ip, , , , "mnInsert(18)"
        '.AddItem "User Template ...", "Inserts an user template", , ip, 64, , , "mnInsert(19)"
        .AddItem "-", , , ip, , , , "mnInsert(20)"
        .AddItem "Date/Time", "Inserts date/time", , ip, 111, , , "mnInsert(21)"
        .AddItem "File Contents ...", "Inserts content from file", , ip, , , , "mnInsert(22)"
        
        'format
        ip = .AddItem("Form&at", , , iMain, , , , "mnuFormatTOP")
        .AddItem "Font" & vbTab & "Shift+Ctrl+F", "Creates a font tag", , ip, 17, , , "mnuFormat(0)"
        .AddItem "Paragraph...", "Creates a formatted paragraph tag", , ip, 18, , , "mnuFormat(1)"
        .AddItem "-", , , ip, , , , "mnuFormat(2)"
        .AddItem "Numbered List", "Converts the selection to numbered list", , ip, 19, , , "mnuFormat(3)"
        .AddItem "Bulleted List", "Converts the selection to bulleted list", , ip, 20, , , "mnuFormat(4)"
        .AddItem "-", , , ip, , , , "mnuFormat(5)"
        .AddItem "Big Text", "Inserts <BIG> tags", , ip, 21, , , "mnuFormat(6)"
        .AddItem "Small Text", "Inserts <SMALL> tags", , ip, 22, , , "mnuFormat(7)"
        .AddItem "-", , , ip, , , , "mnuFormat(8)"
        ip2 = .AddItem("Heading", "Inserts heading tags", , ip, 23, , , "mnuFormat(9)")
        iP3 = .AddItem("Heading 1", "Inserts heading 1 tag", , ip2, , , , "mnuFormat_Hea(1)")
        iP3 = .AddItem("Heading 2", "Inserts heading 2 tag", , ip2, , , , "mnuFormat_Hea(2)")
        iP3 = .AddItem("Heading 3", "Inserts heading 3 tag", , ip2, , , , "mnuFormat_Hea(3)")
        iP3 = .AddItem("Heading 4", "Inserts heading 4 tag", , ip2, , , , "mnuFormat_Hea(4)")
        iP3 = .AddItem("Heading 5", "Inserts heading 5 tag", , ip2, , , , "mnuFormat_Hea(5)")
        iP3 = .AddItem("Heading 6", "Inserts heading 6 tag", , ip2, , , , "mnuFormat_Hea(6)")
        .AddItem "-", , , ip, , , , "mnuFormat(10)"
        .AddItem "Bold" & vbTab & "Ctrl+B", "Creates bold tag", , ip, 24, , , "mnuFormat(11)"
        .AddItem "Italic" & vbTab & "Ctrl+I", "Creates italic tag", , ip, 25, , , "mnuFormat(12)"
        .AddItem "Underline" & vbTab & "Ctrl+U", "Creates underline tag", , ip, 26, , , "mnuFormat(13)"
        .AddItem "Paragraph", "Creates a paragraph tag", , ip, 27, , , "mnuFormat(14)"
        .AddItem "Preformat", "Creates a preformat tag", , ip, , , , "mnuFormat(15)"
        .AddItem "-", , , ip, , , , "mnuFormat(16)"
        ip2 = .AddItem("Align", "Paragraph align", , ip, , , , "mnuFormat(17)")
        iP3 = .AddItem("Left", "Align paragraph to left", , ip2, 28, , , "mnuFormat_Align(0)")
        iP3 = .AddItem("Center", "Align paragraph to center", , ip2, 29, , , "mnuFormat_Align(1)")
        iP3 = .AddItem("Right", "Align paragraph to right", , ip2, 30, , , "mnuFormat_Align(2)")
        iP3 = .AddItem("Justify", "Justify paragraph", , ip2, 118, , , "mnuFormat_Align(3)")
        .AddItem "-", , , ip, , , , "mnuFormat(8)"
        .AddItem "Format Special", "Format text using special and custom tags", , ip, , , , "mnuFormat(20)"
        
        'javascript
        ip = .AddItem("&JavaScript", , , iMain, , , , "mnuJavascriptTOP")
        .AddItem "Add Function", "Creates a new JavaScript function", , ip, 89, , , "mnuFunctions(0)"
        .AddItem "Add Library", "Expand development environment using your custom own libraries ...", , ip, 167, , , "mnuJavascript(26)"
        .AddItem "Ofuscate File", "Minimize the size of your JavaScript files ...", , ip, , , , "mnuJavascript(28)"
        
        .AddItem "-", , , ip, , , , "mnuFunctions(3)"
        .AddItem "Navigator Versions", "Display a list with JavaScript and Navigator supported versions", , ip, , , , "mnuJavascript(9)"
        .AddItem "Object Browser" & vbTab & "F2", "Display JavaScript objects with methods, properties and events", , ip, 65, , , "mnuJavascript(10)"
        .AddItem "Reserved Words", "Information about JavaScript reserved words", , ip, , , , "mnuJavascript(11)"
        .AddItem "-", , , ip, , , , "mnuJavascript(12)"
        .AddItem "Array", "Inserts array", , ip, 58, , , "mnuJavascript(13)"
        .AddItem "Block Code", "Inserts a block code", , ip, 66, , , "mnuJavascript(14)"
        .AddItem "Line End", "Inserts a JavaScript line end", , ip, 69, , , "mnuJavascript(15)"
        .AddItem "Escape Character", "Inserts a JavaScript escape character", , ip, 70, , , "mnuJavascript(16)"
        .AddItem "Multiline Comment", "Inserts multiline comment", , ip, 68, , , "mnuJavascript(17)"
        .AddItem "Regular Expression", "Inserts regular expression", , ip, 61, , , "mnuJavascript(18)"
        .AddItem "Single Comment", "Inserts a single comment", , ip, 67, , , "mnuJavascript(19)"
        .AddItem "Variable", "Inserts a new variable", , ip, 59, , , "mnuJavascript(20)"
        .AddItem "-", , , ip, , , , "mnuJavascript(21)"
        .AddItem "Statements", "Inserts JavaScript statements", , ip, 62, , , "mnuJavascript(22)"
        .AddItem "Windows ...", "Inserts JavaScript windows", , ip, , , , "mnuJavascript(23)"
        .AddItem "-", , , ip, , , , "mnuJavascript(21)"
        .AddItem "JavaScript Reference 1.3", "Display JavaScript Reference 1.3", , ip, 85, , , "mnuJavascript(1)"
        .AddItem "JavaScript Reference 1.4", "Display JavaScript Reference 1.4", , ip, 85, , , "mnuJavascript(2)"
        .AddItem "JavaScript Reference 1.5", "Display JavaScript Reference 1.5", , ip, 85, , , "mnuJavascript(3)"
        .AddItem "-", , , ip, , , , "mnuJavascript(4)"
        .AddItem "JavaScript Guide 1.3", "Display JavaScript Guide 1.3", , ip, 85, , , "mnuJavascript(5)"
        .AddItem "JavaScript Guide 1.4", "Display JavaScript Guide 1.4", , ip, 85, , , "mnuJavascript(6)"
        .AddItem "JavaScript Guide 1.5", "Display JavaScript Guide 1.5", , ip, 85, , , "mnuJavascript(7)"
        .AddItem "-", , , ip, , , , "mnuJavascript(8)"
        .AddItem "Tutorial", "Display JavaScript a JavaScript Tutorial", , ip, 85, , , "mnuJavascript(24)"
        .AddItem "-", , , ip, , , , "mnuJavascript(8)"
        .AddItem "How to enable JavaScript", "Display information about how to enable javascript in a webrowser", , ip, 85, , , "mnuJavascript(27)"
        
        'cargar los plugins
        ip = m_cMenu.AddItem("&Plus!", , , iMain, , , , "mnuPlusTOP")
        .AddItem "Add to Favorites", "Creates add to favorites script", , ip, 42, , , "mnuPlus(0)"
        .AddItem "Calendar", "Creates a Calendar", , ip, 164, , , "mnuPlus(14)"
        .AddItem "Countries Menu", "Creates drop down menu with countries names", , ip, 43, , , "mnuPlus(1)"
        .AddItem "Drop Down Menu", "Creates menu drop down with go button", , ip, 35, , , "mnuPlus(2)"
        .AddItem "Email Link", "Creates an email form", , ip, 44, , , "mnuPlus(3)"
        .AddItem "Iframe Wizard", "Creates iframe", , ip, 45, , , "mnuPlus(4)"
        .AddItem "Image Rollover", "Creates image rollover", , ip, 46, , , "mnuPlus(5)"
        .AddItem "Last Modified", "Inserts a last modified date", , ip, 47, , , "mnuPlus(6)"
        .AddItem "MetaTag Wizard", "Creates a metatag", , ip, 49, , , "mnuPlus(8)"
        .AddItem "Page Transitions", "Creates a page transitions effect", , ip, 50, , , "mnuPlus(9)"
        .AddItem "Popup Window", "Creates a popup window", , ip, 51, , , "mnuPlus(10)"
        .AddItem "SlideShow", "Creates a SlideShow", , ip, 166, , , "mnuPlus(15)"
        
        ip2 = .AddItem("Menus", , , ip, , , , "mnuPlus_Menus")
        .AddItem "Left Menu", "Creates a left menu", , ip2, 48, , , "mnuPlus(7)"
        .AddItem "Popup Menu", "Creates a popup menu", , ip2, 154, , , "mnuPlus(13)"
        .AddItem "Tab Menu", "Creates a tab menu", , ip2, 155, , , "mnuPlus(12)"
        .AddItem "Tree Menu", "Creates a tree menu", , ip2, 156, , , "mnuPlus(11)"
                
        ip2 = .AddItem("CSS", , , ip, , , , "mnuPlus_CSS")
        iP3 = .AddItem("Coloured Scrollbar", "Creates a coloured scrollbar", , ip2, 52, , , "mnuPlus_CSS(0)")
        iP3 = .AddItem("MouseOver Text Links", "Creates a mouseover text links", , ip2, 53, , , "mnuPlus_CSS(1)")
          
        'cargar los plugins
        ip = m_cMenu.AddItem("Analyze", , , iMain, , , , "mnuAnalize")
        .AddItem "Configure", "Configures options for JavaScript Analizer", , ip, 176, , , "mnuJavascript(25)"
        .AddItem "Find Errors", "Searches for JavaScript errors in active file.", , ip, 138, , , "mnuTools(3)"
        .AddItem "Batch Mode", "Searches for JavaScript errors in multiple files.", , ip, 177, , , "mnuAnalize(2)"
        .AddItem "-", , , ip
        
        .AddItem "Validate HTML, XHTML", "Validate document html/xhtml tags", , ip, , , , "mnuTools(5)"
        .AddItem "Validate CSS", "Validate document css tags", , ip, , , , "mnuTools(6)"
        .AddItem "Validate Hyperlinks", "Validate document hyperlinks", , ip, , , , "mnuTools(7)"
        .AddItem "Validate XML", "Validate document structure", , ip, , , , "mnuTools(8)"
        .AddItem "-", , , ip
        .AddItem "Tidy (HTML Validator)", "Tidy is a powerfull html validator", , ip, 127, , , "mnuTools(2)"
        
        'code format
        ip = .AddItem("Code Utilities", "Code Utilities", , iMain, , , , "mnuItemHelp(101)")
        
        .AddItem "Insert text at begining ...", "Inserted text at the begining of selected lines", , ip, , , , "mnuItemHelp(400)"
        .AddItem "Insert text at end ...", "Inserted text at the begining of selected lines", , ip, , , , "mnuItemHelp(401)"
        .AddItem "-", , , ip, , , , "mnuItemHelp(230)"
        .AddItem "Convert to HTML Paragraph", "Format the selected text to HTML Paragraph <P>", , ip, 27, , , "mnuItemHelp(200)"
        .AddItem "Convert to HTML Italic", "Format the selected text to HTML Italic <I>", , ip, 25, , , "mnuItemHelp(201)"
        .AddItem "Convert to HTML Bold", "Format the selected text to HTML Bold <B>", , ip, 24, , , "mnuItemHelp(202)"
        .AddItem "Convert to HTML Underline", "Format the selected text to HTML Underline <U>", , ip, 26, , , "mnuItemHelp(203)"
        .AddItem "Convert to HTML Comment", "Format the selected text to HTML Comment <!-- -->", , ip, 95, , , "mnuItemHelp(204)"
        .AddItem "Convert HTML Entities to Character", "Convert HTML Entities / Tags Into Corresponding Characters", , ip, , , , "mnuItemHelp(205)"
        .AddItem "Convert Charactes to HTML Entities", "Convert Characters to HTML Entities / Tags", , ip, , , , "mnuItemHelp(206)"
        .AddItem "-", , , ip, , , , "mnuItemHelp(103)"
        .AddItem "Convert to document.write statement", "Insert document.write(...) to the selected text", , ip, , , , "mnuItemHelp(300)"
        .AddItem "Convert to array element ...", "Convert the selected text to array element", , ip, 58, , , "mnuItemHelp(301)"
        .AddItem "Encode to url format", "Encode selected text to url format", , ip, , , , "mnuItemHelp(314)"
        .AddItem "-", , , ip, , , , "mnuItemHelp(230)"
        .AddItem "Enclose selected text in quotes ...", "Enclose selected text with " & Chr$(34) & "expression" & Chr$(34), , ip, , , , "mnuItemHelp(309)"
        .AddItem "Enclose selected text in single quotes ...", "Enclose selected with into '" & "expression" & "'", , ip, , , , "mnuItemHelp(310)"
        .AddItem "Enclose selected text in () ...", "Enclose selected with (" & "expression" & ")", , ip, , , , "mnuItemHelp(311)"
        .AddItem "Enclose selected text in [] ...", "Enclose selected with [" & "expression" & "]", , ip, , , , "mnuItemHelp(312)"
        .AddItem "Enclose selected text in {} ...", "Enclose selected with {" & "expression" & "}", , ip, , , , "mnuItemHelp(318)"
        .AddItem "Enclose selected text in <% %> ...", "Enclose selected text with <%" & "expression" & "%>", , ip, , , , "mnuItemHelp(313)"
        .AddItem "-", , , ip, , , , "mnuItemHelp(103)"
        
        .AddItem "Enclose selected text in ASP Reponse.Write() ...", "Enclose selected text in Response.Write (...)", , ip, , , , "mnuItemHelp(315)"
        .AddItem "Enclose selected text in PHP echo ...", "Enclose selected text in echo " & Chr$(34) & "expression" & Chr$(34), , ip, , , , "mnuItemHelp(316)"
        .AddItem "Enclose selected text in JSP out.println ...", "Enclose selected text in out.println " & Chr$(34) & "expression" & Chr$(34), , ip, , , , "mnuItemHelp(317)"
        
        .AddItem "-", , , ip, , , , "mnuItemHelp(104)"
        .AddItem "Create a string variable ...", "Convert the selected text to a var string ...", , ip, , , , "mnuItemHelp(305)"
        .AddItem "Comment select Text ...", "Insert // at the begining of the selected text", , ip, 67, , , "mnuItemHelp(303)"
        .AddItem "Insert a line end ...", "Insert ; at end of the selected text", , ip, 69, , , "mnuItemHelp(304)"
        .AddItem "-", , , ip, , , , "mnuItemHelp(105)"
        .AddItem "Remove White Spaces", "Clear all white spaces from the selected text", , ip, , , , "mnuItemHelp(306)"
        .AddItem "Delete Empty Lines", "Delete all empty lines from the selected text", , ip, , , , "mnuItemHelp(307)"
        .AddItem "Delete Selected Text", "Deletes the selected text", , ip, , , , "mnuItemHelp(308)"
        .AddItem "-", , , ip, , , , "mnuItemHelp(106)"
        .AddItem "Convert to Uppercase", "Converts all characters in the selected text to uppercase", , ip, 78, , , "mnuEdit(16)"
        .AddItem "Convert to Lowercase", "Converts all characters in the selected text to lowercase", , ip, , , , "mnuEdit(17)"
        .AddItem "-", , , ip, , , , "mnuItemHelp(106)"
        .AddItem "Format Special ...", "Enclose text using custom tags", , ip, , , , "mnuItemHelp(319)"
        
        'library
        'cargar los plugins
        ip = m_cMenu.AddItem("Library", , , iMain, , , , "mnuLibrary")
        .AddItem "Language Wizard", "Language Wizard", , ip, 182, , , "mnuLibrary(2)"
        .AddItem "Category Wizard", "Category Wizard", , ip, 183, , , "mnuLibrary(1)"
        .AddItem "Browse Library", "Browse the Code Library", , ip, 185, , , "mnuLibrary(3)"
        .AddItem "-", , , ip, , , , "mnuLibrary(4)"
        .AddItem "Save To Library", "Save active document to library", , ip, 184, , , "mnuLibrary(5)"
        
        'tools
        ip = .AddItem("&Tools", , , iMain, , , , "mnuToolsMisc")
        .AddItem "Color Explorer 1", "Shows the color explorer tool", , ip, 100, , , "mnuTools(0)"
        .AddItem "Color Explorer 2", "Shows the color names explorer tool", , ip, 100, , , "mnuColorNames(0)"
        .AddItem "DOS Console", "A very basic DOS Console ...", , ip, 173, , , "mnuConsole(0)"
        .AddItem "-", , , ip
        .AddItem "Bitmap Extractor", "Extract bitmaps from files", , ip, , , , "mnuTools(100)"
        .AddItem "Image Browser", "Shows the internal image browser window", , ip, 126, , , "mnuTools(1)"
        .AddItem "Image Effect", "Add special effects to images", , ip, 174, , , "mnuTools(13)"
        .AddItem "-", , , ip
        .AddItem "Icon Editor", "Active the icon editor", , ip, 175, , , "mnuTools(140)"
        .AddItem "Icon Extractor", "Extract icons from files", , ip, , , , "mnuTools(14)"
        .AddItem "-", , , ip
        .AddItem "XML Explorer", "Explore the xml contents from a selected file", , ip, 168, , , "mnuTools(9)"
        
        'Add-Ins
        ip = .AddItem("Add-Ins", , , iMain, , , , "mnuAddIns")
        .AddItem "Visual Data Manager", "Connect to Databases using ODBC, run querys & export results ...", , ip, 165, , , "mnuDataBase_Query(1)"
        .AddItem "-", , , ip
        .AddItem "Add-In Manager ...", "Start the Add-In Manager", , ip, 147, , , "mnuTools(4)"
        .AddItem "-", , , ip
        .AddItem "New Add-In ...", "Creates a new Add-In using Add-In Wizard", , ip, , , , "mnuAddIn(5)"
        
        'settings
        ip = .AddItem("&Settings", , , iMain, , , , "mnuToolsTOP")
        .AddItem "Configure Browsers", "Configure paths for selected browsers", , ip, , , , "mnuConfigBrowsers(1)"
        .AddItem "Preferences", "Configure JavaScript Plus!", , ip, 158, , , "mnuOptions(1)"
        .AddItem "-", , , ip
                
        'macro
        ip2 = .AddItem("Macros", , , ip, , , , "mnuMacroTOP")
        .AddItem "Record Keystrokes" & vbTab & "Ctrl+Q", "Starts and Stops keystroke recording", , ip2, 83, , , "mnuMacro(0)"
        .AddItem "-", , , ip2
        For k = 1 To 10
            .AddItem "Playback Recording " & vbTab & "Alt+" & k, "Starts and Stops keystroke recording", , ip2, , , , "mnuMacro(" & k & ")"
        Next k
                       
        .AddItem "-", , , ip
        
        Plugins.Load (App.Path)
        
        .AddItem "-", , , ip
        ip2 = .AddItem("Toolbars", , , ip, , , , "mnuView_Toolbars(0)")
        iP3 = .AddItem("Edit", "Show or hides the Edit toolbar", , ip2, , True, , "mnuView_Toolbars(2)")
        iP3 = .AddItem("File", "Show or hides the File toolbar", , ip2, , True, , "mnuView_Toolbars(1)")
        iP3 = .AddItem("Format", "Show or hides the Format toolbar", , ip2, , True, , "mnuView_Toolbars(3)")
        iP3 = .AddItem("Forms", "Show or hides the Forms toolbar", , ip2, , True, , "mnuView_Toolbars(5)")
        iP3 = .AddItem("Html", "Show or hides the html toolbar", , ip2, , True, , "mnuView_Toolbars(7)")
        iP3 = .AddItem("Javascript", "Show or hides the JavaScript toolbar", , ip2, , True, , "mnuView_Toolbars(4)")
        iP3 = .AddItem("Plus", "Show or hides the Plus toolbar", , ip2, , True, , "mnuView_Toolbars(6)")
        iP3 = .AddItem("Tools", "Show or hides the tools toolbar", , ip2, , True, , "mnuView_Toolbars(8)")
        
        'window
        ip = .AddItem("&Window", , , iMain, , , , "mnuWindow_Top")
        .AddItem "Always on Top", "Makes JavaScript Plus! window stay on top", , ip, , , , "mnuWindow_Top(1)"
        .AddItem "-", , , ip
        .AddItem "Cascade", "Arrange windows so they overlap", , ip, 191, , , "mnuWindow_Top(2)"
        .AddItem "Tile Horizontally", "Arrange windows as non-overlapping tiles", , ip, 192, , , "mnuWindow_Top(3)"
        .AddItem "Tile Vertically", "Arrange windows as non-overlapping tiles", , ip, 193, , , "mnuWindow_Top(4)"
        .AddItem "-", , , ip
        .AddItem "Close", "Closes the active window", , ip, 74, , , "mnuFile(9)"
        .AddItem "Close All", "Closes all opened window", , ip, , , , "mnuFile(10)"
        .AddItem "-", , , ip
        .AddItem "Window List" & vbTab & "F11", "Show list of all opened windows", , ip, , , , "mnuTools(12)"
        
        'help
        ip = .AddItem("&Help", , , iMain, , , , "mnuHelpTOP")
        .AddItem "About VBSoftware", "About VBSoftware", , ip, 85, , , "mnuPdf1"
        .AddItem "Analyzing JavaScript Files", "Analyzing JavaScript Files", , ip, 85, , , "mnuPdf2"
        .AddItem "Help Panels", "Help Panels", , ip, 85, , , "mnuPdf3"
        .AddItem "Menu Reference", "Menu Reference", , ip, 85, , , "mnuPdf4"
        .AddItem "-", , , ip
        .AddItem "Icon Help ...", "Display help about the icons used in JavaScript Plus!", , ip, , , , "mnuHelp(91)"
        .AddItem "Quick Tip ...", "Display the quick tip window", , ip, , , , "mnuHelp(92)"
        .AddItem "-", , , ip
        .AddItem "HTML Reference ...", "Displays Hyper Text Markup Language reference", , ip, 85, , , "mnuHelp(1)"
        .AddItem "CSS Reference ...", "Displays Cascading Style Sheet reference", , ip, 85, , , "mnuHelp(2)"
        .AddItem "Tidy Reference ...", "Displays information about tidy", , ip, 85, , , "mnuHelp(7)"
        .AddItem "-", , , ip
        
        .AddItem "-", , , ip, , , , "mnuJavascript(4)"
        .AddItem "DOM Reference", "Display DOM Reference", , ip, 85, , , "mnuDOM(1)"
        .AddItem "-", , , ip, , , , "mnuJavascript(4)"
        
        .AddItem "JScript Reference", "Display JScript Reference", , ip, 85, , , "mnuJScript(1)"
        .AddItem "-", , , ip
        
        .AddItem "JavaScript Plus! Home Page", "Visit JavaScript Plus! Home Page", , ip, , , , "mnuHelp(3)"
                
        #If LITE = 1 Then
            .AddItem "-", , , ip
            .AddItem "O&rder Now...", "Purchase this software through a secure online web site", , ip, 86, , , "mnuHelp(5)"
        #End If
        .AddItem "-", , , ip
        .AddItem "&Tip of the day ...", "Display the tip of the day", , ip, , , , "mnuHelp(9)"
        .AddItem "-", , , ip
        .AddItem "&About ...", "About this software", , ip, , , , "mnuHelp(6)"
        
        'definicion de las toolbars
        'file
        iMain = .AddItem("FILE", "File Toolbar", , , , , , "FILETOOLBAR")
        ip = .AddItem("New", "New (Ctrl+N)", , iMain, 11, , , "FILE:NEW")
        ip = .AddItem("Open", "Open (Ctrl+O)", , iMain, 12, , , "FILE:OPEN")
        ip2 = .AddItem("Open from FTP ...", "Open from FTP", , ip, 75, , , "FILE:OPEN:FTP")
        ip2 = .AddItem("Open from WEB ...", "Open from WEB", , ip, 180, , , "FILE:OPEN:WEB")
        ip2 = .AddItem("Open Folder ...", "Opens all files from specified folder", , ip, 178, , , "FILE:OPEN:FOLDER")
        ip = .AddItem("Save", "Save (Ctrl+S)", , iMain, 13, , , "FILE:SAVE")
        ip = .AddItem("Save All", "Save All (Shift+Ctrl+S)", , iMain, 14, , , "FILE:SAVEALL")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Print", "Print (Ctrl+P)", , iMain, 15, , , "FILE:PRINT")
        ip = .AddItem("Preview", "Preview", , iMain, 16, , , "FILE:PREVIOUS")
        ip2 = .AddItem("Internet Explorer", "Previous active document in Internet Explorer", , ip, 104, , , "FILE:PREVIEW:1")
        ip2 = .AddItem("Mozilla Firefox", "Previous active document in Mozilla Firefox", , ip, 105, , , "FILE:PREVIEW:2")
        ip2 = .AddItem("Netscape", "Previous active document in Netscape", , ip, 106, , , "FILE:PREVIEW:3")
        ip2 = .AddItem("Opera", "Previous active document in Opera", , ip, 107, , , "FILE:PREVIEW:4")
        ip2 = .AddItem("Google Chrome", "Previous active document in Google Chrome", , ip, 195, , , "FILE:PREVIEW:5")
        ip2 = .AddItem("-", , , ip)
        ip2 = .AddItem("Configure Browsers ...", "Configure browsers options", , ip, , , , "FILE:PREVIEW:5")
        
        'tools
        iMain = .AddItem("TOOLS", "Tools Toolbar", , , , , , "TOOLSTOOLBAR")
        ip = .AddItem("Execute", "Execute Tidy", , iMain, 127, , , "TOOLS:TIDYRUN")
        ip2 = .AddItem("Clean HTML", "Clean HTML", , ip, , , , "TOOLS:TIDYCLEAN")
        ip2 = .AddItem("Convert to XHTML", "Convert active document to XHTML", , ip, , , , "TOOLS:TIDYXHTML")
        ip2 = .AddItem("Convert to XML", "Convert active document to XML", , ip, , , , "TOOLS:TIDYCONXML")
        ip2 = .AddItem("Indent HTML Tags", "Indent HTML Tags", , ip, , , , "TOOLS:TIDYINDHTMTAG")
        ip2 = .AddItem("Upgrade FONT tags to Styles", "Upgrade FONT tags to Styles", , ip, , , , "TOOLS:TIDYUPDFONSTY")
        ip2 = .AddItem("Validate and Fix HTML", "Validate and Fix HTML", , ip, , , , "TOOLS:TIDYVALFIX")
        ip2 = .AddItem("Validate HTML", "Validate HTML", , ip, , , , "TOOLS:TIDYVALHTML")
        ip2 = .AddItem("-", , , ip)
        ip2 = .AddItem("Configure Tidy ...", "Configure Tidy", , ip, 127, , , "TOOLS:TIDYCONFIG")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Insert Color", "Insert Color", , iMain, 100, , , "TOOLS:COLOR")
        ip = .AddItem("Image Browser", "Image Browser", , iMain, 126, , , "TOOLS:IMAGE")
        ip = .AddItem("Analize", "Analize JavaScript", , iMain, 138, , , "TOOLS:JSLINT")
        ip = .AddItem("Database", "Launch Database Tool", , iMain, 165, , , "TOOLS:DATABASE")
        ip = .AddItem("Plugins", "Plugin Manager", , iMain, 147, , , "TOOLS:PLUGINS")
        ip = .AddItem("XML Explorer", "XML Explorer", , iMain, 168, , , "TOOLS:XMLEXPLORER")
        
        ip = .AddItem("DOS Console", "DOS Console", , iMain, 173, , , "TOOLS:DOS")
        ip = .AddItem("Image Effect", "Image Effect", , iMain, 174, , , "TOOLS:IMAGEEFFECT")
        ip = .AddItem("Icon Editor", "Icon Editor", , iMain, 175, , , "TOOLS:ICONEDITOR")
        
        'edit
        iMain = .AddItem("EDIT", " Edit Toolbar", , , , , , "EDTTOOLBAR")
        ip = .AddItem("Cut", "Cut (Ctrl+Z)", , iMain, 0, , , "EDIT:CUT")
        ip = .AddItem("Copy", "Copy (Ctrl+C)", , iMain, 1, , , "EDIT:COPY")
        ip = .AddItem("Paste", "Paste (Ctrl+V)", , iMain, 2, , , "EDIT:PASTE")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Find", "Find (Ctrl+F)", , iMain, 3, , , "EDIT:FIND")
        ip = .AddItem("Find Prev", "Find Previous (Shift+F3)", , iMain, 4, , , "EDIT:FIND:PREV")
        ip = .AddItem("Find Next", "Find Next (F3)", , iMain, 5, , , "EDIT:FIND:NEXT")
        ip = .AddItem("Replace", "Replace (Ctrl+R)", , iMain, 6, , , "EDIT:REPLACE")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Undo", "Undo (Ctrl+Z)", , iMain, 7, , , "EDIT:UNDO")
        ip = .AddItem("Redo", "Redo (Ctrl+Shift+Z)", , iMain, 8, , , "EDIT:REDO")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Outdent", "Outdent (Shift+Tab)", , iMain, 189, , , "EDIT:OUTDENT")
        ip = .AddItem("Indent", "Indent (Tab)", , iMain, 190, , , "EDIT:INDENT")
        
        'format
        iMain = .AddItem("FORMAT", "Format Toolbar", , , , , , "FMTTOOLBAR")
        ip = .AddItem("Font", "Font (Shift+Ctrl+F)", , iMain, 17, , , "FORMAT:FONT")
        ip = .AddItem("FParagraph", "Paragraph Format", , iMain, 18, , , "FORMAT:FPARAGRAPH")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Numbered", "Numbered List", , iMain, 19, , , "FORMAT:NUMBERED")
        ip = .AddItem("Bulleted", "Bulleted List", , iMain, 20, , , "FORMAT:BULLETED")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("BigText", "Big Text", , iMain, 21, , , "FORMAT:BIGTEXT")
        ip = .AddItem("SmallText", "Small Text", , iMain, 22, , , "FORMAT:SMALLTEXT")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Heading", "Heading", , iMain, 23, , , "FORMAT:HEADING")
        ip2 = .AddItem("Heading1", "Heading 1", , ip, , , , "FORMAT:HEADING:1")
        ip2 = .AddItem("Heading2", "Heading 2", , ip, , , , "FORMAT:HEADING:2")
        ip2 = .AddItem("Heading3", "Heading 3", , ip, , , , "FORMAT:HEADING:3")
        ip2 = .AddItem("Heading4", "Heading 4", , ip, , , , "FORMAT:HEADING:4")
        ip2 = .AddItem("Heading5", "Heading 5", , ip, , , , "FORMAT:HEADING:5")
        ip2 = .AddItem("Heading6", "Heading 6", , ip, , , , "FORMAT:HEADING:6")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Bold", "Bold (Ctrl+B)", , iMain, 24, , , "FORMAT:BOLD")
        ip = .AddItem("Italic", "Italic (Ctrl+I)", , iMain, 25, , , "FORMAT:ITALIC")
        ip = .AddItem("Underline (Ctrl+U)", "Underline", , iMain, 26, , , "FORMAT:UNDERLINE")
        ip = .AddItem("Paragraph", "Paragraph", , iMain, 27, , , "FORMAT:PARAGRAPH")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("ALeft", "Left", , iMain, 28, , , "FORMAT:ALEFT")
        ip = .AddItem("ACenter", "Center", , iMain, 29, , , "FORMAT:ACENTER")
        ip = .AddItem("ARight", "Right", , iMain, 30, , , "FORMAT:ARIGHT")
        ip = .AddItem("ARight", "Right", , iMain, 118, , , "FORMAT:JUSTIFY")
        
        'htm
        iMain = .AddItem("HTML", "HTML Toolbar", , , , , , "HTMTOOLBAR")
        ip = .AddItem("Hyperlink", "Hyperlink", , iMain, 96, , , "HTM:HYPERLINK")
        ip = .AddItem("Image", "Image", , iMain, 97, , , "HTM:IMAGE")
        ip = .AddItem("Horizontal Line", "Horizontal Line", , iMain, 109, , , "HTM:HORIZONTAL")
        ip = .AddItem("Comment", "Comment", , iMain, 95, , , "HTM:COMMENT")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Tables", "Tables", , iMain, 80, , , "HTM:TABLE")
        ip = .AddItem("Frame", "Frame", , iMain, 116, , , "HTM:FRAME")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Script", "Script", , iMain, 103, , , "HTM:SCRIPT")
        ip = .AddItem("Symbol", "Symbol", , iMain, 108, , , "HTM:SYMBOL")
        
        'forms
        iMain = .AddItem("FORMS", "Forms Toolbar", , , , , , "FORMSTOOLBAR")
        ip = .AddItem("Form", "Form", , iMain, 31, , , "FORMS:FORM")
        ip = .AddItem("Button", "Button", , iMain, 39, , , "FORMS:BUTTON")
        ip2 = .AddItem("Normal Button", "Normal Button", , ip, , , , "FORMS:BUTTON:1")
        ip2 = .AddItem("Reset Button", "Reset Button", , ip, , , , "FORMS:BUTTON:2")
        ip2 = .AddItem("Submit Button", "Submit Button", , ip, , , , "FORMS:BUTTON:3")
        ip = .AddItem("CheckBox", "CheckBox", , iMain, 32, , , "FORMS:CHECKBOX")
        ip = .AddItem("ComboBox", "ComboBox", , iMain, 35, , , "FORMS:COMBOBOX")
        ip = .AddItem("FileAttach", "File Attach", , iMain, 41, , , "FORMS:FILEATTACH")
        ip = .AddItem("Hidden", "Hidden Entry", , iMain, 40, , , "FORMS:HIDDEN")
        ip = .AddItem("ListBox", "ListBox", , iMain, 34, , , "FORMS:LISTBOX")
        ip = .AddItem("PassWord", "PassWord", , iMain, 37, , , "FORMS:PASSWORD")
        ip = .AddItem("RadioButton", "Radio Button", , iMain, 33, , , "FORMS:RADIO")
        ip = .AddItem("Text", "Text", , iMain, 36, , , "FORMS:TEXT")
        ip = .AddItem("TextArea", "TextArea", , iMain, 38, , , "FORMS:TEXTAREA")
        
        'plus
        iMain = .AddItem("PLUS", "Plus Toolbar", , , , , , "PLUSTOOLBAR")
        ip = .AddItem("Favorites", "Add to Favorites", , iMain, 42, , , "PLUS:FAVORITES")
        ip = .AddItem("Calendar", "Create Calendar", , iMain, 164, , , "PLUS:CALENDAR")
        ip = .AddItem("Countries", "Countries Menu", , iMain, 43, , , "PLUS:COUNTRIES")
        ip = .AddItem("DropDown", "Drop Down Menu", , iMain, 35, , , "PLUS:DROPDOWN")
        ip = .AddItem("Email", "Email Form", , iMain, 44, , , "PLUS:EMAIL")
        ip = .AddItem("IFrame", "Insert IFRAME", , iMain, 45, , , "PLUS:IFRAME")
        ip = .AddItem("Rollover", "Image Rollover", , iMain, 46, , , "PLUS:ROLLOVER")
        ip = .AddItem("LastDate", "Last Date", , iMain, 47, , , "PLUS:LASTDATE")
        ip = .AddItem("LeftMenu", "Left Menu", , iMain, 48, , , "PLUS:LEFTMENU")
        ip = .AddItem("MetaTag", "Meta Tag", , iMain, 49, , , "PLUS:METATAG")
        ip = .AddItem("PageTran", "Page Transitions", , iMain, 50, , , "PLUS:PAGETRAN")
        ip = .AddItem("PopupW", "Popup Window", , iMain, 51, , , "PLUS:POPUPWINDOW")
        ip = .AddItem("ColScroll", "Coloured ScrollBar", , iMain, 52, , , "PLUS:COLSCROLL")
        ip = .AddItem("MouseOver", "MouseOver Text Links", , iMain, 53, , , "PLUS:MOUSEOVER")
        ip = .AddItem("PopupMenu", "PopupMenu", , iMain, 154, , , "PLUS:POPUPMENU")
        ip = .AddItem("SlideShow", "Create a SlideShow", , iMain, 166, , , "PLUS:SLIDESHOW")
        ip = .AddItem("TabMenu", "TabMenu", , iMain, 155, , , "PLUS:TABMENU")
        ip = .AddItem("TreeMenu", "TreeMenu", , iMain, 156, , , "PLUS:TREEMENU")
        .HideInfrequentlyUsed = True

        'javascript
        iMain = .AddItem("JAVASCRIPT", "Javascript Toolbar", , , , , , "JSTOOLBAR")
        ip = .AddItem("Help", "Online Help", , iMain, 54, , , "JS:FAVORITES")
        ip2 = .AddItem("DOM Reference", "Display DOM Reference", , ip, 85, , , "mnuTBDOM(1)")
        ip2 = .AddItem("-", , , ip, , , , "mnuTBJavascript(4)")
        ip2 = .AddItem("JavaScript Reference 1.3", "Display JavaScript Reference 1.3", , ip, 85, , , "mnuTBJavascript(1)")
        ip2 = .AddItem("JavaScript Reference 1.4", "Display JavaScript Reference 1.4", , ip, 85, , , "mnuTBJavascript(2)")
        ip2 = .AddItem("JavaScript Reference 1.5", "Display JavaScript Reference 1.5", , ip, 85, , , "mnuTBJavascript(3)")
        ip2 = .AddItem("-", , , ip, , , , "mnuTBJavascript(4)")
        ip2 = .AddItem("JavaScript Guide 1.3", "Display JavaScript Guide 1.3", , ip, 85, , , "mnuTBJavascript(5)")
        ip2 = .AddItem("JavaScript Guide 1.4", "Display JavaScript Guide 1.4", , ip, 85, , , "mnuTBJavascript(6)")
        ip2 = .AddItem("JavaScript Guide 1.5", "Display JavaScript Guide 1.5", , ip, 85, , , "mnuTBJavascript(7)")
        ip2 = .AddItem("-", , , ip, , , , "mnuTBJavascript(8)")
        ip2 = .AddItem("&JScript Reference", "Display JScript Reference", , ip, 85, , , "mnuTBJScript(1)")
            
        ip = .AddItem("Object", "Object Browser", , iMain, 65, , , "JS:OBJECT")
        ip = .AddItem("Library", "Library Manager", , iMain, 167, , , "JS:LIBRARY")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Alert", "Alert", , iMain, 144, , , "JS:ALERT")
        ip = .AddItem("Confirm", "Confirm", , iMain, 145, , , "JS:CONFIRM")
        ip = .AddItem("Prompt", "Prompt", , iMain, 146, , , "JS:PROMPT")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Array", "Array", , iMain, 58, , , "JS:ARRAY")
        ip = .AddItem("Statements", "Statements", , iMain, 62, , , "JS:STATEMENTS")
        ip2 = .AddItem("do...while", "do...while", , ip, , , , "JS:STATEMENTS:1")
        ip2 = .AddItem("for...", "for...", , ip, , , , "JS:STATEMENTS:2")
        ip2 = .AddItem("for...in", "for...in", , ip, , , , "JS:STATEMENTS:3")
        ip2 = .AddItem("function", "function", , ip, , , , "JS:STATEMENTS:4")
        ip2 = .AddItem("if...then...else", "if...then...else", , ip, , , , "JS:STATEMENTS:5")
        ip2 = .AddItem("switch", "switch", , ip, , , , "JS:STATEMENTS:6")
        ip2 = .AddItem("try...catch", "try...catch", , ip, , , , "JS:STATEMENTS:7")
        ip2 = .AddItem("var", "var", , ip, , , , "JS:STATEMENTS:8")
        ip2 = .AddItem("while", "while", , ip, , , , "JS:STATEMENTS:9")
        ip2 = .AddItem("with", "with", , ip, , , , "JS:STATEMENTS:10")
        ip = .AddItem("Escape", "Escape Chars", , iMain, 70, , , "JS:ESCAPE")
        ip2 = .AddItem("Backspace", "Backspace", , ip, , , , "JS:ESCAPE:1")
        ip2 = .AddItem("Backslash", "Backslash", , ip, , , , "JS:ESCAPE:2")
        ip2 = .AddItem("Carriage Return", "Carriage Return", , ip, , , , "JS:ESCAPE:3")
        ip2 = .AddItem("Double quotation mark", "Double quotation mark", , ip, , , , "JS:ESCAPE:4")
        ip2 = .AddItem("Form feed", "Form feed", , ip, , , , "JS:ESCAPE:5")
        ip2 = .AddItem("Horizontal Tab", "Horizontal Tab", , ip, , , , "JS:ESCAPE:6")
        ip2 = .AddItem("Line feed", "Line feed", , ip, , , , "JS:ESCAPE:7")
        ip2 = .AddItem("Single quotation mark", "Single quotation mark", , ip, , , , "JS:ESCAPE:8")
        ip = .AddItem("RegExp", "Regular Expression", , iMain, 61, , , "JS:REGEXP")
        'ip = .AddItem("PreTemplate", "Pretemplate", , iMain, 63, , , "JS:PRETEMPLATE")
        'ip = .AddItem("UserTemplate", "User Template", , iMain, 64, , , "JS:USERTEMPLATE")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Block", "Block", , iMain, 66, , , "JS:BLOCK")
        ip = .AddItem("Commen1", "Single Comment", , iMain, 67, , , "JS:COMMEN1")
        ip = .AddItem("Commen2", "Multi Comment", , iMain, 68, , , "JS:COMMEN2")
        ip = .AddItem("LineEnd", "Line End", , iMain, 69, , , "JS:LINEEND")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Function", "Function", , iMain, 60, , , "JS:FUNCTION")
   End With

    ip = 0
    With m_cMenuPop
        .Clear
        .AddItem "New ...", "New File", , ip, 11, , , "mnuItemTabX(1)"
        .AddItem "-", , , ip, , , , "mnuItemTabX(2)"
        .AddItem "Save File", "Save File", , ip, 13, , , "mnuItemTabX(3)"
        .AddItem "Save File As ...", "Save File", , ip, , , , "mnuItemTabX(4)"
        .AddItem "Save All", "Save File", , ip, 14, , , "mnuItemTabX(5)"
        .AddItem "-", , , ip, , , , "mnuItemTabX(6)"
        .AddItem "Close", "Close Tab", , ip, 74, , , "mnuItemTabX(7)"
        .AddItem "Close All", "Close All Tabs", , ip, , , , "mnuItemTabX(8)"
        .Store "TABPOPUP"
    End With
    
End Sub
Private Sub buildImageLists()
    
    Dim k As Integer
    Dim j As Integer
    
    'file
    Set m_MainImg = New cVBALImageList
    j = 1
    With m_MainImg
        .IconSizeX = 16: .IconSizeY = 16: .ColourDepth = ILC_COLOR24
        .Create
        For k = 101 To c_file
            .AddFromResourceID k, App.hInstance, IMAGE_ICON, "k" & j
            j = j + 1
        Next k
    End With
    
End Sub

Private Function chevronPress(ctl As vbalDockContainer, ByVal sKey As String, ByVal lX As Long, ByVal lY As Long)
   'Debug.Print "ChevronPress Start", sKey
   Select Case sKey
   Case "MENU"
      tbrMenu.chevronPress lX, lY
      ctl.BandSizeChange "MENU", tbrMenu.ToolbarWidth, tbrMenu.ToolbarHeight, getVerticalHeight(tbrMenu), getVerticalWidth(tbrMenu)
   Case "FILE"
      tbrFile.chevronPress lX, lY
      ctl.BandSizeChange "FILE", tbrFile.ToolbarWidth, tbrFile.ToolbarHeight, getVerticalHeight(tbrFile), getVerticalWidth(tbrFile)
   End Select
   'Debug.Print "ChevronPress End", sKey
End Function

Public Function getVerticalHeight(tbrThis As cToolbar) As Long
Dim l As Long
Dim lHeight As Long
Dim lMaxWidth As Long
Dim lRowHeight As Long
Dim lRowWidth As Long

   lMaxWidth = getVerticalWidth(tbrThis)
   
   For l = 0 To tbrThis.ButtonCount - 1
      If tbrThis.ButtonVisible(l) Then
         If tbrThis.ButtonControl(l) = 0 Then
            
            If tbrThis.ButtonStyle(l) = CTBSeparator Then
               ' we'll start a new row for the next one
               lHeight = lHeight + lRowHeight
               lRowHeight = 0
               lRowWidth = 0
            Else
               If (lRowWidth + tbrThis.ButtonWidth(l) > lMaxWidth) Then
                  ' This button needs to go on a new row:
                  lHeight = lHeight + lRowHeight
                  lRowHeight = 0
                  lRowWidth = lRowWidth + tbrThis.ButtonWidth(l)
                  If (tbrThis.ButtonHeight(l) > lRowHeight) Then
                     lRowHeight = tbrThis.ButtonHeight(l)
                  End If
               Else
                  ' This button goes on this row:
                  If (tbrThis.ButtonHeight(l) > lRowHeight) Then
                     lRowHeight = tbrThis.ButtonHeight(l)
                  End If
                  lRowWidth = lRowWidth + tbrThis.ButtonWidth(l)
               End If
            End If
         End If
      End If
   Next l
   lHeight = lHeight + lRowHeight
   getVerticalHeight = lHeight
End Function

Public Function getVerticalWidth(tbrThis As cToolbar) As Long
Dim l As Long
Dim lMaxWidth As Long
   For l = 0 To tbrThis.ButtonCount - 1
      If tbrThis.ButtonVisible(l) Then
         If tbrThis.ButtonControl(l) = 0 Then
            If (tbrThis.ButtonWidth(l) > lMaxWidth) Then
               lMaxWidth = tbrThis.ButtonWidth(l)
            End If
         End If
      End If
   Next l
   getVerticalWidth = lMaxWidth
   
End Function

Public Sub newEdit()
   
    util.Hourglass hwnd, True
       
    Dim File As New cFile
    
    File.IdDoc = Files.GetId
    
    File.NewFile
    
    Files.Add File
        
    Call lvwOpeFiles.ListItems.Add(, "k" & File.IdDoc, File.Caption)
    
    lbltot.Caption = CStr(lvwOpeFiles.ListItems.count) & " documents"
    
    Set File = Nothing
    
    util.Hourglass hwnd, False
    
End Sub


Private Sub selecttab(ByVal sKey As String, ByVal Indice As Integer)

    Dim tabx As cTab
    
    If Indice = 1 Then
        tabLeft.Pinned = False
        Set tabx = tabLeft.Tabs.ITem(sKey)
        tabx.Selected = True
        tabLeft.Shown = True
        tabLeft.ScrollRight
        'DoEvents
        tabLeft.Pinned = True
    Else
        tabRight.Pinned = False
        Set tabx = tabRight.Tabs.ITem(sKey)
        
        tabRight.ScrollLeft
        tabRight.Shown = True
        tabx.Selected = True
        'DoEvents
        tabRight.Pinned = True
    End If
        
End Sub
Private Sub ShowPdfHelp(ByVal Index As Integer)

   Dim Archivo As String
   
   If Index = 1 Then
      Archivo = "About_JavaScript_Plus.pdf"
   ElseIf Index = 2 Then
      Archivo = "Analyzing_JavaScript_Files.pdf"
   ElseIf Index = 3 Then
      Archivo = "Help_Panels.pdf"
   ElseIf Index = 4 Then
      Archivo = "Menu_Reference Help.pdf"
   End If
   
   Archivo = util.StripPath(App.Path) & "pdf\" & Archivo
   
   If Not ArchivoExiste2(Archivo) Then
        MsgBox "Archivo : " & Archivo & " doesn't exists", vbCritical
        Exit Sub
   End If
    
   On Error Resume Next
   
   Shell Archivo, vbNormalFocus
    
   Err = 0
   
End Sub

Private Sub UploadFilesToFtp()

    Dim Path As String
    
    Path = util.BrowseFolder(Me.hwnd)
    
    If Len(Path) > 0 Then
        
        Dim afiles() As String
            
        Call get_files_from_folder(Path, afiles)
        
        Dim k As Integer
        Dim fCargo As Boolean
        
        For k = 1 To UBound(afiles)
            If Not fCargo Then
                frmFtpFiles.InicializaArreglo
                fCargo = True
            End If
        
            Call frmFtpFiles.CargaArchivos(afiles(k), VBArchivoSinPath(afiles(k)), k, "", "", "")
        Next k
        
        If fCargo Then
            frmFtpFiles.updmulti = True
            frmFtpFiles.Show vbModal
        End If
    End If
    
End Sub

Private Sub visible_toolbars()

    Dim ItemNumber As Long
    Dim valor As String
    Dim ini As String
    'Dim ctl As Control
    
    ini = IniPath
    
    valor = util.LeeIni(ini, "toolbars", "FILE_visible")
    If valor = "0" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(1)")
        m_cMenu.Checked(ItemNumber) = False
    End If
    
    valor = util.LeeIni(ini, "toolbars", "EDIT_visible")
    If valor = "0" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(2)")
        m_cMenu.Checked(ItemNumber) = False
    End If
    
    valor = util.LeeIni(ini, "toolbars", "JS_visible")
    If valor = "0" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(4)")
        m_cMenu.Checked(ItemNumber) = False
    End If
    
    valor = util.LeeIni(ini, "toolbars", "TOOLS_visible")
    If valor = "0" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(8)")
        m_cMenu.Checked(ItemNumber) = False
    End If
    
    valor = util.LeeIni(ini, "toolbars", "FORMS_visible")
    If valor = "0" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(5)")
        m_cMenu.Checked(ItemNumber) = False
    End If
    
    valor = util.LeeIni(ini, "toolbars", "HTM_visible")
    If valor = "0" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(7)")
        m_cMenu.Checked(ItemNumber) = False
    End If
    
    valor = util.LeeIni(ini, "toolbars", "FORMAT_visible")
    If valor = "0" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(3)")
        m_cMenu.Checked(ItemNumber) = False
    End If
    
    valor = util.LeeIni(ini, "toolbars", "PLUS_visible")
    If valor = "0" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(6)")
        m_cMenu.Checked(ItemNumber) = False
    End If
    
End Sub

Private Sub word_count()

    On Error Resume Next
    
    Dim us As Long
    Dim ut As Long
    'Dim k As Integer
    
    If Not frmMain.ActiveForm Is Nothing Then
        With frmMain.ActiveForm
            us = Len(.txtCode.Text)
            ut = .txtCode.LineCount
            MsgBox "Characters:" & us & Chr(10) & "Lines: " & ut, vbOKOnly + vbInformation
        End With
    End If
    
End Sub


Private Sub cmdFiles_Click()

    Dim frm As Form
    Dim k As Integer
    Dim found As Boolean
    Dim IdDoc As Integer
    
    For k = lvwOpeFiles.ListItems.count To 1 Step -1
    
        found = False
        
        For Each frm In Forms
            If TypeName(frm) = "frmEdit" Then
                If frm.Caption = lvwOpeFiles.ListItems(k).Text Then
                    found = True
                    Exit For
                End If
            End If
        Next
        
        If Not found Then
            IdDoc = CInt(Mid$(lvwOpeFiles.ListItems(k).key, 2))
            Files.Remove IdDoc
            lvwOpeFiles.ListItems.Remove k
        End If
    Next k

    lbltot.Caption = CStr(lvwOpeFiles.ListItems.count) & " documents"
    
End Sub

Private Sub CodeLibrary1_FileSelected(ByVal Archivo As String)

   util.Hourglass hwnd, True
   
   frmMain.newEdit
   
   frmMain.ActiveForm.txtCode.OpenFile Archivo
   
   util.Hourglass hwnd, False
   
End Sub

Private Sub filExp_FileClicked(ByVal File As String)
    
    opeEdit File
        
End Sub

Private Sub HlpExp_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)

    If InStr(itm.Tag, "tutorial") > 0 Then
        frmHelp.url = StripPath(App.Path) & itm.Tag
        Load frmHelp
        frmHelp.Show
    Else
        util.ShellFunc itm.Tag, vbNormalFocus
    End If
    
End Sub


Private Sub jsHlp_ElementClicked(ByVal Value As String)
    If Not ActiveForm Is Nothing Then
        If frmMain.ActiveForm.Name = "frmEdit" Then
            ActiveForm.Insertar Value
            ActiveForm.txtCode.SetFocus
        End If
    End If
End Sub

Private Sub lvwOpeFiles_ItemClick(ByVal ITem As MSComctlLib.ListItem)

    Dim k As Integer
    Dim frm As Form
    
    If Not ITem Is Nothing Then
        util.Hourglass hwnd, True
        
        Dim fvisible As Boolean
        
        fvisible = False
        
        For Each frm In Forms
            If TypeName(frm) = "frmEdit" Then
                If frm.Caption = ITem.Text Then
                    fvisible = True
                    Exit For
                End If
            End If
        Next
                
        If fvisible Then
            For Each frm In Forms
                If TypeName(frm) = "frmEdit" Then
                    If CInt(frm.Tag) = CInt(Mid$(ITem.key, 2)) Then
                        On Error Resume Next
                        Call mdiforms_TabClick(frm.hwnd, 0, 0, 0)
                        frm.SetFocus
                        frm.txtCode.SetFocus
                        Err = 0
                        Exit For
                    End If
                End If
            Next
        Else
            Dim File As cFile
            
            For k = 1 To Files.Files.count
                Set File = New cFile
                Set File = Files.Files.ITem(k)
                If File.Caption = ITem.Text Then
                    File.OpenFile
                    Exit For
                End If
            Next
        End If
        util.Hourglass hwnd, False
    End If
    
End Sub


Public Sub m_cMenu_Click(ItemNumber As Long)
    action_menu m_cMenu, ItemNumber
End Sub
Private Function SaveMacros(ByVal sFileName As String, ByVal nMacroNum As Long) As Boolean
  On Error Resume Next
    Dim bArr() As Byte
    Dim hFile As Integer
    'Dim g As CodeSenseCtl.Globals
    'Set g = New CodeSenseCtl.Globals
    CSGlobals.GetMacro nMacroNum, bArr
    If UBound(bArr) >= 0 Then
        hFile = FreeFile
        On Error Resume Next
        Open sFileName For Binary Access Write As #hFile
          Put #hFile, , bArr
        Close #hFile
        If Err.Number Then
            Exit Function
        End If
        SaveMacros = True
    End If
End Function

Public Sub CreateAlert()

    Dim str As New cStringBuilder
    'Dim Msg As String
    
    If Not ActiveForm Is Nothing Then
        If ActiveForm.Name = "frmEdit" Then
            str.Append "function my_alert()" & vbNewLine
            str.Append "{" & vbNewLine
            str.Append "" & vbTab & "alert(" & Chr$(34) & "Your text" & Chr$(34) & ");" & vbNewLine
            str.Append "}" & vbNewLine & vbNewLine
            
            frmMain.ActiveForm.Insertar str.ToString
        End If
    End If
    
    Set str = Nothing
    
End Sub
Public Sub CreateConfirm()

    Dim str As New cStringBuilder
    Dim Msg As String
    
    If Not ActiveForm Is Nothing Then
        If ActiveForm.Name = "frmEdit" Then
            Msg = InputStr("Confirm:", "New Confirm")
        
            If Len(Msg) > 0 Then
                str.Append "function my_confirm()" & vbNewLine
                str.Append "{" & vbNewLine
                str.Append "" & vbTab & "if (confirm(" & Chr$(34) & Msg & Chr$(34) & "))" & vbNewLine
                str.Append "" & vbTab & "{" & vbNewLine
                str.Append "" & vbTab & "alert(" & Chr$(34) & "do something" & Chr$(34) & ");" & vbNewLine
                str.Append "" & vbTab & "}" & vbNewLine
                str.Append "" & vbTab & "else" & vbNewLine
                str.Append "" & vbTab & "{" & vbNewLine
                str.Append "" & vbTab & "alert(" & Chr$(34) & "do something" & Chr$(34) & ");" & vbNewLine
                str.Append "" & vbTab & "}" & vbNewLine
                str.Append "}" & vbNewLine & vbNewLine
                Call frmMain.ActiveForm.Insertar(str.ToString)
            End If
        End If
    End If
    
    Set str = Nothing

End Sub

Public Sub CreatePrompt()

    Dim str As New cStringBuilder
    Dim Msg As String
    
    If Not ActiveForm Is Nothing Then
        If ActiveForm.Name = "frmEdit" Then
            Msg = InputStr("Prompt:", "New Prompt")
            
            If Len(Msg) > 0 Then
                str.Append "function my_prompt()" & vbNewLine
                str.Append "{" & vbNewLine
                str.Append "    var msg=prompt(" & Chr$(34) & Msg & Chr$(34) & ");" & vbNewLine
                str.Append "    //here the code ..." & vbNewLine
                str.Append "    //here the code ..." & vbNewLine
                str.Append "}" & vbNewLine & vbNewLine
                
                frmMain.ActiveForm.Insertar str.ToString
            End If
        End If
    End If
    
    Set str = Nothing
    
End Sub


Public Sub SingleComent()

    'Dim Indice As Integer
    Dim r As CodeSenseCtl.IRange
    'Dim p As CodeSenseCtl.IPosition
    Dim p As New CodeSenseCtl.Position
    Dim k As Integer
    'Dim buffer As String
    
    If Not frmMain.ActiveForm Is Nothing Then
        If frmMain.ActiveForm.Name = "frmEdit" Then
            With frmMain.ActiveForm
                'get current cursor position
                Set r = .txtCode.GetSel(True)
                Set p = .txtCode.GetSelFromPoint(r.StartLineNo, r.StartColNo)
                
                If Not p Is Nothing Then
                    For k = r.StartLineNo To r.EndLineNo
                        p.LineNo = k
                        p.ColNo = 0
                        Call .txtCode.InsertText("//", p)
                    Next k
                End If
            End With
        End If
    End If
    
End Sub
Public Sub CreateEndLine()

    'Dim Indice As Integer
    Dim r As CodeSenseCtl.IRange
    'Dim p As CodeSenseCtl.IPosition
    Dim p As New CodeSenseCtl.Position
    Dim k As Integer
    
    'get current cursor position
    If Not frmMain.ActiveForm Is Nothing Then
        If frmMain.ActiveForm.Name = "frmEdit" Then
            With frmMain.ActiveForm
                Set r = .txtCode.GetSel(True)
                Set p = .txtCode.GetSelFromPoint(r.StartLineNo, r.StartColNo)
            
                If Not p Is Nothing Then
                    For k = r.StartLineNo To r.EndLineNo
                        p.LineNo = k
                        p.ColNo = Len(.txtCode.GetLine(k))
                        Call .txtCode.InsertText(";", p)
                    Next k
                End If
            End With
        End If
    End If
    
End Sub
Private Sub BlockComment()

    'Dim Indice As Integer
    Dim r As CodeSenseCtl.IRange
    'Dim p As CodeSenseCtl.IPosition
    Dim p As New CodeSenseCtl.Position
    'Dim k As Integer
    
    If Not frmMain.ActiveForm Is Nothing Then
        If frmMain.ActiveForm.Name = "frmEdit" Then
        With frmMain.ActiveForm
            'get current cursor position
            Set r = .txtCode.GetSel(True)
            Set p = .txtCode.GetSelFromPoint(r.StartLineNo, r.StartColNo)
            
            If Not p Is Nothing Then
                If r.StartLineNo <> r.EndLineNo Then
                    p.LineNo = r.StartLineNo - 1
                    p.ColNo = 0
                    Call .txtCode.InsertText("/*" & vbNewLine, p)
                    p.LineNo = r.EndLineNo + 2
                    Call .txtCode.InsertText("*/" & vbNewLine, p)
                Else
                    p.LineNo = r.StartLineNo
                    p.ColNo = 0
                    Call .txtCode.InsertText("/* ", p)
                    p.ColNo = Len(.txtCode.GetLine(p.LineNo))
                    Call .txtCode.InsertText(" */" & vbNewLine, p)
                End If
            End If
        End With
        End If
    End If
    
End Sub
Private Sub InsertBlock()

    Dim str As New cStringBuilder
    
    If Not frmMain.ActiveForm Is Nothing Then
        str.Append "{" & vbNewLine
        str.Append "" & vbNewLine
        str.Append "" & vbNewLine
        str.Append "}" & vbNewLine
        
        If frmMain.ActiveForm.Name = "frmEdit" Then
            Call frmMain.ActiveForm.Insertar(str.ToString)
        End If
    End If
    
End Sub
Private Sub InsertarArchivo()

    Dim glosa As String
    'Dim Indice As Integer
    Dim r As CodeSenseCtl.IRange
    'Dim p As CodeSenseCtl.IPosition
    Dim p As New CodeSenseCtl.Position
    Dim Archivo As String
    Dim LastPath As String
    
    LastPath = App.Path
    
    glosa = strGlosa()
        
    If Not Cdlg.VBGetOpenFileName(Archivo, , , , , , glosa, , , "Insert a file ...", , Me.hwnd) Then
        Exit Sub
    End If

    If frmMain.ActiveForm.Name = "frmEdit" Then
        Set r = frmMain.ActiveForm.txtCode.GetSel(True)
        Set p = frmMain.ActiveForm.txtCode.GetSelFromPoint(r.StartLineNo, r.StartColNo)
                
        If Not p Is Nothing Then
            p.ColNo = r.StartColNo
            p.LineNo = r.StartLineNo
            
            Call frmMain.ActiveForm.txtCode.InsertFile(Archivo, p)
            Call frmMain.ActiveForm.txtCode.SetCaretPos(p.LineNo, p.ColNo + 1)
            Call frmMain.ActiveForm.txtCode.SetFocus
        End If
    End If
    
End Sub






Private Sub HomePageTitle()

    Dim src As New cStringBuilder
    Dim Title As String
    
    Title = InputBox("Page title", "Insert Home Page Title")
    
    If Len(Title) > 0 Then
        src.Append "<title>" & Title & "</title>"
        If frmMain.ActiveForm.Name = "frmEdit" Then
            Call ActiveForm.Insertar(src.ToString)
        End If
    End If
    
    Set src = Nothing
    
End Sub
Private Sub BreakSpace()

    Dim src As New cStringBuilder
    
    src.Append "&nbsp"
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call ActiveForm.Insertar(src.ToString)
    End If
    
    Set src = Nothing
    
End Sub


Public Sub HtmlBody()

    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call ActiveForm.Insertar(LoadSnipet("hbody.txt"))
    End If
    
End Sub
Private Sub PrintPreview()

    On Error GoTo ErrorPrintPreview
        
    If ActiveForm Is Nothing Then
        Exit Sub
    End If
    
    If ActiveForm.Name <> "frmEdit" Then
        Exit Sub
    End If
    
    Call util.Hourglass(hwnd, True)
    
    If util.GetPrinterSettings(Printer.DeviceName, Printer.hdc) = False Then
        MsgBox "An error at try to get printer information.", vbCritical
        Exit Sub
    End If
        
    'pendiente print preview aqui va la nueva dll
    Dim pp As New cPrintPreview
    Dim Archivo As String
    Dim nFreeFile As Long
    
    Archivo = util.StripPath(App.Path) & "ppreview.txt"
    nFreeFile = FreeFile
    
    Open Archivo For Output As #nFreeFile
        If frmMain.ActiveForm.txtCode.SelLength > 0 Then
            Print #1, frmMain.ActiveForm.txtCode.SelText
        Else
            Print #1, frmMain.ActiveForm.txtCode.Text
        End If
    Close #nFreeFile
    
    pp.filename = Archivo
    pp.StartPreview
        
    Exit Sub
ErrorPrintPreview:
    MsgBox "PrintPreview : " & Err & " " & Error$, vbCritical
        
End Sub
Private Sub m_cMenu_ItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
    picStatus.Panels(1).Text = m_cMenu.HelpText(ItemNumber)
End Sub


Private Sub m_cMenu_MenuExit()
    picStatus.Panels(1).Text = "Ready"
End Sub


Private Sub m_cMenuPop_Click(ItemNumber As Long)
    
    Select Case m_cMenuPop.ItemKey(ItemNumber)
        Case "mnuItemTabX(1)"   'New
            Call newEdit
        Case "mnuItemTabX(3)"   'Save
            Call savEdit
        Case "mnuItemTabX(4)"   'Save As
            Call savEdit(True)
        Case "mnuItemTabX(5)"   'Save All
            Call savEdit(False, True)
        Case "mnuItemTabX(7)"   'Close
            Call cloEdit
        Case "mnuItemTabX(8)"   'Close All
            Call cloEdit(True)
    End Select
    
End Sub

Private Sub MarkHlp_ElementClicked(ByVal Value As String)
    If Not ActiveForm Is Nothing Then
        If frmMain.ActiveForm.Name = "frmEdit" Then
            ActiveForm.Insertar Value
        End If
    End If
End Sub

Private Sub MarkHlp_TagClicked(ByVal Value As String)
    If Not ActiveForm Is Nothing Then
        If ActiveForm.Name = "frmEdit" Then
            ActiveForm.Insertar Value
        End If
    End If
End Sub



Private Sub MDIForm_Activate()
    
    If Not fdisplaywelcome Then
        fdisplaywelcome = True
        #If LITE = 1 Then
            If fLoading Then
                Do While fLoading
                    DoEvents
                Loop
            End If
            
            frmInfoAbout.Show vbModal
        #End If
        
        Dim valor
        
        valor = util.LeeIni(IniPath, "reference", "install")
        If valor = "0" Then
            If Confirma("Do you want to install JavaScript Reference, Core Guide and Ajax libraries (Recommended)") = vbYes Then
                frmInstallHelp.Show vbModal
            Else
                MsgBox "You can install JavaScript Reference, Core Guide and Ajax libraries from Help Menu", vbInformation
                util.GrabaIni IniPath, "reference", "install", "1"
            End If
            frmSymbols.Show vbModal
        End If
        Call DesdeLineaComando
        
        Dim quicktip As String
        
        quicktip = util.LeeIni(IniPath, "quicktip", "show")
        
        If quicktip = "1" Then
            frmQuickTip.Show
        End If
    End If
    
    '#If LITE = 0 Then
    '    If Not fchecknewver Then
    '        fchecknewver = True
    '        frmChkUpd.Show vbModal
    '    End If
    '#End If
    
End Sub

Private Sub DesdeLineaComando()
    
    If Len(Command) > 0 Then
    
        If Right$(Command, 3) <> "" Then
            opeEdit Replace(Command, Chr$(34), "")
        Else
            
        End If
    End If
        
End Sub

Private Sub MDIForm_Load()
    
    SaveMru = True
    
    debug_startup "StatusBar ...(1)"
    'crear status bar:
    frmSplash.lblmsg.Caption = "StatusBar ..."
    frmSplash.pgb.Value = 5
    DoEvents
                     
    debug_startup "Menu ...(2)"
    'crear el menu del formulario
    frmSplash.lblmsg.Caption = "Menu ..."
    frmSplash.pgb.Value = 10
    DoEvents
    
    Set m_cMenu = New cPopupMenu
    m_cMenu.hWndOwner = Me.hwnd
    m_cMenu.OfficeXpStyle = True
        
    debug_startup "Imagelist ...(3)"
    'construir las imagenes
    frmSplash.lblmsg.Caption = "Imagelist ..."
    frmSplash.pgb.Value = 15
    DoEvents
    
    buildImageLists
    m_cMenu.ImageList = m_MainImg.hIml
        
    Set m_cMenuPop = New cPopupMenu
    m_cMenuPop.hWndOwner = Me.hwnd
    m_cMenuPop.OfficeXpStyle = True
    m_cMenuPop.ImageList = m_MainImg.hIml
    
    'crear el menu
    debug_startup "Creating Menus ...(4)"
    frmSplash.lblmsg.Caption = "Creating Menus ..."
    frmSplash.pgb.Value = 20
    DoEvents
    createMenu
            
    debug_startup "Toolbars ...(5)"
    frmSplash.lblmsg.Caption = "Toolbars ..."
    frmSplash.pgb.Value = 25
    DoEvents
    
    'setear la menubar
    With tbrMenu
        .DrawStyle = CTBDrawOfficeXPStyle
        .CreateFromMenu2 m_cMenu, CTBMenuStyle, "MENUBAR"
        .Wrappable = False
        .ChevronButton(CTBChevronAdditionalAddorRemove) = False
        .ChevronButton(CTBChevronAdditionalReset) = False
        .ChevronButton(CTBChevronAdditionalCustomise) = False
    End With
        
    'crear las toolbars
    debug_startup "Toolbars ...(6)"
    frmSplash.lblmsg.Caption = "Menu Bars ..."
    frmSplash.pgb.Value = 30
    DoEvents
    
    crea_toolbars
            
    debug_startup "Containers ...(7)"
    frmSplash.lblmsg.Caption = "Containers ..."
    frmSplash.pgb.Value = 35
    DoEvents
    
    'setear los contenedores
    Call activa_toolbars
            
    debug_startup "Help Panels ...(8)"
    frmSplash.lblmsg.Caption = "Help Panels ..."
    frmSplash.pgb.Value = 40
    DoEvents
    
    'configurar panel
    debug_startup "crea_tabs"
    Call crea_tabs
    frmSplash.pgb.Value = 45
    DoEvents
    
    debug_startup "Extras ...(9)"
    frmSplash.lblmsg.Caption = "Extras ..."
    frmSplash.pgb.Value = 50
    DoEvents
    
    debug_startup "filExp.Load"
    filExp.inifile = util.StripPath(App.Path) & "filelist.ini"
    filExp.Load
    
    debug_startup "Markup Help ...(10)"
    frmSplash.lblmsg.Caption = "Markup Help ..."
    frmSplash.pgb.Value = 55
    DoEvents
    
    MarkHlp.inifile = util.StripPath(App.Path) & "config\htmlhelp.ini"
    debug_startup "MarkHlp.Prepare"
    MarkHlp.Prepare
    debug_startup "MarkHlp.Load"
    MarkHlp.Load
    
    'dhtml help
    debug_startup "DHTML Help ...(10.5)"
    frmSplash.lblmsg.Caption = "DHTML Help ..."
    frmSplash.pgb.Value = 60
    DoEvents
    
    vbSDhtml1.inifile = util.StripPath(App.Path) & "config\dhtml.ini"
    debug_startup "vbSDhtml1.Load"
    vbSDhtml1.Load
    
    debug_startup "vbSCSS1.Load"
    frmSplash.lblmsg.Caption = "CSS Help ..."
    frmSplash.pgb.Value = 65
    vbsCSS1.Load
        
    debug_startup "vbsXHTML.Load"
    frmSplash.lblmsg.Caption = "XML Help ..."
    frmSplash.pgb.Value = 70
    vbsXHTML1.Load
    
    debug_startup "JavaScript Help ...(11)"
    frmSplash.lblmsg.Caption = "JavaScript Help ..."
    frmSplash.pgb.Value = 75
    DoEvents
    
    debug_startup "JavaScript Help ...(12)"
    jsHlp.inifile = util.StripPath(App.Path) & "config\jshelp.ini"
    debug_startup "jsHlp.Prepare"
    jsHlp.Prepare
    debug_startup "jsHlp.Load"
    jsHlp.Load
    
    Set tboClp.JScVBALImageList = m_MainImg
    debug_startup "tboClp.Load"
    tboClp.Load
    
    debug_startup "ColPicker1.Load"
    ColPicker1.Load
    frmSplash.lblmsg.Caption = "Color Picker ..."
    frmSplash.pgb.Value = 80
    DoEvents
    
    CodeLibrary1.Prepare
    
    'informacion para intellisense
    debug_startup "Intellisense Help ...(16)"
    frmSplash.lblmsg.Caption = "Intellisense Help ..."
    frmSplash.pgb.Value = 85
    DoEvents
    debug_startup "CargaAyuda"
    Call CargaAyuda
        
    debug_startup "ListaLangs.Load"
    ListaLangs.Load

    debug_startup "Intellisense Help ...(17)"
    frmSplash.lblmsg.Caption = "Edit ..."
    frmSplash.pgb.Value = 85
    DoEvents
    
    'Call CargarColoresEditor
    debug_startup "Colors ...(18)"
    frmSplash.lblmsg.Caption = "Colors ..."
    frmSplash.pgb.Value = 90
    DoEvents
    
    debug_startup "LoadMacros"
    Call LoadMacros
    frmSplash.lblmsg.Caption = "Macros ..."
    frmSplash.pgb.Value = 95
    DoEvents
    
    'fLoading = True
    
    ' Load a new child form, and start
    Dim startup
    
    debug_startup "Util.LeeIni("
    startup = util.LeeIni(IniPath, "startup", "document")
    If startup = "" Then startup = "0"
    
    debug_startup "Macros ...(20)"
    
    If startup = "0" Then
        Dim start_template As String
        
        debug_startup "Macros ...(20-1)"
        start_template = util.LeeIni(IniPath, "startup", "start_template")
        If Len(start_template) > 0 Then
            If ArchivoExiste2(start_template) Then
                opeEdit start_template
            Else
                newEdit
            End If
        Else
            newEdit
        End If
    ElseIf startup = "1" Then
        'abrir los ultimos documentos abiertos
        Dim arr_files() As String
        Dim fileini As String
        
        fileini = util.StripPath(App.Path) & "files.ini"
        debug_startup "Macros ...(20-2)"
        If ArchivoExiste2(fileini) Then
            get_info_section "files", arr_files, fileini
            Dim k As Integer
            
            For k = 1 To UBound(arr_files)
                If ArchivoExiste2(arr_files(k)) Then
                    opeEdit arr_files(k)
                    fLoading = True
                    DoEvents
                End If
            Next k
            fLoading = False
        End If
    End If
    
    'estilo ui
    Dim Style As String
    
    Style = util.LeeIni(IniPath, "style", "type")
    Select Case Style
        Case "XP"
            mdiforms.Style = 0
        Case "2000"
            mdiforms.Style = 1
        Case "2003"
            mdiforms.Style = 2
        Case Else
            mdiforms.Style = 0
    End Select
    
    debug_startup "SetAppHelp hwnd"
    SetAppHelp hwnd
    
    debug_startup "InIDE() (22)"
    If Not InIDE() Then
        debug_startup "RestoreLayout (23)"
        Call RestoreLayout
        DoEvents
        debug_startup "visible_toolbars (24)"
        Call visible_toolbars
    End If
    DoEvents
    
    debug_startup "Finished ...(21)"
    frmSplash.lblmsg.Caption = "Finished ..."
    frmSplash.pgb.Value = 100
    DoEvents
    
    #If LITE = 1 Then
        picStatus.Panels(6).Text = "Unregistered version"
    #Else
        If tipo_version = 1 Then
            picStatus.Panels(6).Text = "Registered version for Single Developer"
        ElseIf tipo_version = 2 Then
            picStatus.Panels(6).Text = "Registered version for commercial purposes"
        ElseIf tipo_version = 3 Then
            picStatus.Panels(6).Text = "Registered version for educational purposes"
        ElseIf tipo_version = 4 Then
            picStatus.Panels(6).Text = "Version for testing purposes"
        End If
    #End If
    
    debug_startup "Starting ..."
    frmSplash.lblmsg.Caption = "Starting JavaScript Plus!. Enjoy ..."
    DoEvents
    
    fLoading = False
    
    Me.OLEDropMode = vbOLEDropManual
    
    debug_startup "Unload frmSplash (23)"
    Unload frmSplash
    DoEvents
    
End Sub
Public Sub DocumentWrite()
    
    Dim src As New cStringBuilder
    Dim r As CodeSenseCtl.IRange
    Dim p As New CodeSenseCtl.Position
    Dim k As Integer
    
    If Not frmMain.ActiveForm Is Nothing Then
        If frmMain.ActiveForm.Name = "frmEdit" Then
            With frmMain.ActiveForm
                Set r = .txtCode.GetSel(True)
                Set p = .txtCode.GetSelFromPoint(r.StartLineNo, r.StartColNo)
                
                If Not p Is Nothing Then
                    For k = r.StartLineNo To r.EndLineNo
                        p.LineNo = k
                        p.ColNo = 0 'Len(.txtCode.GetLine(k))
                        Call .txtCode.InsertText("document.write(" & Chr$(34), p)
                    Next k
                    
                    For k = r.StartLineNo To r.EndLineNo
                        p.LineNo = k
                        p.ColNo = Len(.txtCode.GetLine(k))
                        Call .txtCode.InsertText(Chr$(34) & ");", p)
                    Next k
                End If
                
            End With
        End If
    End If
    
    Set src = Nothing
    
End Sub



Private Sub LoadMacros()
  On Error Resume Next
  Dim s As String
  s = Dir(App.Path & "\macros\")
  Do Until s = ""
    If VBA.Right$(s, 3) = "dem" Then
      AddMacro App.Path & "\macros\" & s, Left(s, InStr(1, s, ".") - 1)
    End If
    s = Dir
  Loop
End Sub

Private Sub AddMacro(File As String, macNum As Long)
  On Error Resume Next
  'Dim p As CodeSenseCtl.Globals
  'Set p = New CodeSenseCtl.Globals
  Dim fFile As Integer, bBar() As Byte
  fFile = FreeFile()
  Open File For Binary Access Read As #fFile
    ReDim bBar(0 To LOF(fFile))
    Get fFile, , bBar
  Close #fFile
  CSGlobals.SetMacro macNum, bBar
End Sub

Private Sub CargaAyuda()

    'cargar objetos
    Dim k As Integer
    Dim j As Integer
    Dim i As Integer
    Dim C As Integer
    'Dim num As Integer
    'Dim num2 As Integer
    Dim glosa As String
    Dim Glosa2 As String
    Dim ini As String
    Dim objarray As New Collection
    Dim icono
    
    Dim sSections() As String
    Dim Atributos() As String
    
    Dim nFreeFile As Long
    nFreeFile = FreeFile
    
    ini = util.StripPath(App.Path) & "config\jshelp.ini"
    get_info_section "objetos", sSections, ini
    
    C = 1
    For k = 2 To UBound(sSections)
        glosa = sSections(k)
        ReDim Preserve udtObjetos(C)
        udtObjetos(C) = glosa
        
        'elementos del objeto
        get_info_section glosa, Atributos, ini
        
        For j = 2 To UBound(Atributos)
            Glosa2 = Atributos(j)
            If InStr(Glosa2, "#") > 0 Then
                icono = util.Explode(Glosa2, 2, "#")
                Glosa2 = util.Explode(Glosa2, 1, "#")
                If InStr(Glosa2, "(") > 0 Then
                    AgregaFuncionJs Glosa2
                    Glosa2 = VBA.Left$(Glosa2, InStr(1, Glosa2, "(") - 1)
                End If
                    
                If icono = "1" Then
                    objarray.Add "P" & Glosa2
                ElseIf icono = "2" Then
                    objarray.Add "F" & Glosa2
                ElseIf icono = "3" Then
                    objarray.Add "E" & Glosa2
                ElseIf icono = "4" Then
                    objarray.Add "C" & Glosa2
                ElseIf icono = "5" Then
                    objarray.Add "X" & Glosa2
                ElseIf icono = "6" Then
                    objarray.Add "O" & Glosa2
                Else
                    objarray.Add "F" & Glosa2
                End If
            End If
        Next j
        
        AddObject glosa, objarray
            
        i = objarray.count
        For j = i To 1 Step -1
            objarray.Remove j
        Next j
        C = C + 1
    Next k
    
    get_info_section "functions", sSections, ini
    For k = 2 To UBound(sSections)
        glosa = util.Explode(sSections(k), 1, "#")
        AgregaFuncionJs glosa
    Next k
        
    'carga ayuda de ajax
    'Call carga_info_ajax
    
    Dim arr_files() As String
    Dim Path As String
    Dim Archivo As String
    
    Path = util.StripPath(App.Path) & "libraries"
    
    get_files_from_folder Path, arr_files
    
    For k = 1 To UBound(arr_files)
        Archivo = arr_files(k)
        glosa = util.LeeIni(Archivo, "information", "active")
        If glosa = "Y" Then
            get_info_section "functions", sSections, Archivo
            For j = 2 To UBound(sSections)
                glosa = util.Explode(sSections(j), 1, "#")
                AgregaFuncionJs glosa
            Next j
        End If
    Next k
    
End Sub

Private Sub UnloadAll()
    Dim i As Integer
    
    ' if there is any pending opened form
    ' just unload them all
    On Error Resume Next
    For i = 0 To Forms.count - 1
        Unload Forms(i)
    Next
End Sub
Private Sub RestoreLayout()

    On Error GoTo ErrorRestoreLayout
    
    Dim sFile As String
    
    sFile = util.StripPath(App.Path) & "layout.xml"
        
    If Not ArchivoExiste2(sFile) Then
        activa_toolbars_dock vbalDockContainer1, "FILE", True
        activa_toolbars_dock vbalDockContainer1, "EDIT", True
        activa_toolbars_dock vbalDockContainer1, "FORMAT", True
        activa_toolbars_dock vbalDockContainer1, "JS", True
        activa_toolbars_dock vbalDockContainer1, "FORMS", True
        activa_toolbars_dock vbalDockContainer1, "PLUS", True
        activa_toolbars_dock vbalDockContainer1, "HTML", True
        activa_toolbars_dock vbalDockContainer1, "TOOLS", True
        Exit Sub
    End If
    
    Dim iFile As Integer
    iFile = FreeFile
    Open sFile For Binary Access Read As #iFile
        Dim sXml As String
        sXml = String(LOF(iFile), " ")
        Get #iFile, , sXml
    Close #iFile
    
    ReDim sKey(1 To 9) As String
    Dim hwnd(1 To 9) As Long
    sKey(1) = "MENU": hwnd(1) = tbrMenu.hwnd
    sKey(2) = "FILE": hwnd(2) = tbrFile.hwnd
    sKey(3) = "EDIT": hwnd(3) = tbrEdit.hwnd
    sKey(4) = "JS": hwnd(4) = tbrJs.hwnd
    sKey(5) = "FORMS": hwnd(5) = tbrForms.hwnd
    sKey(6) = "FORMAT": hwnd(6) = tbrFormat.hwnd
    sKey(7) = "PLUS": hwnd(7) = tbrPlus.hwnd
    sKey(8) = "TOOLS": hwnd(8) = tbrTools.hwnd
    sKey(9) = "HTM": hwnd(9) = tbrHtm.hwnd
    
    Dim ctl As Control
    Dim ctlDock As vbalDockContainer
    For Each ctl In Me.Controls
       If TypeName(ctl) = "vbalDockContainer" Then
          Set ctlDock = ctl
          ctlDock.RestoreLayout sXml, sKey(), hwnd()
       End If
    Next
    
    Exit Sub
ErrorRestoreLayout:
    MsgBox "RestoreLayout : " & Err & " " & Error$, vbCritical
    
End Sub


Private Sub SaveLayout()

    Dim sMsg As String
    Dim ctl As Control
    Dim ctlDock As vbalDockContainer
    sMsg = "<Layout>"
    
    For Each ctl In Me.Controls
        If TypeName(ctl) = "vbalDockContainer" Then
            Set ctlDock = ctl
            sMsg = sMsg & vbCrLf & ctlDock.SaveLayout()
        End If
    Next
    sMsg = sMsg & vbCrLf & "</Layout>"
   
    'IndicateClear
    'Indicate sMsg
   
    Dim sFile As String
    sFile = util.StripPath(App.Path) & "layout.xml"
    On Error Resume Next
    util.BorrarArchivo (sFile)
    On Error GoTo ErrorHandler
   
    Dim iFile As Integer
    iFile = FreeFile
    Open sFile For Binary Access Write As #iFile
    Put #iFile, , sMsg
    Close #iFile
    Exit Sub
   
ErrorHandler:
   MsgBox "Failed to save layout: " & Err.description, vbExclamation
   Close #iFile
   Exit Sub
End Sub


Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo salir
    
    Dim k As Integer
    
    For k = 1 To Data.Files.count
        opeEdit Data.Files(k)
    Next k
    
   Exit Sub
   
salir:
   Err = 0
   
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error Resume Next
    
    If fLoading Then Exit Sub
    
    If Not fexpired Then
        Dim ask As String
        ask = util.LeeIni(IniPath, "startup", "ask")
    
        If ask = "1" Or ask = "" Then
            If MsgBox("Are you sure to quit?", vbQuestion + vbYesNoCancel + vbDefaultButton3) <> vbYes Then
                Cancel = 1
                Exit Sub
            End If
        End If
    End If
    
    'guardar lista de ultimos archivos utilizados
    Dim k As Integer
    Dim File As cFile
    Dim Archivo As String
    Dim j As Integer
    
    Archivo = util.StripPath(App.Path) & "files.ini"
        
    Call util.BorrarArchivo(Archivo)
    
    j = 1
    With Files
        For k = 1 To .Files.count
            Set File = New cFile
            Set File = .Files.ITem(k)
            If Not File.Ftp Then
                If Len(File.filename) > 0 Then
                    Call util.GrabaIni(Archivo, "files", "file" & j, File.filename)
                    j = j + 1
                End If
            End If
            Set File = Nothing
        Next k
    End With
    
    'Load frmWait
    'frmWait.Caption = "Closing JavaScript Plus!"
    'frmWait.lbl(0).Caption = "Please wait while JavaScript Plus! closes opened files...."
    'frmWait.Show
    'DoEvents
    
    'cerrar todos los archivos
    If Not Files.CloseAll() Then
        Cancel = 1
        Exit Sub
    End If
    
    'borrar archivos temporales
    borrar_archivos_tmp
    
    'actualizar lista de ultimos archivos utilizados
    Dim cR As New cRegistry
    cR.ClassKey = HKEY_CURRENT_USER
    cR.SectionKey = "Software\vbsoftware\MRU"
    
    If SaveMru Then
        m_cMRU.Save cR
    End If
    
    'guardar layout de toolbars
    If Not InIDE() Then
        Call SaveLayout
    End If
    
    'limpiar imagelist
    m_MainImg.Destroy

    'limpiar las toolbar
    Dim ctl As Control
    
    For Each ctl In Me.Controls
        If TypeName(ctl) = "cToolbar" Then
            ctl.DestroyToolBar
        End If
    Next
    
    If glbquickon Then
        util.GrabaIni IniPath, "quicktip", "show", "1"
    Else
        util.GrabaIni IniPath, "quicktip", "show", "0"
    End If
    
    QuitHelp
    
    'cerrar todos los archivos abiertos
    UnloadAll
    
    'desbloquear acceso unico
    #If LITE = 1 Then
        Call util.MutexCleanUp
    #End If
    
    Err = 0
    
End Sub

Private Sub MDIForm_Resize()

    LockWindowUpdate hwnd
    LockWindowUpdate False
    
End Sub

Private Sub MDIForm_Terminate()
    'If Forms.Count = 0 Then
        UnloadApp
    'End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'    End
End Sub

Private Sub mdiforms_TabClick(TabHwnd As Long, Button As Integer, x As Long, y As Long)
        
    Dim frm As Form
    Dim File As New cFile
    Dim k As Integer
    
    If Button = vbRightButton Then
        Dim pt As POINTAPI
        
        GetCursorPos pt
        
        With m_cMenuPop
            .Restore "TABPOPUP"
            Call .ShowPopupAbsolute(pt.x, pt.y)
        End With
        Exit Sub
    Else
        For Each frm In Forms
            If TypeName(frm) = "frmEdit" Then
                If frm.hwnd = TabHwnd Then
                    For k = 1 To Files.Files.count
                        Set File = New cFile
                        Set File = Files.Files.ITem(k)
                        If File.Caption = Replace(frm.Caption, "*", "") Then
                            If Len(File.filename) > 0 Then
                                If InStr(File.filename, "\") Then
                                    picStatus.Panels(1).Text = util.PathArchivo(File.filename)
                                Else
                                    If Right$(File.RemoteFolder, 1) <> "/" Then
                                        picStatus.Panels(1).Text = File.RemoteFolder & "/" & File.filename
                                    Else
                                        picStatus.Panels(1).Text = File.RemoteFolder & File.filename
                                    End If
                                End If
                                picStatus.Panels(1).Tag = File.filename
                            Else
                                picStatus.Panels(1).Text = vbNullString
                                picStatus.Panels(1).Tag = vbNullString
                            End If
                            GoTo salir
                        End If
                        Set File = Nothing
                    Next k
                End If
            End If
        Next
    End If
    
salir:
End Sub



Private Sub picRight_Resize(Index As Integer)

    If Index = 0 Then
        HlpExp.Move 0, 0, picRight(Index).Width, picRight(Index).Height
    ElseIf Index = 1 Then
        ColPicker1.Move 0, 0, picRight(Index).Width, picRight(Index).Height
    ElseIf Index = 2 Then
        tboClp.Move 0, 0, picRight(Index).Width, picRight(Index).Height
    ElseIf Index = 3 Then
        'HlpHttpCodes.Move 0, 0, picRight(Index).Width, picRight(Index).Height
    ElseIf Index = 4 Then
        CodeLibrary1.Move 0, 0, picRight(Index).Width, picRight(Index).Height
    End If
    
End Sub


Private Sub picSizeLeft_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    picSizeLeft.BackColor = &H8000000C
End Sub


Private Sub picSizeLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        tabLeft.Width = picSizeLeft.Left + x
    End If
End Sub


Private Sub picSizeLeft_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picSizeLeft.BackColor = &H8000000F
End Sub

Private Sub picSizeRight_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    picSizeRight.BackColor = &H8000000C
End Sub


Private Sub picSizeRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        tabRight.Width = Me.Width - (picSizeRight.Left + x)
    End If
End Sub


Private Sub picSizeRight_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picSizeRight.BackColor = &H8000000F
End Sub

Private Sub picTab_Resize(Index As Integer)

    On Error Resume Next
    
    'Dim Top As Long
    
    If Index = 0 Then       'file explorer
        filExp.Move 0, 0, picTab(Index).Width, picTab(Index).Height
    ElseIf Index = 1 Then   'javascript explorer
        jsHlp.Move 0, 0, picTab(Index).Width, picTab(Index).Height
    ElseIf Index = 2 Then   'help explorer
        MarkHlp.Move 0, 0, picTab(Index).Width, picTab(Index).Height
    ElseIf Index = 3 Then   'cliboard explorer
        vbsXHTML1.Move 0, 0, picTab(Index).Width, picTab(Index).Height
    ElseIf Index = 4 Then   'dhtml help
        vbSDhtml1.Move 0, 0, picTab(Index).Width, picTab(Index).Height
    ElseIf Index = 5 Then   'css help
        vbsCSS1.Move 0, 0, picTab(Index).Width, picTab(Index).Height
    ElseIf Index = 6 Then   'file manager
        picTbFiles.Move 0, 0, picTab(Index).Width
        lvwOpeFiles.Move 0, picTbFiles.Height + 1, picTab(Index).Width, picTab(Index).Height - picTbFiles.Height
    End If
        
    Err = 0
    
End Sub



Private Sub tabLeft_Pinned()
    picSizeLeft.Visible = True
End Sub

Private Sub tabLeft_TabClick(theTab As vbalDTab6.cTab, ByVal iButton As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
'    MsgBox "tabLeft_TabClick"
End Sub

Private Sub tabLeft_TabSelected(theTab As vbalDTab6.cTab)
'    MsgBox "tabLeft_TabSelected"
End Sub


Private Sub tabLeft_UnPinned()
    picSizeLeft.Visible = False
End Sub

Private Sub tabRight_Pinned()
    picSizeRight.Visible = True
End Sub

Private Sub tabRight_TabClick(theTab As vbalDTab6.cTab, ByVal iButton As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
    tabrightsel = theTab.Index
End Sub


Private Sub tabRight_TabSelected(theTab As vbalDTab6.cTab)
    tabrightsel = theTab.Index
End Sub

Private Sub tabRight_UnPinned()
    picSizeRight.Visible = False
End Sub

Private Sub tbrEdit_ButtonClick(ByVal lButton As Long)

    Select Case tbrEdit.ButtonKey(lButton)
        Case "EDIT:CUT"
            Call edtOpe(3)
        Case "EDIT:COPY"
            Call edtOpe(4)
        Case "EDIT:PASTE"
            Call edtOpe(5)
        Case "EDIT:FIND"
            Call edtOpe(13)
        Case "EDIT:FIND:PREV"
            Call edtOpe(15)
        Case "EDIT:FIND:NEXT"
            Call edtOpe(14)
        Case "EDIT:REPLACE"
            Call edtOpe(16)
        Case "EDIT:UNDO"
            Call edtOpe(1)
        Case "EDIT:REDO"
            Call edtOpe(2)
        Case "EDIT:INDENT"
            Call edtOpe(8)
        Case "EDIT:OUTDENT"
            Call edtOpe(9)
    End Select
    
End Sub

Private Sub tbrFile_ButtonClick(ByVal lButton As Long)

    Select Case tbrFile.ButtonKey(lButton)
        Case "FILE:NEW"
            'Call newEdit
            frmNewDoc.Show vbModal
        Case "FILE:OPEN"
            Call opeEdit
        Case "FILE:OPEN:FTP"
            frmFtpFiles.Show vbModal
            'frmSites.tipo_conexion = 0
            'frmSites.Show vbModal
        Case "FILE:OPEN:WEB"
            'open from web
            frmOpeWeb.Show vbModal
                        
        Case "FILE:OPEN:FOLDER"
            'open folder
            'Call open_folder
            frmOpenFolder.Show vbModal
        Case "FILE:SAVE"
            Call savEdit
        Case "FILE:SAVEALL"
            Call savEdit(False, True)
        Case "FILE:PRINT"
            Call prnEdit
        Case "FILE:PREVIOUS"
        
    End Select
    
End Sub



Private Sub tbrFormat_ButtonClick(ByVal lButton As Long)

    Select Case tbrFormat.ButtonKey(lButton)
        Case "FORMAT:FONT"
            'font
            If Not ActiveForm Is Nothing Then
                frmFont.Show vbModal
            End If
        Case "FORMAT:FPARAGRAPH"
            'format paragraph
            If Not ActiveForm Is Nothing Then
                frmFParagraph.Show vbModal
            End If
        Case "FORMAT:NUMBERED"
            'numbered
            If Not ActiveForm Is Nothing Then
                frmListType.mycaption = "Ordered List"
                frmListType.mytype = 1
                frmListType.Show vbModal
            End If
        Case "FORMAT:BULLETED"
            'bulleted
            If Not ActiveForm Is Nothing Then
                frmListType.mycaption = "Unordered List"
                frmListType.mytype = 1
                frmListType.Show vbModal
            End If
        Case "FORMAT:BIGTEXT"
            'big
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<big></big>")
                End If
            End If
        Case "FORMAT:SMALLTEXT"
            'small
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<small></small>")
                End If
            End If
        Case "FORMAT:HEADING:1"
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<h1></h1>")
                End If
            End If
        Case "FORMAT:HEADING:2"
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<h2></h2>")
                End If
            End If
        Case "FORMAT:HEADING:3"
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<h3></h3>")
                End If
            End If
        Case "FORMAT:HEADING:4"
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<h4></h4>")
                End If
            End If
        Case "FORMAT:HEADING:5"
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<h5></h5>")
                End If
            End If
        Case "FORMAT:HEADING:6"
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<h6></h6>")
                End If
            End If
        Case "FORMAT:BOLD"
            'bold
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<b></b>")
                End If
            End If
        Case "FORMAT:ITALIC"
            'bold
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<i></i>")
                End If
            End If
        Case "FORMAT:UNDERLINE"
            'bold
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<u></u>")
                End If
            End If
        Case "FORMAT:PARAGRAPH"
            'parrafo
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<p></p>")
                End If
            End If
        Case "FORMAT:ALEFT"
            'left
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<p align=""left""></p>")
                End If
            End If
        Case "FORMAT:ACENTER"
            'center
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<center></center>")
                End If
            End If
        Case "FORMAT:ARIGHT"
            'right
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<p align=""right""></p>")
                End If
            End If
        Case "FORMAT:JUSTIFY"
            'justify
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<p align=""justify""></p>")
                End If
            End If
    End Select
            
End Sub

Private Sub tbrForms_ButtonClick(ByVal lButton As Long)

    Select Case tbrForms.ButtonKey(lButton)
        Case "FORMS:FORM"
            'forms
            If Not ActiveForm Is Nothing Then
                frmInsForm.Show vbModal
            End If
        Case "FORMS:CHECKBOX"
            'checkbox
            If Not ActiveForm Is Nothing Then
                frmHtmlCheck.tipo_control = "checkbox"
                frmHtmlCheck.Show vbModal
            End If
        Case "FORMS:COMBOBOX"
            'combobox
            If Not ActiveForm Is Nothing Then
                frmHtmlCombo.Show vbModal
            End If
        Case "FORMS:FILEATTACH"
            'file attach
            If Not ActiveForm Is Nothing Then
                frmHtmlFileAttach.Show vbModal
            End If
        Case "FORMS:HIDDEN"
            'hidden entry
            If Not ActiveForm Is Nothing Then
                frmHtmlHidden.Show vbModal
            End If
        Case "FORMS:LISTBOX"
            'listbox
            If Not ActiveForm Is Nothing Then
                frmHtmlListbox.Show vbModal
            End If
        Case "FORMS:PASSWORD"
            If Not ActiveForm Is Nothing Then
                frmHtmlText.tipo_texto = "password"
                frmHtmlText.Show vbModal
            End If
        Case "FORMS:RADIO"
            'radio button
            If Not ActiveForm Is Nothing Then
                frmHtmlCheck.tipo_control = "radio"
                frmHtmlCheck.Show vbModal
            End If
        Case "FORMS:TEXT"
            'textbox
            If Not ActiveForm Is Nothing Then
                frmHtmlText.tipo_texto = "text"
                frmHtmlText.Show vbModal
            End If
        Case "FORMS:TEXTAREA"
            'textarea
            If Not ActiveForm Is Nothing Then
                frmHtmlTextArea.Show vbModal
            End If
    End Select
            
End Sub

Private Sub tbrHtm_ButtonClick(ByVal lButton As Long)

    Select Case tbrHtm.ButtonKey(lButton)
        Case "HTM:HYPERLINK"
            'hyperlink
            If Not ActiveForm Is Nothing Then
                frmHyperlink.Show vbModal
                'Call Hiperlink
            End If
        Case "HTM:IMAGE"
            'image
            If Not ActiveForm Is Nothing Then
                frmImage.Show vbModal
            End If
        Case "HTM:HORIZONTAL"
            'horizontal line/ruler
            If Not ActiveForm Is Nothing Then
                frmRuler.Show vbModal
            End If
        Case "HTM:COMMENT"
            If Not ActiveForm Is Nothing Then
                If ActiveForm.Name = "frmEdit" Then
                    Call ActiveForm.Insertar("<!-- -->")
                End If
            End If
        Case "HTM:TABLE"
            'tabla
            If Not ActiveForm Is Nothing Then
                frmTabla.Show vbModal
            End If
        Case "HTM:FRAME"
            'frameset
            If Not ActiveForm Is Nothing Then
                frmFramesWiz.Show vbModal
            End If
        Case "HTM:SCRIPT"
            'script
            If Not ActiveForm Is Nothing Then
                frmScript.Show vbModal
            End If
        Case "HTM:SYMBOL"
            'symbol
            If Not ActiveForm Is Nothing Then
                frmCharExp.inifile = util.StripPath(App.Path) & "config\htmlmap.ini"
                
                frmCharExp.Show vbModal
            End If
    End Select
    
End Sub

Private Sub tbrJs_ButtonClick(ByVal lButton As Long)

    'Dim archivo As String
    
    Select Case tbrJs.ButtonKey(lButton)
        'Case "JS:FAVORITES"
        '    Archivo = util.StripPath(App.path) & "help\javascript.chm"
            
        '    If Not ArchivoExiste2(Archivo) Then
        '        MsgBox "File " & Archivo & " not found!", vbCritical
        '        Exit Sub
        '    End If
            
        '    util.ShellFunc Archivo, vbNormalFocus
        Case "JS:OBJECT"
            frmObjExa.Show vbModal
        Case "JS:LIBRARY"
            frmLibraryManager.Show vbModal
        Case "JS:CONFIRM"
            Call CreateConfirm
        Case "JS:PROMPT"
            Call CreatePrompt
        Case "JS:ALERT"
            Call CreateAlert
        Case "JS:ARRAY"
            If Not frmMain.ActiveForm Is Nothing Then
                frmArray.Show vbModal
            End If
        Case "JS:STATEMENTS"
            If Not frmMain.ActiveForm Is Nothing Then
                frmStatements.Show vbModal
            End If
        Case "JS:ESCAPE"
            If Not frmMain.ActiveForm Is Nothing Then
                frmEscChar.Show vbModal
            End If
        Case "JS:REGEXP"
            If Not frmMain.ActiveForm Is Nothing Then
                frmRegExp.Show vbModal
            End If
        Case "JS:PRETEMPLATE"
            'If Not frmMain.ActiveForm Is Nothing Then
            '    If frmMain.ActiveForm.Name = "frmEdit" Then
            '        frmPreTemplates.Show vbModal
            '    End If
            'End If
        Case "JS:USERTEMPLATE"
            'If Not frmMain.ActiveForm Is Nothing Then
            '    If frmMain.ActiveForm.Name = "frmEdit" Then
            '        frmUserTemplate.Show vbModal
            '    End If
            'End If
        Case "JS:BLOCK"
            Call InsertBlock
        Case "JS:COMMEN1"
            If Not ActiveForm Is Nothing Then
                Call SingleComent
            End If
        Case "JS:COMMEN2"
            Call BlockComment
        Case "JS:LINEEND"
            Call CreateEndLine
        Case "JS:FUNCTION"
            If Not ActiveForm Is Nothing Then
                Dim funcion As String
                funcion = InputBox("Function Name:", "New Function")
                If Len(Trim$(funcion)) > 0 Then
                    frmMain.ActiveForm.Insertar "function " & funcion & "()" & vbNewLine & "{" & vbNewLine & vbNewLine & "}"
                End If
            End If
    End Select
            
End Sub

Private Sub tbrPlus_ButtonClick(ByVal lButton As Long)

    Select Case tbrPlus.ButtonKey(lButton)
        Case "PLUS:FAVORITES"
            'add to favorites
            If Not ActiveForm Is Nothing Then
                frmAddFavorites.Show vbModal
            End If
        Case "PLUS:CALENDAR"
            'calendar
            If Not ActiveForm Is Nothing Then
                frmCalendar.Show vbModal
            End If
        Case "PLUS:SLIDESHOW"
            'slideshow
            If Not ActiveForm Is Nothing Then
                frmSlideShow.Show vbModal
            End If
        Case "PLUS:COUNTRIES"
            'countries menu
            If Not ActiveForm Is Nothing Then
                frmCountryMenus.Show vbModal
            End If
        Case "PLUS:DROPDOWN"
            'drop down menu
            If Not ActiveForm Is Nothing Then
                frmDropDownMenu.Show vbModal
            End If
        Case "PLUS:EMAIL"
            'email link
            If Not ActiveForm Is Nothing Then
                frmCreateEmail.Show vbModal
            End If
        Case "PLUS:IFRAME"
            'iframe
            If Not ActiveForm Is Nothing Then
                frmIframe.Show vbModal
            End If
        Case "PLUS:ROLLOVER"
            'image rollover
            If Not ActiveForm Is Nothing Then
                frmRollover.Show vbModal
            End If
        Case "PLUS:LASTDATE"
            'last date
            If Not ActiveForm Is Nothing Then
                frmLastModDate.Show vbModal
            End If
        Case "PLUS:LEFTMENU"
            'left menu
            If Not ActiveForm Is Nothing Then
                frmLeftMenu.Show vbModal
            End If
        Case "PLUS:METATAG"
            'metatag
            If Not ActiveForm Is Nothing Then
                frmMetaTag.Show vbModal
            End If
        Case "PLUS:PAGETRAN"
            'page tran
            If Not ActiveForm Is Nothing Then
                frmPageTran.Show vbModal
            End If
        Case "PLUS:POPUPWINDOW"
            'popup
            If Not ActiveForm Is Nothing Then
                frmPopup.Show vbModal
            End If
        Case "PLUS:COLSCROLL"
            'colored scrollbar
            If Not ActiveForm Is Nothing Then
                frmCreateColScrollbar.Show vbModal
            End If
        Case "PLUS:MOUSEOVER"
            'mouse over
            If Not ActiveForm Is Nothing Then
                frmMouseOverLinks.Show vbModal
            End If
        Case "PLUS:POPUPMENU"
            'mouse over
            If Not ActiveForm Is Nothing Then
                frmPopupMenu.Show vbModal
            End If
        Case "PLUS:TABMENU"
            'mouse over
            If Not ActiveForm Is Nothing Then
                frmTabMenu.Show vbModal
            End If
        Case "PLUS:TREEMENU"
            'mouse over
            If Not ActiveForm Is Nothing Then
                frmTreeMenu.Show vbModal
            End If
    End Select
    
End Sub

Public Sub check_properties()

    If Not ActiveForm Is Nothing Then
        Call get_word
    End If
    
End Sub
Private Sub get_word()

    Dim r As CodeSenseCtl.IRange
    'Dim p As CodeSenseCtl.IPosition
    Dim p As New CodeSenseCtl.Position
    Dim linea As String
    'Dim lineawrk As String
    
    Dim attrib As String
    'Dim attrib2 As String
    Dim k As Integer
    Dim j As Integer
    'Dim pos As Integer
    'Dim ultpos As Integer
    Dim tagaux As String
    'Dim valor As String
    'Dim prop As String
    
    With ActiveForm
        Set r = .txtCode.GetSel(False)
        
        Set p = .txtCode.GetSelFromPoint(r.StartLineNo, r.StartColNo)
        
        'If p Is Nothing Then
        '    Exit Sub
        'End If
        
        If r.StartLineNo = r.EndLineNo Then
            If r.StartColNo = 0 Then
                linea = .txtCode.GetLine(r.StartLineNo)
            Else
                linea = .txtCode.GetLine(r.StartLineNo)
            End If
                                                                            
            If Len(Trim$(linea)) = 0 Then
                frmMain.ActiveForm.tagaux = vbNullString
                tagaux = vbNullString
                Exit Sub
            End If
            tagaux = Trim$(Left$(linea, r.StartColNo))
            
            If VBA.Right$(tagaux, 1) = ">" Then
                tagaux = vbNullString
                GoTo seguir
            End If
            
            For k = Len(tagaux) To 1 Step -1
                If Mid$(tagaux, k, 1) = "<" Then
                    tagaux = Mid$(tagaux, k + 1)
                    If InStr(tagaux, " ") Then
                        attrib = Mid$(tagaux, InStr(tagaux, " "))
                        tagaux = Left$(tagaux, InStr(tagaux, " ") - 1)
                        If InStr(tagaux, ">") Then
                            tagaux = Left$(tagaux, InStr(tagaux, ">") - 1)
                        End If
                        'buscar los atributos desde el inicio hasta el fin de etiqueta
                        For j = r.StartColNo To Len(linea)
                            If Mid$(linea, j, 1) = ">" Then
                                If Len(linea) < r.StartColNo Then
                                    attrib2 = tagaux & attrib & Mid$(linea, r.StartColNo + 1, j - r.StartColNo - 1)
                                Else
                                    attrib2 = tagaux & attrib
                                End If
                                Exit For
                            End If
                        Next j
                        
                        If Len(attrib2) = 0 Then
                            attrib2 = tagaux & attrib '& Mid$(linea, 1, j)
                        End If
                    End If
                    Exit For
                End If
            Next k
            linea = Mid$(linea, r.StartColNo)
            
        End If
        
seguir:
        p.ColNo = r.StartColNo
        p.LineNo = r.StartLineNo
        .txtCode.SetFocus
        Call .txtCode.SetCaretPos(p.LineNo, p.ColNo)
    End With
    ActiveForm.tagaux = tagaux
    floading_prop = False
End Sub



Private Sub tbrTools_ButtonClick(ByVal lButton As Long)

    Dim Archivo As String
    
    Select Case tbrTools.ButtonKey(lButton)
        Case "TOOLS:TIDYRUN"
            HTidy.ExecuteDefault
        Case "TOOLS:COLOR"
            If Not frmMain.ActiveForm Is Nothing Then
                frmColorBrowser.Show vbModal
            End If
        Case "TOOLS:IMAGE"
            Dim ib As New cImageBrowser
            ib.StartBrowse
            Set ib = Nothing
        Case "TOOLS:JSLINT"
            Call ejecuta_jslint
        Case "TOOLS:DATABASE"
            ejecuta_query_studio
        Case "TOOLS:PLUGINS"
            If Not ActiveForm Is Nothing Then
                frmPlugMan.Show vbModal
            End If
        Case "TOOLS:XMLEXPLORER"
            'xml explorer
            Dim glosaxml As String
                        
            glosaxml = "Xml Files (*.xml)|*.xml|"
            glosaxml = glosaxml & "All Files (*.*)|*.*"
        
            If Cdlg.VBGetOpenFileName(Archivo, , , , , , glosaxml, , LastPath, , "XML", Me.hwnd) Then
                Dim xmlexp As New cXmlExplorer
                xmlexp.filename = Archivo
                xmlexp.LangPath = util.StripPath(App.Path) & "languages\"
                xmlexp.StartExplorer
            End If
        Case "TOOLS:DOS"
            Dim cons As New cDOS
            cons.StartConsole
            Set cons = Nothing
        Case "TOOLS:IMAGEEFFECT"
            'image editor
            Call ejecuta_image_editor
        Case "TOOLS:ICONEDITOR"
            ExecuteTool ("IconEditor.exe")
    End Select

End Sub

Private Sub vbalDockContainer1_BarClose(ByVal sKey As String, bCancel As Boolean)
    
    On Error Resume Next
    
    Dim ItemNumber As Long
    
    If sKey = "FILE" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(1)")
        m_cMenu.Checked(ItemNumber) = False
    ElseIf sKey = "EDIT" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(2)")
        m_cMenu.Checked(ItemNumber) = False
    ElseIf sKey = "JS" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(4)")
        m_cMenu.Checked(ItemNumber) = False
    ElseIf sKey = "TOOLS" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(8)")
        m_cMenu.Checked(ItemNumber) = False
    ElseIf sKey = "FORMS" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(5)")
        m_cMenu.Checked(ItemNumber) = False
    ElseIf sKey = "HTM" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(7)")
        m_cMenu.Checked(ItemNumber) = False
    ElseIf sKey = "FORMAT" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(3)")
        m_cMenu.Checked(ItemNumber) = False
    ElseIf sKey = "PLUS" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(6)")
        m_cMenu.Checked(ItemNumber) = False
    End If
    util.GrabaIni IniPath, "toolbars", sKey & "_visible", "0"
    
    Err = 0
    
End Sub

Private Sub vbalDockContainer1_Docked(ByVal key As String)
    On Error Resume Next
    util.GrabaIni IniPath, "toolbars", key & "_visible", "1"
    Err = 0
End Sub

Private Sub vbalDockContainer2_BarClose(ByVal sKey As String, bCancel As Boolean)

    On Error Resume Next
    
    Dim ItemNumber As Long
    
    If sKey = "FORMS" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(5)")
        m_cMenu.Checked(ItemNumber) = False
    ElseIf sKey = "HTM" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(7)")
        m_cMenu.Checked(ItemNumber) = False
    End If
    
    util.GrabaIni IniPath, "toolbars", sKey & "_visible", "0"
    
    Err = 0
    
End Sub

Private Sub vbalDockContainer2_ChevronPress(ByVal key As String, ByVal x As Long, ByVal y As Long)
    chevronPress vbalDockContainer2, key, x, y
End Sub

Private Sub vbalDockContainer2_Docked(ByVal key As String)
    
    On Error Resume Next
    util.GrabaIni IniPath, "toolbars", key & "_docked", "1"
    Err = 0
    
End Sub

Private Sub vbalDockContainer2_SizeChanged()
    picStatus.Move 0, vbalDockContainer2.Height - vbalDockContainer2.NonDockingAreaSize * Screen.TwipsPerPixelY, vbalDockContainer2.Width, vbalDockContainer2.NonDockingAreaSize * Screen.TwipsPerPixelY
End Sub


Private Sub vbalDockContainer2_Undocked(ByVal key As String)

    On Error Resume Next
    util.GrabaIni IniPath, "toolbars", key & "_docked", "0"
    Err = 0
    
End Sub

Private Sub vbalDockContainer3_BarClose(ByVal sKey As String, bCancel As Boolean)

    On Error Resume Next
    
    Dim ItemNumber As Long
    
    If sKey = "FORMAT" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(3)")
        m_cMenu.Checked(ItemNumber) = False
    End If
        
    util.GrabaIni IniPath, "toolbars", sKey & "_visible", "0"
    
    Err = 0
    
End Sub

Private Sub vbalDockContainer3_Docked(ByVal key As String)

    On Error Resume Next
    util.GrabaIni IniPath, "toolbars", key & "_docked", "1"
    Err = 0
    
End Sub

Private Sub vbalDockContainer3_Undocked(ByVal key As String)
    On Error Resume Next
    util.GrabaIni IniPath, "toolbars", key & "_docked", "0"
    Err = 0
End Sub


Private Sub vbalDockContainer4_BarClose(ByVal sKey As String, bCancel As Boolean)

    On Error Resume Next
    
    Dim ItemNumber As Long
    
    If sKey = "PLUS" Then
        ItemNumber = frmMain.m_cMenu.IndexForKey("mnuView_Toolbars(6)")
        m_cMenu.Checked(ItemNumber) = False
    End If
    
    util.GrabaIni IniPath, "toolbars", sKey & "_visible", "0"
    
    Err = 0
    
End Sub

Private Sub vbalDockContainer4_Docked(ByVal key As String)
    On Error Resume Next
    util.GrabaIni IniPath, "toolbars", key & "_docked", "1"
    Err = 0
End Sub


Private Sub vbalDockContainer4_Undocked(ByVal key As String)
    On Error Resume Next
    util.GrabaIni IniPath, "toolbars", key & "_docked", "0"
    Err = 0
End Sub


Private Sub vbsCSS1_ItemSelected(ByVal Atributo As String)

    If Not ActiveForm Is Nothing Then
        If frmMain.ActiveForm.Name = "frmEdit" Then
            ActiveForm.Insertar Atributo
        End If
    End If
    
End Sub


Private Sub vbSDhtml1_ElementClicked(ByVal Value As String)

    If Not ActiveForm Is Nothing Then
        If frmMain.ActiveForm.Name = "frmEdit" Then
            ActiveForm.Insertar Value
        End If
    End If
    
End Sub

Private Sub vbsXHTML1_ItemSelected(ByVal Atributo As String)

    If Not ActiveForm Is Nothing Then
        If frmMain.ActiveForm.Name = "frmEdit" Then
            ActiveForm.Insertar Atributo
        End If
    End If
    
End Sub


