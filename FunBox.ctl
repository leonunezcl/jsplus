VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{2128BF45-F895-4206-84CD-F4DE2DD8D6B1}#2.0#0"; "vbsTbar6.ocx"
Object = "{E7106799-3A07-4335-80BA-4F20E8E5E2E9}#2.0#0"; "vbsODCL6.ocx"
Object = "{98F993CC-3598-405A-9E9A-0D2CF198B250}#2.0#0"; "vbsDkTb6.ocx"
Begin VB.UserControl FunBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   ScaleHeight     =   3600
   ScaleWidth      =   7380
   ToolboxBitmap   =   "FunBox.ctx":0000
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      ScaleHeight     =   210
      ScaleWidth      =   3705
      TabIndex        =   4
      Top             =   375
      Width           =   3735
      Begin VB.Label lbltotfun 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   3255
         TabIndex        =   6
         Top             =   0
         Width           =   90
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Function Count :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   5
         Top             =   0
         Width           =   1425
      End
   End
   Begin MSComctlLib.ListView lvwSort 
      Height          =   750
      Left            =   105
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1350
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1323
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "funcion"
         Object.Width           =   2540
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
      ScaleWidth      =   7320
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   7380
      Begin vbalTBar6.cToolbar tbrTools 
         Height          =   270
         Left            =   120
         Top             =   45
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   476
      End
   End
   Begin vbalDkTb6.vbalDockContainer vbalDockContainer1 
      Align           =   1  'Align Top
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   53
      AllowUndock     =   0   'False
   End
   Begin ODCboLst6.OwnerDrawComboList lstFun 
      Height          =   900
      Left            =   4425
      TabIndex        =   2
      ToolTipText     =   "Clic to select function ..."
      Top             =   1965
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1588
      ExtendedUI      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   4
      MaxLength       =   0
   End
End
Attribute VB_Name = "FunBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Img As cVBALImageList
Private WithEvents m_cMenu As cPopupMenu
Attribute m_cMenu.VB_VarHelpID = -1
Public Event FunctionSelected(ByVal funcion As String)
Private mFunctions As New Collection
Private mCurrentFunction As String
Private m_FileName As String


Public Sub AddFun(ByVal funcion As String)

    lstFun.AddItemAndData funcion, 5
    lbltotfun.Caption = lstFun.ListCount '- 1
    
End Sub


Public Sub Clear()
    lstFun.Clear
    If mFunctions.count > 0 Then
        Dim k As Integer
        For k = mFunctions.count To 1 Step -1
            mFunctions.Remove k
        Next k
    End If
    lbltotfun.Caption = "0"
End Sub



Public Sub Prepare()
    
    Dim iMain As Long
    Dim ip As Long
    
    Set m_cMenu = New cPopupMenu
    m_cMenu.hWndOwner = UserControl.hwnd
    m_cMenu.OfficeXpStyle = True
    
    Set m_Img = New cVBALImageList
    
    With m_Img
        .IconSizeX = 16: .IconSizeY = 16: .ColourDepth = ILC_COLOR24
        .Create
        .AddFromResourceID 190, App.hInstance, IMAGE_ICON, "k1"
        .AddFromResourceID 242, App.hInstance, IMAGE_ICON, "k2"
        .AddFromResourceID 243, App.hInstance, IMAGE_ICON, "k3"
        .AddFromResourceID 114, App.hInstance, IMAGE_ICON, "k4"
        .AddFromResourceID 117, App.hInstance, IMAGE_ICON, "k5"
        .AddFromResourceID 191, App.hInstance, IMAGE_ICON, "k6"
    End With
    
    m_cMenu.ImageList = m_Img.hIml
    lstFun.ImageList = m_Img.hIml
    
    With m_cMenu
        iMain = .AddItem("TOOLS", "Tools Toolbar", , , , , , "TOOLSTOOLBAR")
        ip = .AddItem("New", "New Function", , iMain, 0, , , "TOOLS:NEW")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Ascending", "Sorts function ascending", , iMain, 1, , , "TOOLS:ASC")
        ip = .AddItem("Descending", "Sorts function descending", , iMain, 2, , , "TOOLS:DES")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Save", "Save to file", , iMain, 3, , , "TOOLS:SAVE")
        ip = .AddItem("Preview", "Preview in browser", , iMain, 4, , , "TOOLS:FPREVIEW")
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

Private Sub SaveFunctions()

    Dim Archivo As String
    Dim nFreeFile As Long
    Dim k As Integer
    Dim ret As String
    
    If lstFun.ListCount = 0 Then Exit Sub
        
    ret = "Text Files (*.txt)|*.txt|"
    ret = ret & "All Files (*.*)|*.*"
            
    Call Cdlg.VBGetSaveFileName(Archivo, , , ret, , LastPath, "Save functions ...", "txt", frmMain.hwnd)
               
    If Len(Archivo) > 0 Then
        nFreeFile = FreeFile
        Open Archivo For Output As #nFreeFile
        For k = 0 To lstFun.ListCount - 1
            Print #nFreeFile, lstFun.List(k)
        Next k
        Close #nFreeFile
    End If
    
End Sub
Private Sub SaveHtml()

    On Local Error GoTo ErrorImprimir
    
    Dim Archivo As String
    Dim k As Integer
    Dim Itmx As ListItem
    Dim nFreeFile As Integer
    Dim Fuente As String
    Dim ret As String
    
    If lstFun.ListCount = 0 Then Exit Sub
       
    ret = "HTML Files (*.html)|*.html|"
    ret = ret & "All Files (*.*)|*.*"
            
    Call Cdlg.VBGetSaveFileName(Archivo, , , ret, , LastPath, "Save As ...", "html", frmMain.hwnd)
    
    If Len(Archivo) = 0 Then Exit Sub
    
    nFreeFile = FreeFile
    
    Call Hourglass(hwnd, True)
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    Open Archivo For Output As #nFreeFile
        'cabezera del archivo
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>JavaScript Plus Report</title>"
        Print #nFreeFile, Replace("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>", "'", Chr$(34))
        Print #nFreeFile, "</head>"
        
        'titulo del reporte
        Print #nFreeFile, Replace("<body bgcolor='#FFFFFF' text='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Report Date : " & Now & " </b></p>"
        Print #nFreeFile, "<p><b>File : " & util.VBArchivoSinPath(m_FileName) & "</b></p>"
        Print #nFreeFile, "</font>"
        
        'generar titulos
        Print #nFreeFile, Replace("<table border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bgcolor='#999999' bordercolor='#000000'>", "'", Chr$(34))
        
        Print #nFreeFile, Replace("<td><b>" & Fuente & "Number</font></b></td>", "'", Chr$(34))
        Print #nFreeFile, Replace("<td><b>" & Fuente & "Name</font></b></td>", "'", Chr$(34))
        
        Print #nFreeFile, "</tr>"
        
        For k = 0 To lstFun.ListCount - 1
            
            'imprimir informacion
            Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
            
            'correlativo
            Print #nFreeFile, Replace("<td>" & Fuente & k + 1 & "</font></td>", "'", Chr$(34))
            Print #nFreeFile, Replace("<td>" & Fuente & lstFun.List(k) & "</font></td>", "'", Chr$(34))
                            
            Print #nFreeFile, "</tr>"
        Next k
        
        Print #nFreeFile, "</table>"
                
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
    Close #nFreeFile
        
    GoTo SalirImprimir
    
ErrorImprimir:
    Resume SalirImprimir
    
SalirImprimir:
    Call Hourglass(hwnd, False)
    Err = 0
    
End Sub
Private Sub Sort(ByVal tipo As String)

    Dim k As Integer
    
    If lstFun.ListCount > -1 Then
        lvwSort.ListItems.Clear
        
        For k = 0 To lstFun.ListCount - 1
            lvwSort.ListItems.Add , "k" & k, lstFun.List(k)
        Next k
        
        lvwSort.Sorted = True
        lvwSort.SortKey = 0
        
        If tipo = "A" Then
            lvwSort.SortOrder = lvwAscending
        Else
            lvwSort.SortOrder = lvwDescending
        End If
                            
        lstFun.Clear
        
        For k = 1 To lvwSort.ListItems.count
            AddFun lvwSort.ListItems(k).Text
        Next k
    End If
    
    lbltotfun.Caption = lstFun.ListCount '- 1
End Sub

Private Sub lstFun_Click()
    If Len(lstFun.Text) > 0 Then
        RaiseEvent FunctionSelected(lstFun.Text)
    End If
End Sub

Private Sub tbrTools_ButtonClick(ByVal lButton As Long)

    Select Case tbrTools.ButtonKey(lButton)
        Case "TOOLS:NEW"
            If Not frmMain.ActiveForm Is Nothing Then
                Dim funcion As String
                funcion = InputBox("Function Name:", "New Function")
                If Len(Trim$(funcion)) > 0 Then
                    frmMain.ActiveForm.Insertar "function " & funcion & "()" & vbNewLine & "{" & vbNewLine & vbNewLine & "}"
                End If
            End If
        Case "TOOLS:ASC"
            Call Sort("A")
        Case "TOOLS:DES"
            Call Sort("B")
        Case "TOOLS:SAVE"
            Call SaveFunctions
        Case "TOOLS:FPREVIEW"
            Call SaveHtml
    End Select
    
End Sub

Private Sub UserControl_Resize()
    
    On Error Resume Next
    LockWindowUpdate hwnd
    picBottom.Move 0, picGeneral.Height + 1, UserControl.Width, 240
    lstFun.Move 0, picGeneral.Height + picBottom.Height + 1, UserControl.Width, UserControl.Height - picGeneral.Height - 245
    lbltotfun.Move picGeneral.Width - lbltotfun.Width - 100
    LockWindowUpdate False
    Err = 0
End Sub



Public Property Get GetFunctions() As Collection
    
    Dim k As Integer
    
    If mFunctions.count > 0 Then
        For k = mFunctions.count To 1 Step -1
            mFunctions.Remove k
        Next k
    End If
    
    For k = 0 To lstFun.ListCount - 1
        mFunctions.Add lstFun.List(k), "k" & k
    Next k
    
    Set GetFunctions = mFunctions
    
End Property


Public Property Get CurrentFunction() As String
    CurrentFunction = mCurrentFunction
End Property

Public Property Let CurrentFunction(ByVal pCurrentFunction As String)
    mCurrentFunction = pCurrentFunction
End Property


Public Property Get filename() As String
    filename = m_FileName
End Property

Public Property Let filename(ByVal pFileName As String)
    m_FileName = pFileName
End Property

