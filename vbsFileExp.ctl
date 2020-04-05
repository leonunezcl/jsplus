VERSION 5.00
Object = "{E7106799-3A07-4335-80BA-4F20E8E5E2E9}#2.0#0"; "vbsODCL6.ocx"
Begin VB.UserControl vbsFileExp 
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2745
   LockControls    =   -1  'True
   ScaleHeight     =   4890
   ScaleWidth      =   2745
   ToolboxBitmap   =   "vbsFileExp.ctx":0000
   Begin VB.DriveListBox driMain 
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Top             =   0
      Width           =   2565
   End
   Begin VB.DirListBox dirMain 
      Height          =   990
      Left            =   105
      TabIndex        =   2
      Top             =   420
      Width           =   2490
   End
   Begin VB.ComboBox cboFilFil 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2295
      Width           =   2310
   End
   Begin VB.FileListBox filWrk 
      Height          =   480
      Left            =   3195
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3795
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ODCboLst6.OwnerDrawComboList filMain 
      Height          =   1275
      Left            =   315
      TabIndex        =   5
      ToolTipText     =   "Double clic to open selected file"
      Top             =   3030
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   2249
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
      FullRowSelect   =   -1  'True
      MaxLength       =   0
   End
   Begin VB.Label lblFil 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "View files that match this pattern:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   2070
      Width           =   2325
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuFile_Open 
         Caption         =   "Open ..."
      End
      Begin VB.Menu mnuFile_Delete 
         Caption         =   "Delete ..."
      End
      Begin VB.Menu mnuFile_Refresh 
         Caption         =   "Refresh ...."
      End
      Begin VB.Menu mnuFile_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Properties 
         Caption         =   "Properties ..."
      End
   End
   Begin VB.Menu mnuFolder 
      Caption         =   "Folder"
      Visible         =   0   'False
      Begin VB.Menu mnuFolder_New 
         Caption         =   "New Folder ..."
      End
      Begin VB.Menu mnuFolder_Properties 
         Caption         =   "Properties ...."
      End
   End
End
Attribute VB_Name = "vbsFileExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Private ini As New cInifile
Private m_IniFile As String

Private m_cIL As New cVBALSysImageList
Private arr_filtros() As String
Public Event FileClicked(ByVal File As String)
Public Sub Load()

    Dim V As Variant
    Dim k As Integer
    Dim j As Integer
    Dim ele As String
    
    ReDim arr_filtros(0)
    j = 1
    
    'cargar filtros
    V = LeeIni(m_IniFile, "filelist", "num")
    
    For k = 1 To V
        ele = LeeIni(m_IniFile, "filelist", "ele" & k)
        If Len(ele) > 0 Then
            cboFilFil.AddItem Explode(ele, 1, "|")
            ReDim Preserve arr_filtros(j)
            arr_filtros(j) = Explode(ele, 2, "|")
            j = j + 1
        End If
    Next k
        
    'setear directorio
    Set m_cIL = New cVBALSysImageList
    m_cIL.IconSizeX = 16
    m_cIL.IconSizeY = 16
    m_cIL.Create
      
    filMain.ImageList = m_cIL.hIml
    
    If UBound(arr_filtros) = 0 Then
        dirMain_Change
    Else
        cboFilFil.ListIndex = 0 ' cboFilFil.ListCount - 1
    End If
    
End Sub


Private Function Explode(ElementsList As String, Index As Integer, Optional Separator As String = vbTab) As String
    Dim SubStr2Explode As String
    Dim auxI           As Integer
    Dim Element        As String

    On Error Resume Next

    SubStr2Explode = ElementsList

    For auxI = 1 To Index
        If InStr(SubStr2Explode, Separator) = 0 Then
            If auxI = Index Then
                Element = SubStr2Explode
            Else
                Element = vbNullString
            End If

            Exit For
        End If

        Element = Mid$(SubStr2Explode, 1, InStr(SubStr2Explode, Separator) - 1)
        SubStr2Explode = Mid$(SubStr2Explode, InStr(SubStr2Explode, Separator) + 1)
    Next auxI

    Explode = Element

    On Error GoTo 0
End Function

Public Function LeeIni(ByVal Archivo As String, ByVal seccion As String, ByVal llave As String) As String

    Dim ret As Long
    
    Dim buffer As String
    
    buffer = Space$(1000)
        
    ret = GetPrivateProfileString(seccion, llave, "", buffer, Len(buffer), Archivo)
    
    buffer = Trim$(buffer)
    buffer = VBA.Left$(buffer, Len(buffer) - 1)
    
    LeeIni = buffer
    
End Function

Private Sub cboFilFil_Click()
    If UBound(arr_filtros) > 0 Then
        filWrk.Pattern = arr_filtros(cboFilFil.ListIndex + 1)
        dirMain_Change
        'LoadPath dirMain.path
    End If
End Sub


Private Sub dirMain_Change()
    filWrk.Path = dirMain.Path
    SetPath dirMain.Path
End Sub

Private Sub SetPath(ByVal sPath As String)
   If VBA.Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
   LoadPath sPath
End Sub



Private Sub dirMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuFolder
    End If
    
End Sub

Private Sub driMain_Change()

    On Error GoTo error_path
    dirMain.Path = driMain.Drive
    Exit Sub
error_path:
    dirMain.Path = App.Path

End Sub

Private Sub LoadPath(ByVal sPath As String)

    Dim sName As String
    Dim iC As Long

    util.Hourglass hwnd, True
    filMain.Clear
    filWrk.Refresh
    For iC = 0 To filWrk.ListCount - 1
        sName = filWrk.List(iC)
        
        filMain.AddItemAndData sName, m_cIL.ItemIndex(sPath & sName, True), , , , , , m_cIL.IconSizeX, , eixVCentre
        
        If iC Mod 50 = 0 Then
            DoEvents
        End If
    Next iC
    util.Hourglass hwnd, False
End Sub

Private Sub filMain_DblClick()
    
    If Len(filMain.Text) > 0 Then
        Dim Archivo As String
        Archivo = util.StripPath(dirMain.Path) & filMain.Text
        RaiseEvent FileClicked(Archivo)
    End If
    
End Sub


Private Sub filMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbRightButton Then
        PopupMenu mnuFile
    End If
    
End Sub


Private Sub mnuFile_Delete_Click()

    If Len(filMain.Text) > 0 Then
        Dim Archivo As String
        Archivo = util.StripPath(dirMain.Path) & filMain.Text
        
        If Confirma("Are you sure to delete the selected file") = vbYes Then
            util.BorrarArchivo Archivo
        End If
        dirMain_Change
    End If
    
End Sub

Private Sub mnuFile_Open_Click()
    
    Dim k As Integer
    Dim Archivo As String
    
    For k = 0 To filMain.ListCount - 1
      If filMain.Selected(k) Then
         Archivo = util.StripPath(dirMain.Path) & filMain.List(k)
         RaiseEvent FileClicked(Archivo)
      End If
    Next k
    
End Sub


Private Sub mnuFile_Properties_Click()
    If Len(filMain.Text) > 0 Then
        Dim Archivo As String
        Archivo = util.StripPath(dirMain.Path) & filMain.Text
        util.PropiedadesArchivo Archivo, hwnd
    End If
End Sub

Private Sub mnuFile_Refresh_Click()
    SetPath dirMain.Path
End Sub

Private Sub mnuFolder_New_Click()
    
    Dim Folder As String
    Dim Path As String
    
    Path = util.StripPath(dirMain.Path)
    
    Folder = InputBox("Folder Name", "Name:")
    
    If Len(Folder) > 0 Then
        util.CrearDirectorio Path & Folder
        dirMain.Refresh
        dirMain_Change
    End If
    
End Sub

Private Sub mnuFolder_Properties_Click()
    If Len(dirMain.Path) > 0 Then
        Dim Path As String
        Path = util.StripPath(dirMain.Path)
        util.PropiedadesArchivo Path, hwnd
    End If
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
            
    Const Top = 0
    LockWindowUpdate hwnd
    driMain.Move 0, Top, UserControl.Width
    dirMain.Move 0, (driMain.Height + 1 + Top), UserControl.Width, UserControl.Height / 2
    cboFilFil.Move 0, UserControl.Height - cboFilFil.Height, UserControl.Width
    lblFil.Move 5, cboFilFil.Top - lblFil.Height - 10
    filMain.Move 0, (driMain.Height + dirMain.Height + 1 + Top), UserControl.Width, ((UserControl.Height / 2) - driMain.Height) - Top - cboFilFil.Height - lblFil.Height
    Err = 0
    LockWindowUpdate False
End Sub




Public Property Get inifile() As String
    inifile = m_IniFile
End Property

Public Property Let inifile(ByVal pIniFile As String)
    m_IniFile = pIniFile
End Property
