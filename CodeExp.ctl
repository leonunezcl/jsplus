VERSION 5.00
Object = "{2128BF45-F895-4206-84CD-F4DE2DD8D6B1}#2.0#0"; "vbsTbar6.ocx"
Object = "{8C44B082-B582-4258-9E2C-7D9383CE7DF4}#1.0#0"; "vbsTreeView6.ocx"
Object = "{98F993CC-3598-405A-9E9A-0D2CF198B250}#2.0#0"; "vbsDkTb6.ocx"
Begin VB.UserControl CodeExp 
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3405
   ScaleHeight     =   4590
   ScaleWidth      =   3405
   ToolboxBitmap   =   "CodeExp.ctx":0000
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
      ScaleWidth      =   3345
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   3405
      Begin vbalTBar6.cToolbar tbrTools 
         Height          =   270
         Left            =   120
         Top             =   45
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   476
      End
   End
   Begin vbalTreeViewLib6.vbalTreeView tvwCodeExp 
      Height          =   1260
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Double clic to Collapse/Expand node"
      Top             =   450
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2223
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbalDkTb6.vbalDockContainer vbalDockContainer1 
      Align           =   1  'Align Top
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   375
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   53
      AllowUndock     =   0   'False
   End
End
Attribute VB_Name = "CodeExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cExplorer As New cCodeExplorer
Private m_Img As cVBALImageList
Private WithEvents m_cMenu As cPopupMenu
Attribute m_cMenu.VB_VarHelpID = -1
Private nodImages As cTreeViewNode
Private nodHyper As cTreeViewNode
Private nodSheets As cTreeViewNode
Private nodScripts As cTreeViewNode
Private m_UrlSource As String

Public Event ImageClick(ByVal valor As String)
Public Event LinkClick(ByVal valor As String)
Public Event ScriptClick(ByVal valor As String)
Public Event StyleClick(ByVal valor As String)

Private Type eInfoArchivo
    Archivo As String
    UrlArchivo As String
End Type
Private arr_img() As eInfoArchivo
Private arr_scripts() As eInfoArchivo
Private arr_stylesheets() As eInfoArchivo

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Function DownloadFile(url As String, localfilename As String) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, url, localfilename, 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
    DoEvents
End Function

Private Function BuscaPunto(ByVal info As String) As Boolean

    Dim ret As Boolean
    Dim Archivo As String
    Dim k As Integer
    
    For k = Len(info) To 1 Step -1
        If Mid$(info, k, 1) = "/" Then
            Archivo = Mid$(info, k + 1)
            If InStr(Archivo, ".") > 0 Then
                ret = True
            End If
            Exit For
        End If
    Next k
    
    BuscaPunto = ret
    
End Function

Private Sub DownloadElements()

    Dim bytes() As Byte
    Dim fnum As Integer
    Dim Path As String
    Dim k As Integer
    Dim j As Integer
    Dim C As Integer
    Dim found As Boolean
    
    If cExplorer.Images.count = 0 And cExplorer.Scripts.count = 0 And cExplorer.StyleSheets.count = 0 And cExplorer.Links.count = 0 Then Exit Sub
    
    Path = util.BrowseFolder(frmMain.hwnd)
    
    If Len(Path) > 0 Then
    
        util.Hourglass hwnd, True
        
        Path = util.StripPath(Path)
        
        Call PreparaDescarga
        
        If UBound(arr_img) > 0 Then
            
            Load frmOpenFiles
            frmOpenFiles.Caption = "Downloading Images"
            frmOpenFiles.pgb.Max = UBound(arr_img)
            frmOpenFiles.Show
                
            On Error Resume Next
            
            For k = 1 To UBound(arr_img)
            
                If frmOpenFiles.Cancelo Then
                    Unload frmOpenFiles
                    util.Hourglass hwnd, False
                    Exit Sub
                End If
                                
                frmOpenFiles.lblFile.Caption = arr_img(k).Archivo
                
                DownloadFile arr_img(k).UrlArchivo, Path & arr_img(k).Archivo

                frmOpenFiles.pgb.Value = k
                Err = 0
                j = j + 1
            Next k
            
            Err = 0
            
            Unload frmOpenFiles
        End If
        
        If UBound(arr_scripts) > 0 Then
        
            Load frmOpenFiles
            frmOpenFiles.Caption = "Downloading Scripts"
            frmOpenFiles.pgb.Max = UBound(arr_scripts)
            frmOpenFiles.Show
            
            On Error Resume Next
            
            For k = 1 To UBound(arr_scripts)
            
                If frmOpenFiles.Cancelo Then
                    Unload frmOpenFiles
                    util.Hourglass hwnd, False
                    Exit Sub
                End If
                
                frmOpenFiles.lblFile.Caption = arr_scripts(k).Archivo
                
                DownloadFile arr_img(k).UrlArchivo, Path & arr_img(k).Archivo
                
                frmOpenFiles.pgb.Value = k
                j = j + 1
                Err = 0
            Next k
            Err = 0
            Unload frmOpenFiles
        End If
        
        If UBound(arr_stylesheets) > 0 Then
            
            Load frmOpenFiles
            frmOpenFiles.Caption = "Downloading StyleSheets"
            frmOpenFiles.pgb.Max = UBound(arr_stylesheets)
            frmOpenFiles.Show
            
            On Error Resume Next
            
            For k = 1 To UBound(arr_stylesheets)
            
                If frmOpenFiles.Cancelo Then
                    Unload frmOpenFiles
                    util.Hourglass hwnd, False
                    Exit Sub
                End If
                
                frmOpenFiles.lblFile.Caption = arr_stylesheets(k).Archivo
                
                DownloadFile arr_img(k).UrlArchivo, Path & arr_img(k).Archivo
                
                frmOpenFiles.pgb.Value = k
                j = j + 1
                Err = 0
            Next k
            Err = 0
            Unload frmOpenFiles
        End If
        
        If j > 0 Then
            MsgBox CStr(j) & " files downloaded to " & Path, vbInformation
        End If
        
        util.Hourglass hwnd, False
    End If
    
End Sub

Private Sub eliminar_nodos()

    On Error Resume Next
    nodImages.Delete
    nodHyper.Delete
    nodSheets.Delete
    nodScripts.Delete
    Err = 0
    
End Sub

Public Sub LoadCode(ByVal url As String)
    
    code_explorer url
        
End Sub
Public Sub code_explorer(ByVal Archivo As String)
 
    Dim k As Integer
    
    cExplorer.filename = Archivo
    If Not cExplorer.Explore Then
        Exit Sub
    End If
       
    Call eliminar_nodos
        
    Set nodImages = tvwCodeExp.Nodes.Add(, etvwChild, "kimgs", "Images", 0)
    nodImages.Bold = True
    
    Set nodHyper = tvwCodeExp.Nodes.Add(, etvwChild, "khyper", "Hyperlinks", 0)
    nodHyper.Bold = True
    
    Set nodSheets = tvwCodeExp.Nodes.Add(, etvwChild, "ksheets", "Style Sheets", 0)
    nodSheets.Bold = True
    
    Set nodScripts = tvwCodeExp.Nodes.Add(, etvwChild, "kscripts", "Scripts", 0)
    nodScripts.Bold = True
        
    With cExplorer
        For k = 1 To .Images.count
            nodImages.AddChildNode "kimgs_" & k, .Images.ITem(k), 1 '97
        Next k
        
        For k = 1 To .Links.count
            nodHyper.AddChildNode "khyper_" & k, .Links.ITem(k), 2 '96
        Next k
        
        For k = 1 To .Scripts.count
            nodScripts.AddChildNode "kscripts_" & k, .Scripts.ITem(k), 3 '103
        Next k
        
        For k = 1 To .StyleSheets.count
            nodSheets.AddChildNode "ksheets_" & k, .StyleSheets.ITem(k), 4 '125
        Next k
    End With
   
    nodImages.Expanded = True
    nodHyper.Expanded = True
    nodScripts.Expanded = True
    nodSheets.Expanded = True
    nodImages.Selected = True
    
End Sub
Private Sub BuildImageList()
    
    Dim iMain As Long
    Dim ip As Long
    
    Set m_cMenu = New cPopupMenu
    m_cMenu.hWndOwner = UserControl.hwnd
    m_cMenu.OfficeXpStyle = True
    
    Set m_Img = New cVBALImageList
    
    With m_Img
        .IconSizeX = 16: .IconSizeY = 16: .ColourDepth = ILC_COLOR24
        .Create
        .AddFromResourceID 203, App.hInstance, IMAGE_ICON, "k1"
        .AddFromResourceID 198, App.hInstance, IMAGE_ICON, "k2"
        .AddFromResourceID 197, App.hInstance, IMAGE_ICON, "k3"
        .AddFromResourceID 204, App.hInstance, IMAGE_ICON, "k4"
        .AddFromResourceID 226, App.hInstance, IMAGE_ICON, "k5"
        .AddFromResourceID 114, App.hInstance, IMAGE_ICON, "k6"
        .AddFromResourceID 282, App.hInstance, IMAGE_ICON, "k7"
        .AddFromResourceID 284, App.hInstance, IMAGE_ICON, "k8"
    End With
    
    m_cMenu.ImageList = m_Img.hIml
    tvwCodeExp.ImageList = m_Img.hIml
    
    With m_cMenu
        iMain = .AddItem("TOOLS", "Tools Toolbar", , , , , , "TOOLSTOOLBAR")
        ip = .AddItem("Save", "Save to file", , iMain, 5, , , "TOOLS:SAVE")
        ip = .AddItem("Download", "Download", , iMain, 6, , , "TOOLS:DOWNLOAD")
        ip = .AddItem("-", , , iMain)
        ip = .AddItem("Config", "Configure extension", , iMain, 7, , , "TOOLS:CONFIG")
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

Private Sub PreparaDescarga()

    Dim k As Integer
    Dim C As Integer
    Dim j As Integer
    Dim found As Boolean
    Dim Archivo As String
    Dim dato As String
    Dim url As String
    
    ReDim arr_img(0)
    ReDim arr_scripts(0)
    ReDim arr_stylesheets(0)
    
    If Right$(m_UrlSource, 1) <> "/" Then
        url = m_UrlSource & "/"
    Else
        url = m_UrlSource
    End If
                        
    With cExplorer
        
        If .Images.count > 0 Then
            C = 1
            For k = 1 To .Images.count
                found = False
                
                dato = .Images.ITem(k)
                
                Archivo = ""
                
                If BuscaPunto(dato) Then
                    If Left$(dato, 7) = "http://" Then
                        Archivo = Replace(dato, "'", "")
                        Archivo = Replace(Archivo, "/", "\")
                        Archivo = util.VBArchivoSinPath(Archivo)
                    ElseIf Len(url) > 0 Then
                    
                        If Left$(dato, 1) = "/" Then
                            dato = url & Mid$(dato, 2)
                        Else
                            dato = url & dato
                        End If
                        
                        Archivo = Replace(dato, "'", "")
                        Archivo = Replace(Archivo, "/", "\")
                        Archivo = util.VBArchivoSinPath(Archivo)
                    End If
                    
                    If Len(Archivo) > 0 Then
                        For j = 1 To UBound(arr_img)
                            If LCase$(arr_img(j).Archivo) = LCase$(Archivo) Then
                                found = True
                                Exit For
                            End If
                        Next j
                        
                        If Not found Then
                            ReDim Preserve arr_img(C)
                            arr_img(C).Archivo = Archivo
                            arr_img(C).UrlArchivo = Replace(dato, "'", "")
                            C = C + 1
                        End If
                    End If
                End If
            Next k
        End If
        
        If .Scripts.count > 0 Then
            C = 1
            For k = 1 To .Scripts.count
                found = False
                
                dato = .Scripts.ITem(k)
                
                Archivo = ""
                
                If BuscaPunto(dato) Then
                    If Left$(dato, 7) = "http://" Then
                        Archivo = Replace(dato, "'", "")
                        Archivo = Replace(Archivo, "/", "\")
                        Archivo = util.VBArchivoSinPath(Archivo)
                    ElseIf Len(url) > 0 Then
                    
                        If Left$(dato, 1) = "/" Then
                            dato = url & Mid$(dato, 2)
                        Else
                            dato = url & dato
                        End If
                        
                        Archivo = Replace(dato, "'", "")
                        Archivo = Replace(Archivo, "/", "\")
                        Archivo = util.VBArchivoSinPath(Archivo)
                    End If
                    
                    If Len(Archivo) > 0 Then
                        For j = 1 To UBound(arr_scripts)
                            If LCase$(arr_scripts(j).Archivo) = LCase$(Archivo) Then
                                found = True
                                Exit For
                            End If
                        Next j
                        
                        If Not found Then
                            ReDim Preserve arr_scripts(C)
                            arr_scripts(C).Archivo = Archivo
                            arr_scripts(C).UrlArchivo = Replace(dato, "'", "")
                            C = C + 1
                        End If
                    End If
                End If
            Next k
        End If
        
        If .StyleSheets.count > 0 Then
            C = 1
            For k = 1 To .StyleSheets.count
                found = False
                
                dato = .StyleSheets.ITem(k)
                
                Archivo = ""
                
                If BuscaPunto(dato) Then
                    If Left$(dato, 7) = "http://" Then
                        Archivo = Replace(dato, "'", "")
                        Archivo = Replace(Archivo, "/", "\")
                        Archivo = util.VBArchivoSinPath(Archivo)
                    ElseIf Len(url) > 0 Then
                    
                        If Left$(dato, 1) = "/" Then
                            dato = url & Mid$(dato, 2)
                        Else
                            dato = url & dato
                        End If
                        
                        Archivo = Replace(dato, "'", "")
                        Archivo = Replace(Archivo, "/", "\")
                        Archivo = util.VBArchivoSinPath(Archivo)
                    End If
                    
                    If Len(Archivo) > 0 Then
                        For j = 1 To UBound(arr_stylesheets)
                            If LCase$(arr_stylesheets(j).Archivo) = LCase$(Archivo) Then
                                found = True
                                Exit For
                            End If
                        Next j
                        
                        If Not found Then
                            ReDim Preserve arr_stylesheets(C)
                            arr_stylesheets(C).Archivo = Archivo
                            arr_stylesheets(C).UrlArchivo = Replace(dato, "'", "")
                            C = C + 1
                        End If
                    End If
                End If
            Next k
        End If
    End With
End Sub
Private Sub SaveElements()

    On Error GoTo ErrorSaveElements
    
    Dim Archivo As String
    Dim nFreeFile As Long
    Dim k As Integer
    Dim ret As String
    
    If cExplorer.Images.count = 0 And cExplorer.Scripts.count = 0 And cExplorer.StyleSheets.count = 0 And cExplorer.Links.count = 0 Then Exit Sub
    
    ret = "Text Files (*.txt)|*.txt|"
    ret = ret & "All Files (*.*)|*.*"
            
    Call Cdlg.VBGetSaveFileName(Archivo, , , ret, , LastPath, "Save elements ...", "txt", frmMain.hwnd)
               
    util.Hourglass hwnd, True
    
    If Len(Archivo) > 0 Then
        nFreeFile = FreeFile
        Open Archivo For Output As #nFreeFile
        
        With cExplorer
            Print #nFreeFile, "------"
            Print #nFreeFile, "Images"
            Print #nFreeFile, "------"
            For k = 1 To .Images.count
                Print #nFreeFile, .Images.ITem(k)
            Next k
            
            Print #nFreeFile, "------"
            Print #nFreeFile, "Links"
            Print #nFreeFile, "------"
            For k = 1 To .Links.count
                Print #nFreeFile, .Links.ITem(k)
            Next k
            
            Print #nFreeFile, "------"
            Print #nFreeFile, "Scripts"
            Print #nFreeFile, "------"
            For k = 1 To .Scripts.count
                Print #nFreeFile, .Scripts.ITem(k)
            Next k
            
            Print #nFreeFile, "------"
            Print #nFreeFile, "StyleSheets"
            Print #nFreeFile, "------"
            For k = 1 To .StyleSheets.count
                Print #nFreeFile, .StyleSheets.ITem(k)
            Next k
            Print #nFreeFile, "------"
        End With
        Close #nFreeFile
    End If
    
    util.Hourglass hwnd, False
    
    Exit Sub
    
ErrorSaveElements:
    util.Hourglass hwnd, False
    If nFreeFile > 0 Then Close #nFreeFile
    MsgBox "CodeExp.SaveElements - Error : " & Error$, vbCritical
        
End Sub

Private Sub tbrTools_ButtonClick(ByVal lButton As Long)

    Select Case tbrTools.ButtonKey(lButton)
        Case "TOOLS:SAVE"
            Call SaveElements
        Case "TOOLS:DOWNLOAD"
            Call DownloadElements
        Case "TOOLS:CONFIG"
            frmSetTreeExp.origen = 0
            frmSetTreeExp.Show vbModal
    End Select
    
End Sub


Private Sub tvwCodeExp_NodeClick(Node As vbalTreeViewLib6.cTreeViewNode)

    If Not Node Is Nothing Then
        If InStr(Node.key, "khyper_") Then
            RaiseEvent LinkClick(Node.Text)
        ElseIf InStr(Node.key, "kimgs_") Then
            RaiseEvent ImageClick(Node.Text)
        ElseIf InStr(Node.key, "kscripts_") Then
            RaiseEvent ScriptClick(Node.Text)
        ElseIf InStr(Node.key, "ksheets_") Then
            RaiseEvent StyleClick(Node.Text)
        End If
    End If
    
End Sub


Private Sub UserControl_Initialize()
    BuildImageList
End Sub

Private Sub UserControl_Resize()
    
    On Error Resume Next
    LockWindowUpdate hwnd
    tvwCodeExp.Move 0, picGeneral.Height + 1, UserControl.Width, UserControl.Height - picGeneral.Height - 245
    LockWindowUpdate False
    Err = 0
    
End Sub

Public Property Get UrlSource() As String
    UrlSource = m_UrlSource
End Property

Public Property Let UrlSource(ByVal pUrlSource As String)
    m_UrlSource = pUrlSource
End Property
