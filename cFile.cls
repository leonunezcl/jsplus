VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_IdDoc As Integer      'id del documento
Private m_FileName As String    'nombre del archivo
Private m_Ftp As Boolean        'archivo desde ftp
Private m_TempFile As String   'nombre archivo temporal para ftp
Private m_SiteName As String    'nombre del sitio para ftp
Private m_Url As String         'path del archivo en ftp
Private m_User As String        'usuario ftp
Private m_Pwd As String         'pwd del usuario ftp
Private m_Port As Long          'puerto ftp del archivo
Private m_FtpFile As String
Private m_Changed As Boolean    'cambio en archivo destino
Private m_Caption As String     'caption del archivo en tab
Private m_RemoteFolder As String 'carpeta remota
Private m_Passive As Integer

Public Function CloseFile(frm As Form) As Boolean

    'If frm.txtCode.Modified Then
    Dim Msg As String
    Dim File As String
    Dim rpta As VbMsgBoxResult
    
    If InStr(frm.Caption, "*") Then
        File = Replace(Me.Caption, "*", "")
        Msg = "Do you want to save changes to " & File & "?"
        
        rpta = MsgBox(Msg, vbQuestion + vbYesNoCancel + vbDefaultButton3)
        If rpta = vbYes Then
        'If Confirma(Msg) = vbYes Then
            If SaveFile(frm, False) Then
                If m_Ftp Then
                    util.BorrarArchivo m_TempFile
                End If
                CloseFile = True
            Else
                'error al grabar no se cierra
                CloseFile = False
            End If
        ElseIf rpta = vbCancel Then 'Cancelar
            CloseFile = False
        Else    'NO Grabar cambios
            frm.NoSave = True
            CloseFile = True
        End If
    Else
        CloseFile = True
    End If
    
End Function
Public Sub NewFile()
    
    Dim ft As New frmEdit
    
    m_Caption = "Document " & m_IdDoc
    
    ft.Show
    ft.Caption = m_Caption
    ft.Tag = m_IdDoc
    
    ListaLangs.SetLang "", ft.txtCode
    
    On Error Resume Next
    
    ft.txtCode.SetFocus
    
End Sub

Public Function OpenFile() As Boolean

    On Error GoTo ErrorOpenFile
    
    Dim ft As New frmEdit
    
    Load ft
    
    ft.Caption = m_Caption
    
    If m_Ftp Then
        ft.txtCode.OpenFile m_TempFile
        ft.LoadFunctions m_TempFile
        ft.CodeExp1.LoadCode m_TempFile
        ft.Tag = m_IdDoc
        ListaLangs.SetLang m_FileName, ft.txtCode
    Else
        ft.txtCode.OpenFile m_FileName
        ft.LoadFunctions m_FileName
        ft.CodeExp1.LoadCode m_FileName
        ft.Tag = m_IdDoc
        ListaLangs.SetLang m_FileName, ft.txtCode
        
        AddRecentFile m_FileName
    End If
    
    ft.Tag = m_IdDoc
    ft.Show
        
    OpenFile = True
    
    Exit Function
ErrorOpenFile:
    MsgBox "cFile.OpenFile : " & Err & " " & Error$, vbCritical
    
End Function


Public Function SaveFile(frm As Form, ByVal fSaveAs As Boolean, Optional ByVal SaveFtpDialog As Boolean = True) As Boolean
    
    Dim ret As Boolean
    
    'verificar si es un archivo nuevo o por actualizar
    If Len(m_FileName) > 0 Then
        'actualizar archivo con nuevo nombre o mantener
        If fSaveAs Then
            'preguntar como se va a llamar el archivo nuevo
            
            Files.filename = m_FileName
            
            If Files.GetSaveFile(m_FileName) Then
                On Error Resume Next
                
                If m_Ftp Then
                    m_Ftp = False
                    m_SiteName = vbNullString
                    m_User = vbNullString
                    m_Pwd = vbNullString
                    m_Port = 0
                    m_Url = vbNullString
                    m_TempFile = vbNullString
                    m_FtpFile = vbNullString
                End If
                
                Call frm.txtCode.SaveFile(m_FileName, False)
                If Err = 0 Then
                    'm_FileName = Files.FileName
                    m_Caption = util.StripFile(m_FileName)
                    frm.Caption = m_Caption
                    'actualizar funciones
                    frm.LoadFunctions m_FileName
                    
                    'actualizar explorador de codigo
                    frm.CodeExp1.Load m_FileName
                    
                    'actualizar explorador de codigo
                    frm.HTMLTree1.Load m_FileName, frm.txtCode.Text
                                        
                    'colorizar la nueva extension
                    ListaLangs.SetLang m_FileName, frm.txtCode
                    
                    With frmMain
                        .lvwOpeFiles.ListItems("k" & m_IdDoc).Text = m_Caption
                        .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(1) = PathArchivo(m_FileName)
                        .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(2) = vbNullString
                        .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(3) = vbNullString
                        .lblTot.Caption = CStr(.lvwOpeFiles.ListItems.count) & " documents"
                    End With
                    ret = True
                Else
                    MsgBox "Failed to save file", vbCritical
                End If
            End If
        Else
            'si se guarda archivo ftp entonces guardarlo en el servidor
            If m_Ftp Then
                
                If SaveFtpDialog Then
                    frmUpload.tipo_accion = 1
                    Load frmUpload
                    frmUpload.Show
                    frmUpload.lblFile.Caption = m_FileName
                    DoEvents
                End If
                
                'info del sitio
                FTPManager.SiteName = m_SiteName
                FTPManager.User = m_User
                FTPManager.pwd = m_Pwd
                FTPManager.Port = m_Port
                FTPManager.url = m_Url
                FTPManager.FtpFile = m_FtpFile
                
                'abrir sitio remoto
                If FTPManager.open_site2 Then
                    'colocar el directorio donde se saco el archivo
                    If FTPManager.set_current_directory(m_RemoteFolder) Then
                        'actualizar copia local
                        On Error Resume Next
                        Call frm.txtCode.SaveFile(m_TempFile, False)
                        If Err = 0 Then
                            'actualizar el archivo en el servidor
                            If FTPManager.put_file(m_TempFile, m_FileName) Then
                                'cerrar el sitio remoto
                                FTPManager.close_site
                                
                                m_Caption = "FTP:" & m_FileName
                                frm.Caption = m_Caption
                                
                                'actualizar funciones
                                frm.LoadFunctions m_TempFile
                                
                                'actualizar explorador de codigo
                                frm.CodeExp1.Load m_TempFile
                                
                                'actualizar explorador de codigo
                                frm.HTMLTree1.Load m_FileName, frm.txtCode.Text
                                 
                                'colorizar la nueva extension
                                ListaLangs.SetLang m_FileName, frm.txtCode 'm_TempFile, frm.txtCode
                                
                                With frmMain
                                    .lvwOpeFiles.ListItems("k" & m_IdDoc).Text = m_Caption
                                    .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(1) = vbNullString
                                    .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(2) = m_SiteName
                                    .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(3) = m_RemoteFolder
                                    .lblTot.Caption = CStr(.lvwOpeFiles.ListItems.count) & " documents"
                                End With
                                ret = True
                            Else
                                If SaveFtpDialog Then Unload frmUpload
                                MsgBox "Failed to create/update remote file : " & m_FileName, vbCritical
                            End If
                        Else
                            If SaveFtpDialog Then Unload frmUpload
                            MsgBox "Failed to save temp file.", vbCritical
                        End If
                    Else
                        If SaveFtpDialog Then Unload frmUpload
                        MsgBox "Failed to set remote folder : " & m_RemoteFolder, vbCritical
                    End If
                Else
                    If SaveFtpDialog Then Unload frmUpload
                    MsgBox "Failed to open sitename : " & m_SiteName, vbCritical
                End If
                
                If SaveFtpDialog Then Unload frmUpload
            Else
                'actualizar el contenido del archivo activo
                On Error Resume Next
                Call frm.txtCode.SaveFile(m_FileName, False)
                If Err = 0 Then
                    m_Caption = util.StripFile(m_FileName)
                    frm.Caption = m_Caption
                    'actualizar funciones
                    frm.LoadFunctions m_FileName
                    'actualizar explorador de codigo
                    frm.CodeExp1.Load m_FileName
                    'actualizar explorador de codigo
                    frm.HTMLTree1.Load m_FileName, frm.txtCode.Text
                    'colorizar la nueva extension
                    ListaLangs.SetLang m_FileName, frm.txtCode
                    
                    With frmMain
                        .lvwOpeFiles.ListItems("k" & m_IdDoc).Text = m_Caption
                        .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(1) = PathArchivo(m_FileName)
                        .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(2) = vbNullString
                        .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(3) = vbNullString
                        .lblTot.Caption = CStr(.lvwOpeFiles.ListItems.count) & " documents"
                    End With
                    ret = True
                Else
                    MsgBox "Failed to save file", vbCritical
                End If
            End If
        End If
    Else
        'preguntar como se va a llamar el archivo nuevo
        If m_Ftp Then
            'pantalla de espera
            If SaveFtpDialog Then
                frmUpload.tipo_accion = 1
                Load frmUpload
                frmUpload.Show
                frmUpload.lblFile.Caption = m_FileName
                DoEvents
            End If
            
            'info del sitio
            FTPManager.SiteName = m_SiteName
            FTPManager.User = m_User
            FTPManager.pwd = m_Pwd
            FTPManager.Port = m_Port
            FTPManager.url = m_Url
            FTPManager.FtpFile = m_FtpFile
            
            'abrir sitio remoto
            If FTPManager.open_site2 Then
                'colocar el directorio donde se saco el archivo
                If FTPManager.set_current_directory(m_RemoteFolder) Then
                    'actualizar copia local
                    On Error Resume Next
                    Call frm.txtCode.SaveFile(m_TempFile, False)
                    If Err = 0 Then
                        'actualizar el archivo en el servidor
                        If FTPManager.put_file(m_TempFile, m_FileName) Then
                            'cerrar el sitio remoto
                            FTPManager.close_site
                            
                            m_Caption = "FTP:" & m_FileName
                            frm.Caption = m_Caption
                            'actualizar funciones
                            frm.LoadFunctions m_TempFile
                            'actualizar explorador de codigo
                            frm.CodeExp1.Load m_TempFile
                            'actualizar explorador de codigo
                            frm.HTMLTree1.Load m_FileName, frm.txtCode.Text
                            'colorizar la nueva extension
                            ListaLangs.SetLang m_TempFile, frm.txtCode
                            
                            With frmMain
                                .lvwOpeFiles.ListItems("k" & m_IdDoc).Text = m_Caption
                                .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(1) = vbNullString
                                .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(2) = m_SiteName
                                .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(3) = m_RemoteFolder
                                .lblTot.Caption = CStr(.lvwOpeFiles.ListItems.count) & " documents"
                            End With
                                
                            ret = True
                        Else
                            If SaveFtpDialog Then Unload frmUpload
                            MsgBox "Failed to create/update remote file : " & m_FileName, vbCritical
                        End If
                    Else
                        If SaveFtpDialog Then Unload frmUpload
                        MsgBox "Failed to save temp file.", vbCritical
                    End If
                Else
                    If SaveFtpDialog Then Unload frmUpload
                    MsgBox "Failed to set remote folder : " & m_RemoteFolder, vbCritical
                End If
            Else
                If SaveFtpDialog Then Unload frmUpload
                MsgBox "Failed to open sitename : " & m_SiteName, vbCritical
            End If
            
            If SaveFtpDialog Then Unload frmUpload
        Else
            If Files.GetSaveFile(m_FileName) Then
                On Error Resume Next
                Call frm.txtCode.SaveFile(m_FileName, False)
                If Err = 0 Then
                    'm_FileName = Files.FileName
                    m_Caption = util.StripFile(m_FileName)
                    frm.Caption = m_Caption
                    'actualizar funciones
                    frm.LoadFunctions m_FileName
                    'actualizar explorador de codigo
                    frm.CodeExp1.Load m_FileName
                    'actualizar explorador de codigo
                    frm.HTMLTree1.Load m_FileName, frm.txtCode.Text
                    'colorizar la nueva extension
                    ListaLangs.SetLang m_FileName, frm.txtCode
                    
                    With frmMain
                        .lvwOpeFiles.ListItems("k" & m_IdDoc).Text = m_Caption
                        .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(1) = PathArchivo(m_FileName)
                        .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(2) = vbNullString
                        .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(3) = vbNullString
                        .lblTot.Caption = CStr(.lvwOpeFiles.ListItems.count) & " documents"
                    End With
                    
                    frm.Refresh
                    ret = True
                Else
                    MsgBox "Failed to save file : " & m_FileName, vbCritical
                End If
            End If
        End If
    End If
    
    SaveFile = ret
    
End Function
Public Property Get filename() As String
    filename = m_FileName
End Property

Public Property Let filename(ByVal pFileName As String)
    m_FileName = pFileName
End Property

Public Property Get Changed() As Boolean
    Changed = m_Changed
End Property

Public Property Let Changed(ByVal pChanged As Boolean)
    m_Changed = pChanged
End Property


Public Function SaveFile2(frm As Form, ByVal PathFile As String) As Boolean

    Dim ret As Boolean
    
    'verificar si es un archivo nuevo o por actualizar
    If Len(m_FileName) > 0 Then
        If m_Ftp Then
            m_FileName = PathFile & m_FileName
        Else
            m_FileName = PathFile & util.VBArchivoSinPath(m_FileName)
        End If
            
        On Error Resume Next
        
        If m_Ftp Then
            m_Ftp = False
            m_SiteName = vbNullString
            m_User = vbNullString
            m_Pwd = vbNullString
            m_Port = 0
            m_Url = vbNullString
            m_TempFile = vbNullString
            m_FtpFile = vbNullString
        End If
            
        Call frm.txtCode.SaveFile(m_FileName, False)
        If Err = 0 Then
            'm_FileName = Files.FileName
            m_Caption = util.StripFile(m_FileName)
            frm.Caption = m_Caption
            'actualizar funciones
            frm.LoadFunctions m_FileName
            'actualizar explorador de codigo
            frm.CodeExp1.Load m_FileName
            
            'actualizar explorador de codigo
            frm.HTMLTree1.Load m_FileName, frm.txtCode.Text
                    
            'colorizar la nueva extension
            ListaLangs.SetLang m_FileName, frm.txtCode
            
            With frmMain
                .lvwOpeFiles.ListItems("k" & m_IdDoc).Text = m_Caption
                .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(1) = PathArchivo(m_FileName)
                .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(2) = vbNullString
                .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(3) = vbNullString
                .lblTot.Caption = CStr(.lvwOpeFiles.ListItems.count) & " documents"
            End With
                    
            ret = True
        Else
            MsgBox "Failed to save file : " & m_FileName, vbCritical
        End If
    Else
        'ret = True
        If Files.GetSaveFile(m_FileName) Then
            On Error Resume Next
            Call frm.txtCode.SaveFile(m_FileName, False)
            If Err = 0 Then
                'm_FileName = Files.FileName
                m_Caption = util.StripFile(m_FileName)
                frm.Caption = m_Caption
                'actualizar funciones
                frm.LoadFunctions m_FileName
                'actualizar explorador de codigo
                frm.CodeExp1.Load m_FileName
                'actualizar explorador de codigo
                frm.HTMLTree1.Load m_FileName, frm.txtCode.Text
                'colorizar la nueva extension
                ListaLangs.SetLang m_FileName, frm.txtCode
                
                With frmMain
                    .lvwOpeFiles.ListItems("k" & m_IdDoc).Text = m_Caption
                    .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(1) = PathArchivo(m_FileName)
                    .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(2) = vbNullString
                    .lvwOpeFiles.ListItems("k" & m_IdDoc).SubItems(3) = vbNullString
                    .lblTot.Caption = CStr(.lvwOpeFiles.ListItems.count) & " documents"
                End With
            
                frm.Refresh
                ret = True
            Else
                MsgBox "Failed to save file : " & m_FileName, vbCritical
            End If
        End If
    End If
    
    SaveFile2 = ret

End Function

Public Property Get SiteName() As String
    SiteName = m_SiteName
End Property

Public Property Let SiteName(ByVal pSiteName As String)
    m_SiteName = pSiteName
End Property

Public Property Get url() As String
    url = m_Url
End Property

Public Property Let url(ByVal pUrl As String)
    m_Url = pUrl
End Property

Public Property Get User() As String
    User = m_User
End Property

Public Property Let User(ByVal pUser As String)
    m_User = pUser
End Property

Public Property Get pwd() As String
    pwd = m_Pwd
End Property

Public Property Let pwd(ByVal pPwd As String)
    m_Pwd = pPwd
End Property

Public Property Get Port() As Long
    Port = m_Port
End Property

Public Property Let Port(ByVal pPort As Long)
    m_Port = pPort
End Property

Public Property Get IdDoc() As Integer
    IdDoc = m_IdDoc
End Property

Public Property Let IdDoc(ByVal pIdDoc As Integer)
    m_IdDoc = pIdDoc
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal pCaption As String)
    m_Caption = pCaption
End Property

Public Property Get Ftp() As Boolean
    Ftp = m_Ftp
End Property

Public Property Let Ftp(ByVal pFtp As Boolean)
    m_Ftp = pFtp
End Property

Public Property Get TempFile() As String
    TempFile = m_TempFile
End Property

Public Property Let TempFile(ByVal pTempfile As String)
    m_TempFile = pTempfile
End Property

Public Property Get RemoteFolder() As String
    RemoteFolder = m_RemoteFolder
End Property

Public Property Let RemoteFolder(ByVal pRemoteFolder As String)
    m_RemoteFolder = pRemoteFolder
End Property

Public Property Get Passive() As Integer
    Passive = m_Passive
End Property

Public Property Let Passive(ByVal pPassive As Integer)
    m_Passive = pPassive
End Property

Public Property Get FtpFile() As String
    FtpFile = m_FtpFile
End Property

Public Property Let FtpFile(ByVal pFtpFile As String)
    m_FtpFile = pFtpFile
End Property
