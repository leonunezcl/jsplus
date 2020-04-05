VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFtp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'handlers de la coneccion
Private hConnection As Long
Private hOpen As Long
Private sOrgPath  As String
Private FTPInfo As Ftp

Private m_SiteName As String
Private m_Url As String
Private m_Port As Integer
Private m_User As String
Private m_Pwd As String
Private m_Passive As Integer
Private m_lastdir As String
Private m_str_error As String
Private m_File As String

Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000     ' don't write this item to the cache
Private Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
Private Const FTP_TRANSFER_TYPE_ASCII = &H1
Private Const FTP_TRANSFER_TYPE_BINARY = &H2
Private Const INTERNET_DEFAULT_FTP_PORT = 21               ' default for FTP servers
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_FLAG_PASSIVE = &H8000000            ' used for FTP connections
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0                    ' use registry configuration
Private Const INTERNET_OPEN_TYPE_DIRECT = 1                        ' direct to net
Private Const INTERNET_OPEN_TYPE_PROXY = 3                         ' via named proxy
Private Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY = 4   ' prevent using java/script/INS
Private Const MAX_PATH = 260

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hConnect As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Boolean
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpCommand Lib "wininet.dll" Alias "FtpCommandA" (ByVal hFtpSession As Long, ByVal fExpectedReponse As Boolean, ByVal dwFlags As Long, ByVal lpszCommand As String, ByVal dwContext As Long, ByVal phFtpCommand As Long) As Boolean
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long

Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
                            (ByVal dwFlags As Long, lpSource As Any, _
                            ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
                            ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
'Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As Any) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

Private Const rDayZeroBias As Double = 109205#   ' Abs(CDbl(#01-01-1601#))
Private Const rMillisecondPerDay As Double = 10000000# * 60# * 60# * 24# / 10000#

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Public Function execute_command(ByVal sCommand As String, ByVal sFile As String) As Boolean

    Dim ret As Long
    
    ret = FtpCommand(hConnection, False, FTP_TRANSFER_TYPE_ASCII, sCommand & " " & sFile, 0, 0)
    
    execute_command = IIf(ret, True, False)
    
End Function


Public Function open_site3() As Boolean

    'Dim ret As String
    Dim lErr As Long
    Dim sErr As String
    Dim lenBuf As Long
    
    'm_SiteName = FTPManager.SiteName
    'm_User = FTPManager.User
    'm_Pwd = FTPManager.Pwd
    'm_Url = FTPManager.Url
    'm_Port = FTPManager.Port
    'm_lastdir = FTPManager.lastdir
        
    'cerrar conexiones anteriores
    If hOpen <> 0 Then Call close_site
    
    'open an internet connection
    hOpen = InternetOpen(SiteName, INTERNET_OPEN_TYPE_DIRECT, "", "", INTERNET_FLAG_NO_CACHE_WRITE)
    
    If hOpen <> 0 Then
        'connect to the FTP server
        hConnection = InternetConnect(hOpen, m_Url, m_Port, m_User, m_Pwd, INTERNET_SERVICE_FTP, IIf(m_Passive, INTERNET_FLAG_PASSIVE, 0), &H0)
                    
        If hConnection <> 0 Then
            If Len(m_lastdir) > 0 Then
                'get the directory
                If Not set_current_directory(m_lastdir) Then
                    m_lastdir = get_current_directory()
                    set_current_directory (m_lastdir)
                End If
            Else
                m_lastdir = get_current_directory()
                set_current_directory (m_lastdir)
            End If
        Else
            'get the required buffer size
            InternetGetLastResponseInfo lErr, sErr, lenBuf
            'create a buffer
            sErr = String(lenBuf, 0)
            'retrieve the last respons info
            InternetGetLastResponseInfo lErr, sErr, lenBuf
            
            Exit Function
        End If
    Else
        'get the required buffer size
        InternetGetLastResponseInfo lErr, sErr, lenBuf
        'create a buffer
        sErr = String(lenBuf, 0)
        'retrieve the last respons info
        InternetGetLastResponseInfo lErr, sErr, lenBuf
        
        Exit Function
    End If
    
    open_site3 = True
    

End Function

Public Function open_from_web(ByVal sUrl As String, ByRef tmpfile As String) As Boolean

    On Error GoTo Erroropen_from_web
    
    Dim hOpen As Long, hFile As Long, sBuffer As String, ret As Long
    Dim src As New cStringBuilder
    Dim nFreeFile As Long
    Dim tmpdir As String
        
    'Create a buffer for the file we're going to download
    'sBuffer = Space$(1024)
    
    'Create an internet connection
    hOpen = InternetOpen("openweb", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    
    'Open the url
    hFile = InternetOpenUrl(hOpen, sUrl, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
    
    Do
        If frmOpenFiles.Cancelo Then Exit Do
        
        sBuffer = Space$(4096)
        
        InternetReadFile hFile, sBuffer, 4096, ret
    
        If ret > 0 Then
            src.Append sBuffer
        End If
        DoEvents
    Loop Until ret <= 0
  
    'clean up
    InternetCloseHandle hFile
    InternetCloseHandle hOpen

    'archivo de salida
    tmpfile = util.ArchivoTemporal()
    tmpdir = util.StripPath(App.Path) & "temp"
    util.CrearDirectorio tmpdir
    tmpfile = tmpdir & "\" & util.StripFile(tmpfile)
    nFreeFile = FreeFile
    
    Open tmpfile For Output As #nFreeFile
        Print #1, src.ToString
    Close #nFreeFile
    
    open_from_web = True
        
    sBuffer = vbNullString
    
    Exit Function
    
Erroropen_from_web:
    MsgBox "open_from_web : " & Err & " " & Error$, vbCritical
    
End Function
Public Sub close_site(Optional ByVal updaccount As Boolean = True)

    If hOpen = 0 Then Exit Sub
    
    'actualizar info de la cuenta
    If updaccount Then upd_info_account
    
    'close the FTP connection
    InternetCloseHandle hConnection
    
    'close the internet connection
    InternetCloseHandle hOpen

    hConnection = 0
    hOpen = 0
    
End Sub

Public Function get_current_directory() As Boolean

    Dim ret As Long
    Dim buffer As String
    
    buffer = Space$(255)
    
    ret = FtpGetCurrentDirectory(hConnection, buffer, Len(buffer))
        
    get_current_directory = IIf(ret, True, False)
    
    If Len(Trim$(buffer)) > 0 Then
        m_lastdir = util.StripNulls(buffer)
    End If
    
End Function

Public Sub get_files_from_dir(ByVal filter As String, ByRef arr_files() As String)
    
    Dim pData As WIN32_FIND_DATA, hFind As Long, lRet As Long
    Dim C As Integer
    Dim File As String
    Dim str As New cStringBuilder
    
    ReDim arr_files(0)
    
    pData.cFileName = String(MAX_PATH, 0)

    hFind = FtpFindFirstFile(hConnection, filter, pData, 0, 0)
    
    If hFind = 0 Then Exit Sub
    
    C = 1
    ReDim Preserve arr_files(C)
    
    File = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
    str.Append File
        
    If (pData.dwFileAttributes And vbDirectory) Then
        str.Append "|D|"
    Else
        str.Append "|F|"
    End If
    
    str.Append CStr(pData.nFileSizeLow) & "|"
    str.Append CStr(Win32ToVbTime(pData.ftCreationTime)) & "|"
    str.Append CStr(Win32ToVbTime(pData.ftLastWriteTime)) & "|"
        
    arr_files(C) = str.ToString
    
    C = C + 1
    
    Do
        pData.cFileName = String(MAX_PATH, 0)

        lRet = InternetFindNextFile(hFind, pData)

        If lRet = 0 Then Exit Do
        
        ReDim Preserve arr_files(C)
        File = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
        str.Append File
        
        If (pData.dwFileAttributes And vbDirectory) Then
            str.Append "|D|"
        Else
            str.Append "|F|"
        End If
        
        str.Append CStr(pData.nFileSizeLow) & "|"
        str.Append CStr(Win32ToVbTime(pData.ftCreationTime)) & "|"
        str.Append CStr(Win32ToVbTime(pData.ftLastWriteTime)) & "|"
                
        arr_files(C) = str.ToString
        
        Set str = Nothing
        
        C = C + 1
    Loop
    
    InternetCloseHandle hFind
    
End Sub

Public Function OpenSiteInformation(ByVal File As String) As Boolean

    On Error GoTo ErrorOpenSiteInformation
    
    Dim fFile As Long
    Dim lErr As Long
    Dim sErr As String
    Dim lenBuf As Long
    
    If Not ArchivoExiste2(File) Then
        MsgBox "File not found : " & File, vbCritical
        Exit Function
    End If
        
    fFile = FreeFile()
        
    Open File For Binary Access Read As #fFile
        Get #fFile, , FTPInfo
    Close #fFile
    
    m_SiteName = FTPInfo.Name
    m_User = FTPInfo.UserName
    m_Pwd = Base64Decode(FTPInfo.Password)
    m_Url = FTPInfo.url
    m_Port = FTPInfo.PortNum
    m_lastdir = FTPInfo.lastdir
    m_File = File
    
    OpenSiteInformation = True
    
    Exit Function
ErrorOpenSiteInformation:
    If fFile > 0 Then Close #fFile
    MsgBox "OpenSiteInformation : " & Err.Number & " " & Err.description, vbCritical
    
End Function
Private Function Win32ToVbTime(ft1 As FILETIME)
    
    Dim Ft2 As FILETIME
    Dim SysTime As SYSTEMTIME
    
    ' convertir el formato de fecha UTC al formato de fecha local
    If FileTimeToLocalFileTime(Ft2, ft1) Then
        'llamar api y formatear
        FileTimeToSystemTime ft1, SysTime

        'regresar la fecha convertida
        Win32ToVbTime = Format(LTrim(str$(SysTime.wDay)), "00") + "/" + Format(str$(SysTime.wMonth), "00") + "/" + LTrim(str$(SysTime.wYear))
    Else
        MsgBox Err.LastDllError
    End If
    
End Function
Public Function delete_file(ByVal File As String) As Boolean

    Dim ret As Long
    
    'delete the file from the FTP server
    ret = FtpDeleteFile(hConnection, File)

    delete_file = IIf(ret, True, False)
    
End Function

Public Function get_file(ByVal input_file As String, ByVal output_file As String) As Boolean

    Dim ret As Long
    
    'retrieve the file from the FTP server
    ret = FtpGetFile(hConnection, input_file, output_file, False, 0, FTP_TRANSFER_TYPE_ASCII, 0)
    
    get_file = IIf(ret, True, False)
    
End Function

Public Function get_binary_file(ByVal input_file As String, ByVal output_file As String) As Boolean

    Dim ret As Long
    
    'retrieve the file from the FTP server
    ret = FtpGetFile(hConnection, input_file, output_file, False, 0, FTP_TRANSFER_TYPE_BINARY, 0)
    
    get_binary_file = IIf(ret, True, False)
    
End Function


Public Function make_directory(ByVal Folder As String) As Boolean
    
    Dim ret As Long
    
    'create a new directory 'testing'
    ret = FtpCreateDirectory(hConnection, Folder)
    
    make_directory = IIf(ret, True, False)
    
End Function

Public Function open_site(ByVal File As String) As Boolean

    'Dim ret As String
    Dim fFile As Long
    Dim lErr As Long, sErr As String, lenBuf As Long
    
    If Not ArchivoExiste2(File) Then
        MsgBox "File not found : " & File, vbCritical
        Exit Function
    End If
        
    fFile = FreeFile()
        
    On Error Resume Next
    Open File For Binary Access Read As #fFile
        Get #fFile, , FTPInfo
    Close #fFile
    
    If Err <> 0 Then
        open_site = False
        Exit Function
    End If
    
    m_SiteName = FTPInfo.Name
    m_User = FTPInfo.UserName
    m_Pwd = Base64Decode(FTPInfo.Password)
    m_Url = FTPInfo.url
    m_Port = FTPInfo.PortNum
    m_lastdir = FTPInfo.lastdir
    m_File = File
    
    'cerrar conexiones anteriores
    If hOpen <> 0 Then Call close_site
    
    'open an internet connection
    hOpen = InternetOpen(SiteName, INTERNET_OPEN_TYPE_DIRECT, "", "", INTERNET_FLAG_NO_CACHE_WRITE)
    
    If hOpen <> 0 Then
        'connect to the FTP server
        hConnection = InternetConnect(hOpen, m_Url, m_Port, m_User, m_Pwd, INTERNET_SERVICE_FTP, IIf(m_Passive, INTERNET_FLAG_PASSIVE, 0), &H0)
                    
        If hConnection <> 0 Then
            If Len(m_lastdir) > 0 Then
                'get the directory
                If Not set_current_directory(m_lastdir) Then
                    m_lastdir = get_current_directory()
                    set_current_directory (m_lastdir)
                End If
            Else
                m_lastdir = get_current_directory()
                set_current_directory (m_lastdir)
            End If
        Else
            'get the required buffer size
            InternetGetLastResponseInfo lErr, sErr, lenBuf
            'create a buffer
            sErr = String(lenBuf, 0)
            'retrieve the last respons info
            InternetGetLastResponseInfo lErr, sErr, lenBuf
            
            Exit Function
        End If
    Else
        'get the required buffer size
        InternetGetLastResponseInfo lErr, sErr, lenBuf
        'create a buffer
        sErr = String(lenBuf, 0)
        'retrieve the last respons info
        InternetGetLastResponseInfo lErr, sErr, lenBuf
        
        Exit Function
    End If
    
    open_site = True
    
End Function

Public Function open_site2() As Boolean

    'Dim ret As String
    Dim lErr As Long
    Dim sErr As String
    Dim lenBuf As Long
    
    m_SiteName = FTPManager.SiteName
    m_User = FTPManager.User
    m_Pwd = FTPManager.pwd
    m_Url = FTPManager.url
    m_Port = FTPManager.Port
    m_lastdir = FTPManager.lastdir
        
    'cerrar conexiones anteriores
    If hOpen <> 0 Then Call close_site
    
    'open an internet connection
    hOpen = InternetOpen(SiteName, INTERNET_OPEN_TYPE_DIRECT, "", "", INTERNET_FLAG_NO_CACHE_WRITE)
    
    If hOpen <> 0 Then
        'connect to the FTP server
        hConnection = InternetConnect(hOpen, m_Url, m_Port, m_User, m_Pwd, INTERNET_SERVICE_FTP, IIf(m_Passive, INTERNET_FLAG_PASSIVE, 0), &H0)
                    
        If hConnection <> 0 Then
            If Len(m_lastdir) > 0 Then
                'get the directory
                If Not set_current_directory(m_lastdir) Then
                    m_lastdir = get_current_directory()
                    set_current_directory (m_lastdir)
                End If
            Else
                m_lastdir = get_current_directory()
                set_current_directory (m_lastdir)
            End If
        Else
            'get the required buffer size
            InternetGetLastResponseInfo lErr, sErr, lenBuf
            'create a buffer
            sErr = String(lenBuf, 0)
            'retrieve the last respons info
            InternetGetLastResponseInfo lErr, sErr, lenBuf
            
            Exit Function
        End If
    Else
        'get the required buffer size
        InternetGetLastResponseInfo lErr, sErr, lenBuf
        'create a buffer
        sErr = String(lenBuf, 0)
        'retrieve the last respons info
        InternetGetLastResponseInfo lErr, sErr, lenBuf
        
        Exit Function
    End If
    
    open_site2 = True
    
End Function


Public Sub ftp_error(ByVal dwError As Long, ByRef szFunc As String)
    Dim dwTemp As Long
    Dim szString As String * 2048, szErrorMessage As String
    FormatMessage FORMAT_MESSAGE_FROM_HMODULE, _
                      GetModuleHandle("wininet.dll"), dwError, 0, _
                      szString, 256, 0
    szErrorMessage = szFunc & " error code: " & dwError & " Message: " & szString
    If (dwError = 12003) Then
        ' Extended error information was returned
        InternetGetLastResponseInfo dwTemp, szString, 2048
        szErrorMessage = szString
    End If
    MsgBox szErrorMessage, vbCritical
End Sub



Private Sub upd_info_account()
    
    FTPInfo.Name = m_SiteName
    FTPInfo.UserName = m_User
    FTPInfo.url = m_Url
    FTPInfo.PortNum = m_Port
    FTPInfo.Password = Base64Encode(m_Pwd)
    FTPInfo.lastdir = m_lastdir
    
    If Len(m_File) > 0 Then
        Open m_File For Binary Access Write As #1
            Put #1, , FTPInfo
        Close #1
    End If
        
End Sub


Public Function put_file(ByVal input_file As String, output_file As String)

    Dim ret As Long
    
    'upload the file
    ret = FtpPutFile(hConnection, input_file, output_file, FTP_TRANSFER_TYPE_ASCII, 0)
    
    put_file = IIf(ret, True, False)

End Function

Public Function remove_directory(ByVal Folder As String) As Boolean

    Dim ret As Long

    'remove the direcrtory 'testing'
    ret = FtpRemoveDirectory(hConnection, Folder)

    remove_directory = IIf(ret, True, False)
    
End Function

Public Function rename_file(ByVal older_name As String, ByVal new_name As String) As Boolean

    Dim ret As Long
    
    'rename
    ret = FtpRenameFile(hConnection, older_name, new_name)

    rename_file = IIf(ret, True, False)
    
End Function


Public Function set_current_directory(ByVal Folder As String) As Boolean

    Dim ret As Long
    
    'set the current directory to 'root/testing'
    ret = FtpSetCurrentDirectory(hConnection, Folder)

    set_current_directory = IIf(ret, True, False)
    
    If ret Then
        m_lastdir = Folder
    End If
    
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

Public Property Get Port() As Long
    Port = m_Port
End Property

Public Property Let Port(ByVal pPort As Long)
    m_Port = pPort
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

Public Property Get Passive() As Integer
    Passive = m_Passive
End Property

Public Property Let Passive(ByVal ppasive As Integer)
    m_Passive = ppasive
End Property

Private Sub Class_Initialize()
    hConnection = 0
    hOpen = 0
    sOrgPath = ""
End Sub


Private Sub Class_Terminate()
    hConnection = 0
    hOpen = 0
    sOrgPath = ""
End Sub



Public Property Get lastdir() As String
    lastdir = m_lastdir
End Property

Public Property Let lastdir(ByVal plastdir As String)
    m_lastdir = plastdir
End Property

Public Property Get str_error() As String
    str_error = m_str_error
End Property

Public Property Let str_error(ByVal pstr_error As String)
    m_str_error = pstr_error
End Property

Public Property Get FtpFile() As String
    FtpFile = m_File
End Property

Public Property Let FtpFile(ByVal pFile As String)
    m_File = pFile
End Property
