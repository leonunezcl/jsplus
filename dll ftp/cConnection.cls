VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'WinFTP, created by the KPD-Team 2000
'This file can be downloaded from http://www.allapi.net/
'For questions or comments, contact us at KPDTeam@Allapi.net

' You are free to use this code within your own applications,
' but you are expressly forbidden from selling or otherwise
' distributing this source code without prior written consent.
' This includes both posting free demo projects made from this
' code as well as reproducing the code in text or html format.

Const scUserAgent = "KPD Connection Class"
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_TEMPORARY = &H100
Const FILE_ATTRIBUTE_COMPRESSED = &H800
Const FILE_ATTRIBUTE_OFFLINE = &H1000
Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Const INTERNET_SERVICE_FTP = 1
Const INTERNET_INVALID_PORT_NUMBER = 0
Const MIN_SOCKETS_REQD As Long = 1
Const WS_VERSION_REQD As Long = &H101
Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Const SOCKET_ERROR As Long = -1
Const WSADESCRIPTION_LEN = 257
Const WSASYS_STATUS_LEN = 129
Const MAX_WSADescription = 256
Const MAX_WSASYSStatus = 128
Const MAX_PATH = 260
Const NO_ERROR = 0
Const ERROR_SUCCESS = 0
Const ERROR_NO_MORE_FILES = 18
Const INTERNET_FLAG_PASSIVE = &H8000000
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
Private Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type
Private Type WSADataInfo
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As String
End Type
Private Type Hostent
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type
Private Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type
Private Type IP_ECHO_REPLY
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    data As Long
    Options As IP_OPTION_INFORMATION
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Private Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean
Private Declare Function WSAStartupInfo Lib "WSOCK32" Alias "WSAStartup" (ByVal wVersionRequested As Integer, lpWSAdata As WSADataInfo) As Long
Private Declare Function WSACleanup Lib "WSOCK32" () As Long
Private Declare Function WSAGetLastError Lib "WSOCK32" () As Long
Private Declare Function WSAStartup Lib "WSOCK32" (ByVal wVersionRequired As Long, lpWSAdata As WSAdata) As Long
Private Declare Function gethostname Lib "WSOCK32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function GetHostByName Lib "WSOCK32" Alias "gethostbyname" (ByVal szHost As String) As Long
Private Declare Sub CopyMemoryIP Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetAttemptConnect Lib "wininet" (ByVal dwReserved As Long) As Long
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean
Public hInternet As Long
Public hConnection As Long
Public ServerName As String
Public Username As String
Public Password As String
Public Connected As Boolean
Public FTPService As Boolean
Event StatusChanged(NewStatus As tNewStatus, sOptionalInfo As String)
Public Function CreateConnection(bFTPService As Boolean, Optional sServerName As String, Optional sUserName As String, Optional sPassword As String) As Boolean
    If Connected Then Disconnect
    If InternetAttemptConnect(0) <> 0 Then Exit Function
    hInternet = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hInternet = 0 Then
        RaiseEvent StatusChanged(nsConnectionFailed, "Unable to connect to the Internet.")
        Exit Function
    End If
    If bFTPService Then
        hConnection = InternetConnect(hInternet, sServerName, INTERNET_INVALID_PORT_NUMBER, sUserName, sPassword, INTERNET_SERVICE_FTP, IIf(frmConnect.chkPassive.Value = vbChecked, INTERNET_FLAG_PASSIVE, 0), 0)
        If hConnection = 0 Then
            Disconnect
            RaiseEvent StatusChanged(nsConnectionFailed, "Unable to connect to the specified server.")
            Exit Function
        End If
        ServerName = sServerName
        Username = sUserName
        Password = sPassword
    End If
    FTPService = bFTPService
    Connected = True
    CreateConnection = True
    RaiseEvent StatusChanged(nsConnected, "")
End Function
Public Sub Disconnect()
    InternetCloseHandle hConnection
    InternetCloseHandle hInternet
    Connected = False
    RaiseEvent StatusChanged(nsDisconnected, "")
End Sub
Public Function CanConnect() As Boolean
    CanConnect = Not (InternetAttemptConnect(0))
End Function
Public Function GetLocalIPAddress() As String
    Dim sHostName As String * 256
    Dim lpHost As Long
    Dim Host As Hostent
    Dim dwIPAddr As Long
    Dim tmpIPAddr() As Byte
    Dim I As Integer
    Dim sIPAddr As String
    If Not SocketsInitialize() Then
        GetLocalIPAddress = ""
        Exit Function
    End If
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetLocalIPAddress = ""
        MsgBox "Windows Sockets error " & str$(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    sHostName = Trim$(sHostName)
    lpHost = GetHostByName(sHostName)
    If lpHost = 0 Then
        GetLocalIPAddress = ""
        MsgBox "Windows Sockets are not responding. " & "Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    CopyMemoryIP Host, lpHost, Len(Host)
    CopyMemoryIP dwIPAddr, Host.hAddrList, 4
    ReDim tmpIPAddr(1 To Host.hLen)
    CopyMemoryIP tmpIPAddr(1), dwIPAddr, Host.hLen
    For I = 1 To Host.hLen
        sIPAddr = sIPAddr & tmpIPAddr(I) & "."
    Next
    GetLocalIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    SocketsCleanup
End Function
Public Function GetLocalIPHostName() As String
    Dim sHostName As String * 256
    If Not SocketsInitialize() Then
        GetLocalIPHostName = ""
        Exit Function
    End If
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetLocalIPHostName = ""
        MsgBox "Windows Sockets error " & str$(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    GetLocalIPHostName = VBA.Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup
End Function
Private Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function
Private Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function
Private Sub SocketsCleanup()
    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
End Sub
Private Function SocketsInitialize() As Boolean
    Dim WSAD As WSAdata
    Dim sLoByte As String
    Dim sHiByte As String
    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
        MsgBox "The 32-bit Windows Socket is not responding."
        SocketsInitialize = False
        Exit Function
    End If
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & CStr(MIN_SOCKETS_REQD) & " supported sockets."
        SocketsInitialize = False
        Exit Function
    End If
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        sHiByte = CStr(HiByte(WSAD.wVersion))
        sLoByte = CStr(LoByte(WSAD.wVersion))
        MsgBox "Sockets version " & sLoByte & "." & sHiByte & " is not supported by 32-bit Windows Sockets."
        SocketsInitialize = False
        Exit Function
    End If
    'must be OK, so lets do it
    SocketsInitialize = True
End Function
Public Function PingServer(HostName As String, ByRef PingTime As Long, ByRef RemoteIP As String) As Boolean
    Dim hFile As Long, lpWSAdata As WSAdata
    Dim hHostent As Hostent, AddrList As Long
    Dim Address As Long, rIP As String
    Dim OptInfo As IP_OPTION_INFORMATION
    Dim EchoReply As IP_ECHO_REPLY
    RaiseEvent StatusChanged(nsQuerySent, "")
    Call WSAStartup(&H101, lpWSAdata)
    If GetHostByName(HostName + String(64 - Len(HostName), 0)) <> SOCKET_ERROR Then
        CopyMemory hHostent.hName, ByVal GetHostByName(HostName + String(64 - Len(HostName), 0)), Len(hHostent)
        CopyMemory AddrList, ByVal hHostent.hAddrList, 4
        CopyMemory Address, ByVal AddrList, 4
    End If
    hFile = IcmpCreateFile()
    If hFile = 0 Then
        RaiseEvent StatusChanged(nsQueryFailed, "Error while creating Internet Control Message Protocol handle.")
        Exit Function
    End If
    OptInfo.TTL = 255
    If IcmpSendEcho(hFile, Address, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, 2000) Then
        rIP = CStr(EchoReply.Address(0)) + "." + CStr(EchoReply.Address(1)) + "." + CStr(EchoReply.Address(2)) + "." + CStr(EchoReply.Address(3))
    End If
    If EchoReply.Status = 0 Then
        RemoteIP = rIP
        PingTime = EchoReply.RoundTripTime
        PingServer = True
        RaiseEvent StatusChanged(nsQueryCompleted, "")
    Else
        RaiseEvent StatusChanged(nsQueryFailed, "No reply received.")
    End If
    Call IcmpCloseHandle(hFile)
    Call WSACleanup
End Function
Public Function SetNewDirectory(sNewDir As String)
    If FTPService = False Or Connected = False Then
        RaiseEvent StatusChanged(nsOperationFailed, "Not an FTP session.")
        Exit Function
    End If
    RaiseEvent StatusChanged(nsOperationStarted, "")
    If FtpSetCurrentDirectory(hConnection, sNewDir) = False Then
        RaiseEvent StatusChanged(nsOperationFailed, "Error while moving to the new directory")
        Exit Function
    End If
    RaiseEvent StatusChanged(nsOperationSuceeded, "")
End Function
Public Function ClearCollection(cCol As Collection)
    While cCol.Count > 0
        cCol.Remove cCol.Count
    Wend
End Function
Public Function EnumFiles(EnumItemNameBag As Collection, EnumItemAttributeBag As Collection, EnumItemSizeBag As Collection) As Boolean
    If FTPService = False Or Connected = False Then
        RaiseEvent StatusChanged(nsOperationFailed, "Not an FTP session.")
        Exit Function
    End If
    Dim hFind As Long, nLastError As Long, ptr As Long, pData As WIN32_FIND_DATA
    Dim bRet As Boolean, strItemName As String, Tel As Long
    ClearCollection EnumItemNameBag
    ClearCollection EnumItemAttributeBag
    ClearCollection EnumItemSizeBag
    pData.cFileName = String(MAX_PATH, 0)
    hFind = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)
    nLastError = Err.LastDllError
    If hFind = 0 Then
        If Err.LastDllError = ERROR_NO_MORE_FILES Then
            RaiseEvent StatusChanged(nsOperationSuceeded, "")
        Else
            RaiseEvent StatusChanged(nsOperationFailed, "Error while enumerating the remote files.")
        End If
        Exit Function
    End If
    EnumItemAttributeBag.Add pData.dwFileAttributes
    strItemName = VBA.Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
    EnumItemNameBag.Add strItemName
    EnumItemSizeBag.Add pData.nFileSizeLow
    Do
        Tel = Tel + 1
        frmMainFTP.lblNumFiles.Caption = CStr(Tel)
        DoEvents
        pData.cFileName = String(MAX_PATH, 0)
        bRet = InternetFindNextFile(hFind, pData)
        If bRet = False Then
            If Err.LastDllError = ERROR_NO_MORE_FILES Then
                RaiseEvent StatusChanged(nsOperationSuceeded, "")
            Else
                RaiseEvent StatusChanged(nsOperationFailed, "Error while enumerating the remote files.")
            End If
            InternetCloseHandle hFind
            Exit Function
        Else
            EnumItemAttributeBag.Add pData.dwFileAttributes
            strItemName = VBA.Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
            EnumItemNameBag.Add strItemName
            EnumItemSizeBag.Add pData.nFileSizeLow
       End If
    Loop
End Function
Public Function CreateDirectory(sNewDir As String) As Boolean
    If FTPService = False Or Connected = False Then
        RaiseEvent StatusChanged(nsOperationFailed, "Not an FTP session.")
        Exit Function
    End If
    RaiseEvent StatusChanged(nsOperationStarted, "")
    If FtpCreateDirectory(hConnection, sNewDir) = False Then
        RaiseEvent StatusChanged(nsOperationFailed, "Error while creating the new directory.")
        Exit Function
    End If
    RaiseEvent StatusChanged(nsOperationSuceeded, "")
End Function
Public Function RemoveDirectory(sDir As String) As Boolean
    If FTPService = False Or Connected = False Then
        RaiseEvent StatusChanged(nsOperationFailed, "Not an FTP session.")
        Exit Function
    End If
    RaiseEvent StatusChanged(nsOperationStarted, "")
    If FtpRemoveDirectory(hConnection, sDir) = False Then
        RaiseEvent StatusChanged(nsOperationFailed, "Error while deleteing the directory.")
        Exit Function
    End If
    RaiseEvent StatusChanged(nsOperationSuceeded, "")
End Function
Public Function GetCurrentDirectory() As String
    If FTPService = False Or Connected = False Then
        RaiseEvent StatusChanged(nsOperationFailed, "Not an FTP session.")
        Exit Function
    End If
    GetCurrentDirectory = String(1024, 0)
    If FtpGetCurrentDirectory(hConnection, GetCurrentDirectory, 1024) = 0 Then
        RaiseEvent StatusChanged(nsOperationFailed, "Unable to retrieve the current directory.")
        Exit Function
    Else
        GetCurrentDirectory = VBA.Left$(GetCurrentDirectory, InStr(1, GetCurrentDirectory, Chr$(0)) - 1)
    End If
    RaiseEvent StatusChanged(nsOperationSuceeded, "")
End Function
Public Function GetLastResponseInfo() As String
    Dim nErr As Long, BuffSize As Long
    InternetGetLastResponseInfo nErr, GetLastResponseInfo, BuffSize
    If BuffSize > 0 Then
        GetLastResponseInfo = Space(BuffSize)
        InternetGetLastResponseInfo nErr, GetLastResponseInfo, BuffSize
    End If
End Function
