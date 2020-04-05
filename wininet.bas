Attribute VB_Name = "wininet"
Option Explicit

Public hSession As Long
Public hConnect As Long
Public Const MAX_PATH = 260
Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As Currency
        ftLastAccessTime As Currency
        ftLastWriteTime As Currency
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public Const sReadBuffer = 1024

Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal pub_lngInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToWrite As Long, dwNumberOfBytesWritten As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Long
Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String, ByVal fdwAccess As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Long
Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Long
Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Long
Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Long
Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
Public Declare Function FtpCommand Lib "wininet.dll" Alias "FtpCommandA" (ByVal hFtpSession As Long, ByVal fExpectedReponse As Boolean, ByVal dwFlags As Long, ByVal lpszCommand As String, ByVal dwContext As Long, ByVal phFtpCommand As Long) As Boolean
                  
' Use registry access settings.
Public Const INTERNET_OPEN_TYPE_DIRECT = 1

' Type of service to access.
Public Const INTERNET_SERVICE_FTP = 1

' Brings the data across the wire even if it locally cached.
Public Const INTERNET_FLAG_RELOAD = &H80000000

Public Const FTP_TRANSFER_TYPE_ASCII = &H1

' flags for InternetOpen
Public Const INTERNET_FLAG_PASSIVE = &H8000000

Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000     ' don't write this item to the cache

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000


Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
                            (ByVal dwFlags As Long, lpSource As Any, _
                            ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
                            ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800

Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As Any) As Long

Const rDayZeroBias As Double = 109205#   ' Abs(CDbl(#01-01-1601#))
Const rMillisecondPerDay As Double = 10000000# * 60# * 60# * 24# / 10000#

Function Win32ToVbTime(ft As Currency) As Date
    Dim ftl As Currency
    ' Call API to convert from UTC time to local time
    If FileTimeToLocalFileTime(ft, ftl) Then
        ' Local time is nanoseconds since 01-01-1601
        ' In Currency that comes out as milliseconds
        ' Divide by milliseconds per day to get days since 1601
        ' Subtract days from 1601 to 1899 to get VB Date equivalent
        Win32ToVbTime = CDate((ftl / rMillisecondPerDay) - rDayZeroBias)
    Else
        MsgBox Err.LastDllError
    End If
End Function

Public Function ReturnSize(file As String) As Long
  Dim hFile As Long, dt As WIN32_FIND_DATA
  hFile = FtpFindFirstFile(hConnect, file, dt, INTERNET_FLAG_RELOAD, INTERNET_FLAG_NO_CACHE_WRITE)
  If hFile = 0 Then
    ReturnSize = 0
    Exit Function
  End If
  ReturnSize = dt.nFileSizeLow
  InternetCloseHandle hFile
End Function

Public Function GetFTPDirectory(hConnection As Long) As String
    Dim szDir As String
    szDir = String(1024, Chr$(0))
    If (FtpGetCurrentDirectory(hConnection, szDir, 1024) = False) Then
        Exit Function
    Else
        GetFTPDirectory = VBA.Left(szDir, InStr(1, szDir, String(1, 0), vbBinaryCompare) - 1)
    End If
End Function

Public Sub FTPError(ByVal dwError As Long, ByRef szFunc As String)
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
    MsgBox szErrorMessage
End Sub


