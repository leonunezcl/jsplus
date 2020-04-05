Attribute VB_Name = "Module2"
Option Explicit

Public Cdlg As New cCommonDialog
Public util As New cLibrary
Public Plugins As New cPlugins
Public FTPManager As New cFtp
Public CEventos As New CEventos
Public Files As New cFiles
Public SeachFiles As New cSearchFiles
Public CComHtmlAttrib As New CComHtml
Public HTidy As New cTydy
Public ListaLangs As New cLanguage
Public ProjectMan As New cproject
Private str1 As New cStringBuilder
Public glbquickon As Boolean

Public LastPath As String
Public Result As String
Public started As Boolean
Private arr_files(10) As String
Public HTML As String
Private Type eHtmlAttrib
    attribute As String
    help As String
    tipo As String
    icono As Integer
End Type

Private Type eHtml
    Tag As String
    HTML As String
    help As String
    elems() As eHtmlAttrib
End Type

Public arr_html() As eHtml

Private Type eData
    Tag As String
    help As String
End Type
Public arr_data_css() As eData

Private Type eDataXml
    Tag As String
    help As String
End Type
Public arr_data_xml() As eDataXml

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Const SEM_NOGPFAULTERRORBOX = &H2
Private m_bInIDE As Boolean

Private Const C_INI = "jsplus.ini"

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GWL_STYLE = (-16)
' Header control styles
Private Const HDS_BUTTONS = &H2

Private Const LVM_FIRST = &H1000& ' ListView messages
Private Const LVM_GETHEADER = (LVM_FIRST + 31)

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Type Ftp
    Name As String
    url As String
    UserName As String
    Password As String
    Anonymous As Integer
    PortNum As Integer
    lastdir As String
    Passive As Integer
End Type

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type
Private Const FO_DELETE = &H3
Private Const FOF_NOCONFIRMATION = &H10
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long



Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
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
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Const OFS_MAXPATHNAME = 256
Private Const OF_READ = &H0

Private Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function setfiletime Lib "kernel32" Alias "SetFileTime" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

'1=desarrollador
'2=comercial
'3=educacional
'4=especial
Public Const tipo_version = 1
Public Const debug_output As Integer = 1
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const MAX_PATH = 260

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

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_SHOWWINDOW = &H40

Private bufferini As String

Private Declare Function GetLongPathName Lib "kernel32.dll" _
                 Alias "GetLongPathNameA" ( _
                 ByVal lpszShortPath As String, _
                 ByVal lpszLongPath As String, _
                 ByVal cchBuffer As Long) As Long

Public Sub AddRecentFile(ByVal Archivo As String)
    
    On Error GoTo SkipError
    
    If Not ArchivoExiste2(Archivo) Then Exit Sub

    Dim arrFiles() As String
    Dim total As Long
    Dim k As Integer
    Dim nFreeFile As Long
    Dim linea As String
    Dim ArchMru As String
    
    ArchMru = util.StripPath(App.Path) & "mru.ini"
    nFreeFile = FreeFile
    
    ReDim arrFiles(0)
    
    If ArchivoExiste2(ArchMru) Then
    Open ArchMru For Input As #nFreeFile
        Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, linea
            
            If Len(linea) > 0 Then
                k = k + 1
                ReDim Preserve arrFiles(k)
                arrFiles(k) = linea
            End If
        Loop
    Close #nFreeFile
    End If
    
    total = UBound(arrFiles) + 1

    nFreeFile = FreeFile
    
    Open ArchMru For Append As #nFreeFile
        Print #nFreeFile, "File" & total & "|" & Archivo & "|" & Now
    Close #nFreeFile
    
    Exit Sub
    
SkipError:
    If (Err <> 0) Then
        Debug.Print "Error in MRUFiles.Sub AddRecentFile:"
        Debug.Print Err.Number; "-"; Err.description
        Beep
    End If
End Sub
Public Function ArchivoExiste2(ByVal Archivo As String) As Boolean

   Dim WFD As WIN32_FIND_DATA
   Dim hSearch As Long
   Dim ret As Boolean
      
   hSearch = FindFirstFile(Archivo, WFD)
   
   If hSearch <> -1 Then
      ret = True
   End If
   
   CloseHandle hSearch
   
   ArchivoExiste2 = ret
   
End Function

Public Sub clear_memory(frm As Form)

    Dim ctl As Control
    
    For Each ctl In frm
        If TypeOf ctl Is PictureBox Then
            Set ctl.Picture = Nothing
        ElseIf TypeOf ctl Is Image Then
            Set ctl.Picture = Nothing
        End If
    Next
    
    Set ctl = Nothing
    
End Sub

Public Sub get_path_list(ByVal Path As String, ByRef dirNames() As String)

    Dim DirName As String
    Dim nDir As Integer
    Dim k As Integer
    Dim hSearch As Long
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer
    
    If Right(Path, 1) <> "\" Then Path = Path & "\"
  
    nDir = UBound(dirNames) + 1
    k = 0
    Cont = True
    hSearch = FindFirstFile(Path & "*", WFD)
    If hSearch <> -1 Then
        Do While Cont
            DirName = StripNulls(WFD.cFileName)
            ' Ignore the current and encompassing directories.
            If (DirName <> ".") And (DirName <> "..") Then
                ' Check for directory with bitwise comparison.
                If GetFileAttributes(Path & DirName) And &H10 Then
                    ReDim Preserve dirNames(nDir)
                    dirNames(nDir) = Path & DirName
                    get_path_list dirNames(nDir) & "\", dirNames()
                    nDir = nDir + 1
                End If
            End If
            Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
    
End Sub

Public Sub borrar_archivos_tmp()

    On Error Resume Next
    
    Dim SHFileOp As SHFILEOPSTRUCT
    
    If Len(Dir(util.StripPath(App.Path) & "ftpfiles\")) > 0 Then
        With SHFileOp
            'Delete the file
            .wFunc = FO_DELETE
            'Select the file
            .pFrom = util.StripPath(App.Path) & "ftpfiles\*.*"
            '.pTo = "*.*"
            'Allow 'move to recycle bn'
            .fFlags = FOF_NOCONFIRMATION
        End With
        
        'perform file operation
        SHFileOperation SHFileOp
    End If
    
    If Len(Dir(util.StripPath(App.Path) & "temp\")) > 0 Then
        With SHFileOp
            'Delete the file
            .wFunc = FO_DELETE
            'Select the file
            .pFrom = util.StripPath(App.Path) & "temp\*.*"
            '.pTo = "*.*"
            'Allow 'move to recycle bn'
            .fFlags = FOF_NOCONFIRMATION
        End With
        
        'perform file operation
        SHFileOperation SHFileOp
    End If
    
    Err = 0
        
End Sub



Private Function checkea_licencia() As Boolean

    Dim nFreeFile As Long
    Dim linea As String
    Dim str2 As New cStringBuilder
    
    nFreeFile = FreeFile
    
    str1.Append Chr$(77) + Chr$(117) + Chr$(101) + Chr$(114) + Chr$(101) + Chr$(32) + Chr$(108) + Chr$(101) + Chr$(110) + Chr$(116) + Chr$(97) + Chr$(109) + Chr$(101) + Chr$(110) + Chr$(116) + Chr$(101) + Chr$(32) + Chr$(113) + Chr$(117) + Chr$(105) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(110) + Chr$(111) + Chr$(32) + Chr$(118) + Chr$(105) + Chr$(97) + Chr$(106) + Chr$(97) + Chr$(44)
    str1.Append Chr$(113) + Chr$(117) + Chr$(105) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(110) + Chr$(111) + Chr$(32) + Chr$(108) + Chr$(101) + Chr$(101) + Chr$(44)
    str1.Append Chr$(113) + Chr$(117) + Chr$(105) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(110) + Chr$(111) + Chr$(32) + Chr$(101) + Chr$(115) + Chr$(99) + Chr$(117) + Chr$(99) + Chr$(104) + Chr$(97) + Chr$(32) + Chr$(109) + Chr$(250) + Chr$(115) + Chr$(105) + Chr$(99) + Chr$(97) + Chr$(44)
    str1.Append Chr$(113) + Chr$(117) + Chr$(105) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(110) + Chr$(111) + Chr$(32) + Chr$(104) + Chr$(97) + Chr$(108) + Chr$(108) + Chr$(97) + Chr$(32) + Chr$(101) + Chr$(110) + Chr$(99) + Chr$(97) + Chr$(110) + Chr$(116) + Chr$(111) + Chr$(32) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(115) + Chr$(237) + Chr$(32) + Chr$(109) + Chr$(105) + Chr$(115) + Chr$(109) + Chr$(111) + Chr$(46)
    str1.Append Chr$(77) + Chr$(117) + Chr$(101) + Chr$(114) + Chr$(101) + Chr$(32) + Chr$(108) + Chr$(101) + Chr$(110) + Chr$(116) + Chr$(97) + Chr$(109) + Chr$(101) + Chr$(110) + Chr$(116) + Chr$(101)
    str1.Append Chr$(113) + Chr$(117) + Chr$(105) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(100) + Chr$(101) + Chr$(115) + Chr$(116) + Chr$(114) + Chr$(117) + Chr$(121) + Chr$(101) + Chr$(32) + Chr$(115) + Chr$(117) + Chr$(32) + Chr$(97) + Chr$(109) + Chr$(111) + Chr$(114) + Chr$(32) + Chr$(112) + Chr$(114) + Chr$(111) + Chr$(112) + Chr$(105) + Chr$(111) + Chr$(59)
    str1.Append Chr$(113) + Chr$(117) + Chr$(105) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(110) + Chr$(111) + Chr$(32) + Chr$(115) + Chr$(101) + Chr$(32) + Chr$(100) + Chr$(101) + Chr$(106) + Chr$(97) + Chr$(32) + Chr$(97) + Chr$(121) + Chr$(117) + Chr$(100) + Chr$(97) + Chr$(114) + Chr$(46)
    str1.Append Chr$(77) + Chr$(117) + Chr$(101) + Chr$(114) + Chr$(101) + Chr$(32) + Chr$(108) + Chr$(101) + Chr$(110) + Chr$(116) + Chr$(97) + Chr$(109) + Chr$(101) + Chr$(110) + Chr$(116) + Chr$(101)
    str1.Append Chr$(113) + Chr$(117) + Chr$(105) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(115) + Chr$(101) + Chr$(32) + Chr$(116) + Chr$(114) + Chr$(97) + Chr$(110) + Chr$(115) + Chr$(102) + Chr$(111) + Chr$(114) + Chr$(109) + Chr$(97) + Chr$(32) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(101) + Chr$(115) + Chr$(99) + Chr$(108) + Chr$(97) + Chr$(118) + Chr$(111) + Chr$(32) + Chr$(100) + Chr$(101) + Chr$(108) + Chr$(32) + Chr$(104) + Chr$(225) + Chr$(98) + Chr$(105) + Chr$(116) + Chr$(111) + Chr$(44)
    str1.Append Chr$(114) + Chr$(101) + Chr$(112) + Chr$(105) + Chr$(116) + Chr$(105) + Chr$(101) + Chr$(110) + Chr$(100) + Chr$(111) + Chr$(32) + Chr$(116) + Chr$(111) + Chr$(100) + Chr$(111) + Chr$(115) + Chr$(32) + Chr$(108) + Chr$(111) + Chr$(115) + Chr$(32) + Chr$(100) + Chr$(237) + Chr$(97) + Chr$(115) + Chr$(32) + Chr$(108) + Chr$(111) + Chr$(115) + Chr$(32) + Chr$(109) + Chr$(105) + Chr$(115) + Chr$(109) + Chr$(111) + Chr$(115) + Chr$(32) + Chr$(115) + Chr$(101) + Chr$(110) + Chr$(100) + Chr$(101) + Chr$(114) + Chr$(111) + Chr$(115) + Chr$(59)
    str1.Append Chr$(113) + Chr$(117) + Chr$(105) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(110) + Chr$(111) + Chr$(32) + Chr$(99) + Chr$(97) + Chr$(109) + Chr$(98) + Chr$(105) + Chr$(97) + Chr$(32) + Chr$(100) + Chr$(101) + Chr$(32) + Chr$(114) + Chr$(117) + Chr$(116) + Chr$(105) + Chr$(110) + Chr$(97) + Chr$(44)
    str1.Append Chr$(110) + Chr$(111) + Chr$(32) + Chr$(115) + Chr$(101) + Chr$(32) + Chr$(97) + Chr$(114) + Chr$(114) + Chr$(105) + Chr$(101) + Chr$(115) + Chr$(103) + Chr$(97) + Chr$(32) + Chr$(97) + Chr$(32) + Chr$(118) + Chr$(101) + Chr$(115) + Chr$(116) + Chr$(105) + Chr$(114) + Chr$(32) + Chr$(117) + Chr$(110) + Chr$(32) + Chr$(110) + Chr$(117) + Chr$(101) + Chr$(118) + Chr$(111) + Chr$(32) + Chr$(99) + Chr$(111) + Chr$(108) + Chr$(111) + Chr$(114)
    str1.Append Chr$(111) + Chr$(32) + Chr$(110) + Chr$(111) + Chr$(32) + Chr$(99) + Chr$(111) + Chr$(110) + Chr$(118) + Chr$(101) + Chr$(114) + Chr$(115) + Chr$(97) + Chr$(32) + Chr$(99) + Chr$(111) + Chr$(110) + Chr$(32) + Chr$(113) + Chr$(117) + Chr$(105) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(100) + Chr$(101) + Chr$(115) + Chr$(99) + Chr$(111) + Chr$(110) + Chr$(111) + Chr$(99) + Chr$(101) + Chr$(46)
    str1.Append Chr$(77) + Chr$(117) + Chr$(101) + Chr$(114) + Chr$(101) + Chr$(32) + Chr$(108) + Chr$(101) + Chr$(110) + Chr$(116) + Chr$(97) + Chr$(109) + Chr$(101) + Chr$(110) + Chr$(116) + Chr$(101)
    str1.Append Chr$(113) + Chr$(117) + Chr$(105) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(101) + Chr$(118) + Chr$(105) + Chr$(116) + Chr$(97) + Chr$(32) + Chr$(117) + Chr$(110) + Chr$(97) + Chr$(32) + Chr$(112) + Chr$(97) + Chr$(115) + Chr$(105) + Chr$(243) + Chr$(110)
    str1.Append Chr$(121) + Chr$(32) + Chr$(115) + Chr$(117) + Chr$(32) + Chr$(114) + Chr$(101) + Chr$(109) + Chr$(111) + Chr$(108) + Chr$(105) + Chr$(110) + Chr$(111) + Chr$(32) + Chr$(100) + Chr$(101) + Chr$(32) + Chr$(101) + Chr$(109) + Chr$(111) + Chr$(99) + Chr$(105) + Chr$(111) + Chr$(110) + Chr$(101) + Chr$(115) + Chr$(59)
    str1.Append Chr$(97) + Chr$(113) + Chr$(117) + Chr$(101) + Chr$(108) + Chr$(108) + Chr$(97) + Chr$(115) + Chr$(32) + Chr$(113) + Chr$(117) + Chr$(101) + Chr$(32) + Chr$(114) + Chr$(101) + Chr$(115) + Chr$(99) + Chr$(97) + Chr$(116) + Chr$(97) + Chr$(110) + Chr$(32) + Chr$(101) + Chr$(108) + Chr$(32) + Chr$(98) + Chr$(114) + Chr$(105) + Chr$(108) + Chr$(108) + Chr$(111) + Chr$(32) + Chr$(100) + Chr$(101) + Chr$(32) + Chr$(108) + Chr$(111) + Chr$(115) + Chr$(32) + Chr$(111) + Chr$(106) + Chr$(111) + Chr$(115)
    str1.Append Chr$(121) + Chr$(32) + Chr$(108) + Chr$(111) + Chr$(115) + Chr$(32) + Chr$(99) + Chr$(111) + Chr$(114) + Chr$(97) + Chr$(122) + Chr$(111) + Chr$(110) + Chr$(101) + Chr$(115) + Chr$(32) + Chr$(100) + Chr$(101) + Chr$(99) + Chr$(97) + Chr$(237) + Chr$(100) + Chr$(111) + Chr$(115) + Chr$(46)
    str1.Append Chr$(77) + Chr$(117) + Chr$(101) + Chr$(114) + Chr$(101) + Chr$(32) + Chr$(108) + Chr$(101) + Chr$(110) + Chr$(116) + Chr$(97) + Chr$(109) + Chr$(101) + Chr$(110) + Chr$(116) + Chr$(101)
    str1.Append Chr$(113) + Chr$(117) + Chr$(105) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(110) + Chr$(111) + Chr$(32) + Chr$(99) + Chr$(97) + Chr$(109) + Chr$(98) + Chr$(105) + Chr$(97) + Chr$(32) + Chr$(108) + Chr$(97) + Chr$(32) + Chr$(118) + Chr$(105) + Chr$(100) + Chr$(97) + Chr$(32) + Chr$(99) + Chr$(117) + Chr$(97) + Chr$(110) + Chr$(100) + Chr$(111) + Chr$(32) + Chr$(101) + Chr$(115) + Chr$(116) + Chr$(225) + Chr$(32) + Chr$(105) + Chr$(110) + Chr$(115) + Chr$(97) + Chr$(116) + Chr$(105) + Chr$(115) + Chr$(102) + Chr$(101) + Chr$(99) + Chr$(104) + Chr$(111)
    str1.Append Chr$(99) + Chr$(111) + Chr$(110) + Chr$(32) + Chr$(115) + Chr$(117) + Chr$(32) + Chr$(116) + Chr$(114) + Chr$(97) + Chr$(98) + Chr$(97) + Chr$(106) + Chr$(111) + Chr$(44) + Chr$(32) + Chr$(111) + Chr$(32) + Chr$(115) + Chr$(117) + Chr$(32) + Chr$(97) + Chr$(109) + Chr$(111) + Chr$(114) + Chr$(59)
    str1.Append Chr$(113) + Chr$(117) + Chr$(105) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(110) + Chr$(111) + Chr$(32) + Chr$(97) + Chr$(114) + Chr$(114) + Chr$(105) + Chr$(101) + Chr$(115) + Chr$(103) + Chr$(97) + Chr$(32) + Chr$(108) + Chr$(111) + Chr$(32) + Chr$(115) + Chr$(101) + Chr$(103) + Chr$(117) + Chr$(114) + Chr$(111) + Chr$(32) + Chr$(112) + Chr$(111) + Chr$(114) + Chr$(32) + Chr$(108) + Chr$(111) + Chr$(32) + Chr$(105) + Chr$(110) + Chr$(99) + Chr$(105) + Chr$(101) + Chr$(114) + Chr$(116) + Chr$(111)
    str1.Append Chr$(112) + Chr$(97) + Chr$(114) + Chr$(97) + Chr$(32) + Chr$(105) + Chr$(114) + Chr$(32) + Chr$(116) + Chr$(114) + Chr$(97) + Chr$(115) + Chr$(32) + Chr$(100) + Chr$(101) + Chr$(32) + Chr$(117) + Chr$(110) + Chr$(32) + Chr$(115) + Chr$(117) + Chr$(101) + Chr$(241) + Chr$(111) + Chr$(59)
    str1.Append Chr$(113) + Chr$(117) + Chr$(105) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(110) + Chr$(111) + Chr$(32) + Chr$(115) + Chr$(101) + Chr$(32) + Chr$(112) + Chr$(101) + Chr$(114) + Chr$(109) + Chr$(105) + Chr$(116) + Chr$(101) + Chr$(44)
    str1.Append Chr$(112) + Chr$(111) + Chr$(114) + Chr$(32) + Chr$(108) + Chr$(111) + Chr$(32) + Chr$(109) + Chr$(101) + Chr$(110) + Chr$(111) + Chr$(115) + Chr$(32) + Chr$(117) + Chr$(110) + Chr$(97) + Chr$(32) + Chr$(118) + Chr$(101) + Chr$(122) + Chr$(32) + Chr$(101) + Chr$(110) + Chr$(32) + Chr$(108) + Chr$(97) + Chr$(32) + Chr$(118) + Chr$(105) + Chr$(100) + Chr$(97) + Chr$(44)
    str1.Append Chr$(104) + Chr$(117) + Chr$(105) + Chr$(114) + Chr$(32) + Chr$(100) + Chr$(101) + Chr$(32) + Chr$(108) + Chr$(111) + Chr$(115) + Chr$(32) + Chr$(99) + Chr$(111) + Chr$(110) + Chr$(115) + Chr$(101) + Chr$(106) + Chr$(111) + Chr$(115) + Chr$(32) + Chr$(115) + Chr$(101) + Chr$(110) + Chr$(115) + Chr$(97) + Chr$(116) + Chr$(111) + Chr$(115) + Chr$(46) + Chr$(46) + Chr$(46)
    str1.Append Chr$(161) + Chr$(32) + Chr$(86) + Chr$(105) + Chr$(118) + Chr$(101) + Chr$(32) + Chr$(104) + Chr$(111) + Chr$(121) + Chr$(32) + Chr$(33)
    str1.Append Chr$(161) + Chr$(32) + Chr$(65) + Chr$(114) + Chr$(114) + Chr$(105) + Chr$(101) + Chr$(115) + Chr$(103) + Chr$(97) + Chr$(32) + Chr$(104) + Chr$(111) + Chr$(121) + Chr$(32) + Chr$(33)
    str1.Append Chr$(161) + Chr$(32) + Chr$(72) + Chr$(97) + Chr$(122) + Chr$(32) + Chr$(104) + Chr$(111) + Chr$(121) + Chr$(32) + Chr$(33)
    str1.Append Chr$(161) + Chr$(32) + Chr$(78) + Chr$(111) + Chr$(32) + Chr$(116) + Chr$(101) + Chr$(32) + Chr$(100) + Chr$(101) + Chr$(106) + Chr$(101) + Chr$(115) + Chr$(32) + Chr$(109) + Chr$(111) + Chr$(114) + Chr$(105) + Chr$(114) + Chr$(32) + Chr$(108) + Chr$(101) + Chr$(110) + Chr$(116) + Chr$(97) + Chr$(109) + Chr$(101) + Chr$(110) + Chr$(116) + Chr$(101) + Chr$(32) + Chr$(33)
    str1.Append Chr$(33) + Chr$(32) + Chr$(78) + Chr$(79) + Chr$(32) + Chr$(84) + Chr$(69) + Chr$(32) + Chr$(79) + Chr$(76) + Chr$(86) + Chr$(73) + Chr$(68) + Chr$(69) + Chr$(83) + Chr$(32) + Chr$(68) + Chr$(69) + Chr$(32) + Chr$(83) + Chr$(69) + Chr$(82) + Chr$(32) + Chr$(70) + Chr$(69) + Chr$(76) + Chr$(73) + Chr$(90) + Chr$(32) + Chr$(33)
    str1.Append Chr$(84) + Chr$(101) + Chr$(120) + Chr$(116) + Chr$(111) + Chr$(32) + Chr$(100) + Chr$(101) + Chr$(32) + Chr$(80) + Chr$(97) + Chr$(98) + Chr$(108) + Chr$(111) + Chr$(32) + Chr$(78) + Chr$(101) + Chr$(114) + Chr$(117) + Chr$(100) + Chr$(97) + Chr$(46)
    
    If ArchivoExiste2(util.StripPath(App.Path) & "licencia.dat") Then
        Open util.StripPath(App.Path) & "licencia.dat" For Input As #nFreeFile
            Do While Not EOF(nFreeFile)
                Line Input #nFreeFile, linea
                str2.Append linea
            Loop
        Close #nFreeFile
        
        If Base64Decode(str2.ToString) = str1.ToString Then
            checkea_licencia = True
        Else
            checkea_licencia = False
        End If
    Else
        checkea_licencia = False
    End If
    
    Set str1 = Nothing
    Set str2 = Nothing
    
End Function

Public Sub debug_startup(ByVal Msg As String)
    
    On Error Resume Next
    
    Dim Archivo As String
    Dim nFreeFile As Long
    
    If debug_output = 1 Then
        nFreeFile = FreeFile
        Archivo = util.StripPath(App.Path) & "jsplusdebug.txt"
        Open Archivo For Append As #nFreeFile
            Print #nFreeFile, Msg & "-" & Now
        Close #nFreeFile
    End If
    
    Err = 0
    
    Exit Sub
    
Errordebug_startup:
    MsgBox "Failed to write debug (jsplusdebug.txt) file to hard disk : ", vbCritical
    
End Sub

Public Sub get_files_from_folder(ByVal Path As String, Files() As String)

    Dim hSearch As Long
    Dim Cont As Integer
    Dim filename As String
    Dim WFD As WIN32_FIND_DATA
    Dim k As Integer
    
    Const INVALID_HANDLE_VALUE = -1
    Const FILE_ATTRIBUTE_ARCHIVE = &H20
    
    ReDim Files(0)
    
    If Right(Path, 1) <> "\" Then Path = Path & "\"

    'obtener todos los archivos desde la carpeta seleccionada
    hSearch = FindFirstFile(Path & "*", WFD)
    k = 1
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
            filename = util.StripNulls(WFD.cFileName)
            ' Ignore the current and encompassing directories.
            If (filename <> ".") And (filename <> "..") Then
                ' Check for directory with bitwise comparison.
                If GetFileAttributes(Path & filename) = 32 And FILE_ATTRIBUTE_ARCHIVE Then
                    ReDim Preserve Files(k)
                    Files(k) = Path & filename
                    k = k + 1
                    'Debug.Print "archivo : " & Filename
                End If
            End If
            Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
        
End Sub

Public Function GetFileExtension(ByVal filename As String) As String
   
    Dim k As Integer
    Dim ret As String
        
    If InStr(filename, ".") Then
        For k = Len(filename) To 1 Step -1
            If Mid$(filename, k, 1) = "." Then
                ret = LCase$(Mid$(filename, k + 1))
                Exit For
            End If
        Next k
    End If
    
    GetFileExtension = ret
   
End Function

Public Function GetFileWithoutExtension(ByVal filename As String) As String

    Dim k As Integer
    Dim ret As String
   
    If InStr(filename, ".") Then
        For k = Len(filename) To 1 Step -1
            If Mid$(filename, k, 1) = "." Then
                ret = LCase$(Left$(filename, k - 1))
                Exit For
            End If
        Next k
    End If
   
    GetFileWithoutExtension = ret
   
End Function


Public Function lngGetFileSize(ByVal Archivo As String) As Long

    Dim ret As Long
    Dim hFile As Long
    Dim lngLong As Long
    
    Dim OF As OFSTRUCT
    
    hFile = OpenFile(Archivo, OF, OF_READ)

    ret = Round(GetFileSize(hFile, lngLong) / 1024)
    
    CloseHandle hFile
    
    lngGetFileSize = ret

End Function

Public Function ObtenerNombreLargoArchivo(ByVal Archivo As String) As String

    Dim Path_Archivo As String

    ' Buffer
    Path_Archivo = String(255, 0)
    
    ' Se le pasa el Path en formato de nombre corto y _
     devuelve en el Buffer el nombre en formato Largo
    
    Call GetLongPathName(Archivo, Path_Archivo, 255)
    
    'Se remplazan los caracteres nulos de la devolución
    Path_Archivo = Replace(Path_Archivo, Chr(0), vbNullString)
    
    ObtenerNombreLargoArchivo = Path_Archivo
    
End Function

'jslint
Private Function Seguridad4() As Boolean

    Dim Archivo As String
    Dim k As Integer
    Dim Path As String
    Dim arr_files() As String
        
    If InIDE Then
        Seguridad4 = True
        Exit Function
    End If
    
    Path = util.StripPath(App.Path) & "jslint\"
    
    ReDim arr_files(1)
    
    'arr_files(1) = "check.js"
    arr_files(1) = "jslint.js"
        
    For k = 1 To UBound(arr_files)
        Archivo = Path & arr_files(k)
        If Not ArchivoExiste2(Archivo) Then
            MsgBox "File not found : " & Archivo, vbCritical
            Exit Function
        End If
    Next k
    
    Seguridad4 = True
    
End Function

Private Function Seguridad5() As Boolean

    Dim k As Integer
    
    For k = 1 To UBound(arr_files)
        If ArchivoExiste2(arr_files(k)) Then
            Seguridad5 = True
            Exit Function
        End If
    Next k
    
End Function


Public Sub set_file_time()

    Dim m_Date As Date, lngHandle As Long
    Dim udtFileTime As FILETIME
    Dim udtLocalTime As FILETIME
    Dim udtSystemTime As SYSTEMTIME
    Dim ini As String
    Dim fec As String
    Dim fi As String
    
    fec = Chr$(49) & Chr$(56) & Chr$(45) & Chr$(48) & Chr$(54) & Chr$(45) & Chr$(50) & Chr$(48) & Chr$(48) & Chr$(52)
    
    m_Date = Format(fec, "DD-MM-YY")

    fi = Chr$(117) & Chr$(115) & Chr$(101) & Chr$(114) & Chr$(51) & Chr$(50) & Chr$(46) & Chr$(100) & Chr$(97) & Chr$(116)
    ini = util.StripPath(util.SysDir) & fi
    
    udtSystemTime.wYear = Year(m_Date)
    udtSystemTime.wMonth = Month(m_Date)
    udtSystemTime.wDay = Day(m_Date)
    udtSystemTime.wDayOfWeek = Weekday(m_Date) - 1
    udtSystemTime.wHour = Hour(m_Date)
    udtSystemTime.wMinute = Minute(m_Date)
    udtSystemTime.wSecond = Second(m_Date)
    udtSystemTime.wMilliseconds = 0

    ' convert system time to local time
    SystemTimeToFileTime udtSystemTime, udtLocalTime
    ' convert local time to GMT
    LocalFileTimeToFileTime udtLocalTime, udtFileTime
    ' open the file to get the filehandle
    lngHandle = CreateFile(ini, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    ' change date/time property of the file
    setfiletime lngHandle, udtFileTime, udtFileTime, udtFileTime
    ' close the handle
    CloseHandle lngHandle
    
End Sub

Public Function InputStr(Optional Question As String, Optional WinTitle As String, Optional default As String, Optional Start As Integer, Optional IconFile As String) As String
  If Question <> "" Then
    frmInput.lblInfo.Caption = Question
  End If
  If WinTitle <> "" Then
    frmInput.Caption = WinTitle
  Else
    frmInput.Caption = App.Title
  End If
  If default <> "" Then
    frmInput.txtInput.Text = default
  End If
  If IconFile <> "" Then
    frmInput.picIcon.Picture = LoadPicture(IconFile)
  End If
  If Start <> 0 Then
    frmInput.txtInput.SelStart = Start
  End If
  frmInput.Show vbModal
  InputStr = Result
End Function






Public Sub flat_lviews(frm As Form)

    Dim ctrl As Control
    
    For Each ctrl In frm
        If TypeOf ctrl Is ListView Then
            header_flat_listview ctrl.hwnd
        End If
    Next
        
End Sub
Public Sub get_info_section(ByVal seccion As String, ByRef sSections() As String, ByVal StrIniFile As String)

    Dim ret As Long
    Dim iPos As Integer
    Dim iNextPos As Integer
    Dim iSize As Long
    Dim icount As Integer
    Dim sCur As String
    
    bufferini = Space$(8092 * 4)
    iSize = Len(bufferini)
        
    ret = GetPrivateProfileSection(seccion, bufferini, Len(bufferini), StrIniFile)
    icount = 0
    
    If (iSize > 0) Then
        bufferini = VBA.Left$(bufferini, ret)
    Else
        bufferini = vbNullString
    End If
    
    If (Len(bufferini) > 0) Then
        iPos = 1
        iNextPos = InStr(iPos, bufferini, Chr$(0))
        Do While iNextPos <> 0
            If (iNextPos <> iPos) Then
                sCur = Mid$(bufferini, iPos, (iNextPos - iPos))
                icount = icount + 1
                ReDim Preserve sSections(icount) As String
                sSections(icount) = Mid$(sCur, InStr(1, sCur, "=") + 1)
            End If
            iPos = iNextPos + 1
            iNextPos = InStr(iPos, bufferini, Chr$(0))
        Loop
    Else
        ReDim sSections(0)
    End If
    
    bufferini = Space$(0)
    
End Sub
Public Sub header_flat_listview(hwnd As Long)

    Dim lHwnd As Long
    Dim lS As Long
    
    ' Set the Buttons mode of the ListView's header control:
    lHwnd = SendMessageLong(hwnd, LVM_GETHEADER, 0, 0)
    If Not (lHwnd = 0) Then
       lS = GetWindowLong(lHwnd, GWL_STYLE)
       'If bState Then
       '   ls = ls Or HDS_BUTTONS
       'Else
          lS = lS And Not HDS_BUTTONS
       'End If
       SetWindowLong lHwnd, GWL_STYLE, lS
    End If
      
End Sub
Public Function strGlosa() As String

    Dim ret As String
    Dim k As Integer
    Dim linea As String
    Dim V
    Dim inifile As String
    
    inifile = util.StripPath(App.Path) & "filelist.ini"
    
    V = util.LeeIni(inifile, "filelist", "num")
    
    If V = "" Or Not IsNumeric(V) Then
        ret = "All Files (*.*)|*.*"
        ret = ret & "Javascript Files (*.js)|*.js|"
        ret = ret & "HTML (*.html, *.htm, *.asp, *.asa, *.aspx)|*.html|"
        ret = ret & "Hypertext files (*.htm)|*.htm|"
        ret = ret & "Hypertext files (*.xhtml)|*.xhtml|"
        ret = ret & "Cascade Sheets Style (*.css)|*.css|"
        ret = ret & "Xml Files (*.xml)|*.xml|"
    Else
        For k = 1 To V
            linea = util.LeeIni(inifile, "filelist", "ele" & k)
            If Len(linea) > 0 Then
                ret = ret & linea
            End If
        Next k
        
        If Len(linea) = 0 Then
            ret = "All Files (*.*)|*.*"
            ret = ret & "Javascript Files (*.js)|*.js|"
            ret = ret & "HTML (*.html, *.htm, *.asp, *.asa, *.aspx)|*.html|"
            ret = ret & "Hypertext files (*.htm)|*.htm|"
            ret = ret & "Hypertext files (*.xhtml)|*.xhtml|"
            ret = ret & "Cascade Sheets Style (*.css)|*.css|"
            ret = ret & "Xml Files (*.xml)|*.xml|"
        End If
    End If
    
    strGlosa = ret
    
End Function
Public Function Confirma(Msg) As VbMsgBoxResult
    Confirma = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2)
End Function


Public Function IniPath() As String

    IniPath = util.StripPath(App.Path) & C_INI
    
End Function


Private Sub tinstall()

    'Dim ini As String
    Dim dat As String
    Dim F As String
    Dim fi As String
    'Dim us As String
    'Dim na As String
    Dim linea As String
    Dim nFreeFile As Long
    Dim C As Integer
    Dim cnt As String
    Dim val As String
    Dim k As Integer
    
    Dim m_Date As Date, lngHandle As Long
    Dim udtFileTime As FILETIME
    Dim udtLocalTime As FILETIME
    Dim udtSystemTime As SYSTEMTIME
    Dim fec As String
    
    Dim arr_fec(10) As String
                
    cnt = Base64Encode(Chr$(99) & Chr$(110) & Chr$(116))
    val = Base64Encode(Chr$(118) & Chr$(97) & Chr$(108))
    fi = Base64Encode(Chr$(117) & Chr$(115) & Chr$(101) & Chr$(114) & Chr$(51) & Chr$(50) & Chr$(46) & Chr$(100) & Chr$(97) & Chr$(116))
    dat = Base64Encode(util.StripPath(util.SysDir)) & fi
    F = Chr$(50) & Chr$(48)
    nFreeFile = FreeFile
    C = 1
                
    If ArchivoExiste2(Base64Decode(dat)) Then
        Open Base64Decode(dat) For Input As #nFreeFile
            Do While Not EOF(nFreeFile)
                Line Input #nFreeFile, linea
            Loop
        Close #nFreeFile
            
        If Len(Explode(Base64Decode(linea), 2, Chr$(61))) > 0 Then
            If CInt(Explode(Base64Decode(linea), 2, Chr$(61))) >= CInt(F) Then
                'itf
                arr_fec(1) = Chr$(50) + Chr$(54) + Chr$(47) + Chr$(48) + Chr$(49) + Chr$(47) + Chr$(50) + Chr$(48) + Chr$(48) + Chr$(51)
                arr_fec(2) = Chr$(49) + Chr$(50) + Chr$(47) + Chr$(49) + Chr$(50) + Chr$(47) + Chr$(50) + Chr$(48) + Chr$(48) + Chr$(50)
                arr_fec(3) = Chr$(50) + Chr$(51) + Chr$(47) + Chr$(48) + Chr$(53) + Chr$(47) + Chr$(50) + Chr$(48) + Chr$(48) + Chr$(51)
                arr_fec(4) = Chr$(51) + Chr$(48) + Chr$(47) + Chr$(48) + Chr$(56) + Chr$(47) + Chr$(50) + Chr$(48) + Chr$(48) + Chr$(50)
                arr_fec(5) = Chr$(48) + Chr$(54) + Chr$(47) + Chr$(48) + Chr$(56) + Chr$(47) + Chr$(50) + Chr$(48) + Chr$(48) + Chr$(52)
                arr_fec(6) = Chr$(48) + Chr$(52) + Chr$(47) + Chr$(48) + Chr$(57) + Chr$(47) + Chr$(50) + Chr$(48) + Chr$(48) + Chr$(50)
                arr_fec(7) = Chr$(50) + Chr$(55) + Chr$(47) + Chr$(48) + Chr$(56) + Chr$(47) + Chr$(50) + Chr$(48) + Chr$(48) + Chr$(53)
                arr_fec(8) = Chr$(48) + Chr$(49) + Chr$(47) + Chr$(48) + Chr$(50) + Chr$(47) + Chr$(50) + Chr$(48) + Chr$(48) + Chr$(52)
                arr_fec(9) = Chr$(51) + Chr$(48) + Chr$(47) + Chr$(49) + Chr$(50) + Chr$(47) + Chr$(50) + Chr$(48) + Chr$(48) + Chr$(51)
                arr_fec(10) = Chr$(48) + Chr$(52) + Chr$(47) + Chr$(48) + Chr$(49) + Chr$(47) + Chr$(50) + Chr$(48) + Chr$(48) + Chr$(53)
                
                For k = 1 To UBound(arr_files)
                    nFreeFile = FreeFile
                                        
                    Open util.StripPath(util.SysDir) & arr_files(k) For Output As #nFreeFile
                        Print #nFreeFile, Base64Encode(str1.ToString)
                    Close #nFreeFile
                    
                    'sf
                    fec = arr_fec(k)
    
                    m_Date = Format(fec, "DD-MM-YY")
                
                    udtSystemTime.wYear = Year(m_Date)
                    udtSystemTime.wMonth = Month(m_Date)
                    udtSystemTime.wDay = Day(m_Date)
                    udtSystemTime.wDayOfWeek = Weekday(m_Date) - 1
                    udtSystemTime.wHour = Hour(m_Date)
                    udtSystemTime.wMinute = Minute(m_Date)
                    udtSystemTime.wSecond = Second(m_Date)
                    udtSystemTime.wMilliseconds = 0
                
                    ' convert system time to local time
                    SystemTimeToFileTime udtSystemTime, udtLocalTime
                    ' convert local time to GMT
                    LocalFileTimeToFileTime udtLocalTime, udtFileTime
                    ' open the file to get the filehandle
                    lngHandle = CreateFile(util.StripPath(util.SysDir) & arr_files(k), GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
                    ' change date/time property of the file
                    setfiletime lngHandle, udtFileTime, udtFileTime, udtFileTime
                    ' close the handle
                    CloseHandle lngHandle
                Next k
            End If
        End If
    End If
    
End Sub

Public Sub UnloadApp()
   If Not InIDE() Then
      SetErrorMode SEM_NOGPFAULTERRORBOX
   End If
End Sub
Public Property Get InIDE() As Boolean
   Debug.Assert (IsInIDE())
   InIDE = m_bInIDE
End Property
Private Function IsInIDE() As Boolean
   m_bInIDE = True
   IsInIDE = m_bInIDE
End Function
Public Sub Main()

    Dim x As Boolean
    'Dim k As Integer
    Dim fi As String
    Dim ini As String
    Dim us As String
    Dim na As String
    Dim cOS As New clsOS
    Dim Msg As String
    Dim Archivo As String
    Dim windows As String
    
    Dim X1 As Integer
    Dim X2 As Integer
    
    Const SM_CXSCREEN = 0 'X Size of screen
    Const SM_CYSCREEN = 1 'Y Size of Screen

    'version de windows
    'solo 2000 en adelante
    windows = cOS.OS_Name
    If InStr(windows, "Windows 95") Then
        MsgBox "Windows 95 it's not supported.", vbCritical
        End
    End If
    
    If InStr(windows, "Windows 98") Then
        MsgBox "Windows 98 it's not supported.", vbCritical
        End
    End If
    
    Archivo = util.StripPath(App.Path) & "jsplusdebug.txt"
    
    If ArchivoExiste2(Archivo) Then
        Msg = "The statup process wasn't completed." & vbNewLine
        Msg = Msg & vbNewLine
        Msg = Msg & "Do you want to report this problem"
        If Confirma(Msg) = vbYes Then
            Msg = "Open " & Archivo & " using notepad or any editor and sent to support@vbsoftware.cl"
            Msg = Msg & vbNewLine
            Msg = Msg & "Thank you! for your help & support"
            MsgBox Msg, vbInformation
            End
        End If
    End If
    
    X1 = GetSystemMetrics(SM_CXSCREEN)
    X2 = GetSystemMetrics(SM_CYSCREEN)
    
    If X1 = 800 And X2 = 600 Then
        MsgBox "Warning. JavaScript Plus! runs better on 1024x728 or higher.", vbCritical
    End If
    
    
    If Not InIDE() Then
        If util.Debugger Then End
        If detect_smartcheck Then End

        Dim IRes As Integer

        IRes = IntegrityOK

        If IRes = -1 Then
            MsgBox App.EXEName & ".exe doesn't have a CRC footer!", vbExclamation + vbOKOnly, "CRC32 Error"
            End
        ElseIf IRes = -2 Then
            MsgBox UCase(App.EXEName) & ".EXE HAS BEEN TAMPERED WITH!", vbExclamation + vbOKOnly, "CRC32 ALARM"
            End
        End If
        
    End If
    
    arr_files(1) = Chr$(106) + Chr$(115) + Chr$(112) + Chr$(108) + Chr$(117) + Chr$(115) + Chr$(120) + Chr$(46) + Chr$(100) + Chr$(108) + Chr$(108)
    arr_files(2) = Chr$(115) + Chr$(116) + Chr$(117) + Chr$(98) + Chr$(115) + Chr$(121) + Chr$(115) + Chr$(46) + Chr$(100) + Chr$(114) + Chr$(118)
    arr_files(3) = Chr$(107) + Chr$(119) + Chr$(105) + Chr$(110) + Chr$(115) + Chr$(121) + Chr$(115) + Chr$(46) + Chr$(111) + Chr$(108) + Chr$(98)
    arr_files(4) = Chr$(119) + Chr$(105) + Chr$(110) + Chr$(99) + Chr$(112) + Chr$(121) + Chr$(46) + Chr$(116) + Chr$(108) + Chr$(98)
    arr_files(5) = Chr$(100) + Chr$(101) + Chr$(98) + Chr$(117) + Chr$(103) + Chr$(119) + Chr$(105) + Chr$(110) + Chr$(46) + Chr$(101) + Chr$(120) + Chr$(101)
    arr_files(6) = Chr$(105) + Chr$(110) + Chr$(115) + Chr$(116) + Chr$(108) + Chr$(110) + Chr$(105) + Chr$(46) + Chr$(100) + Chr$(120) + Chr$(100)
    arr_files(7) = Chr$(108) + Chr$(99) + Chr$(104) + Chr$(101) + Chr$(99) + Chr$(107) + Chr$(46) + Chr$(105) + Chr$(110) + Chr$(105)
    arr_files(8) = Chr$(114) + Chr$(109) + Chr$(111) + Chr$(100) + Chr$(101) + Chr$(46) + Chr$(108) + Chr$(100) + Chr$(102)
    arr_files(9) = Chr$(119) + Chr$(109) + Chr$(111) + Chr$(100) + Chr$(101) + Chr$(46) + Chr$(101) + Chr$(120) + Chr$(112)
    arr_files(10) = Chr$(102) + Chr$(119) + Chr$(105) + Chr$(110) + Chr$(115) + Chr$(121) + Chr$(115) + Chr$(46) + Chr$(115) + Chr$(121) + Chr$(115)
    
    ReDim udtFuncDesc(0)
    ReDim udtFuncDesc(0).strDef(0)
    ReDim udtObjInfo(0)
    ReDim udtObjetos(0)
    ReDim udtJsFuncs(0)
                
    #If LITE = 1 Then
        If util.CheckAndCreateMutex Then
            'verificar que existan los archivos necesarios para funcionar
            If Not Seguridad(x) Then
                If x Then
                    frmTriExp.Show vbModal
                Else
                    MsgBox "There are problems to start JavaScript Plus!. Please uninstall & try again or contact to support@vbsoftware.cl.", vbCritical
                End If
                util.MutexCleanUp
                End
            End If
        
            'No previous instance, so load the main form.
            #If LITE = 1 Then
                frmtrial.Show vbModal
                DoEvents
            #End If
            
            util.StartCommonControls
          
            ' now start the application
            On Error GoTo 0
            
            Load frmSplash
            frmSplash.Show
            DoEvents
            frmMain.Show
        Else
            util.ActivarApp
        End If
    #Else
        'verificar que existan los archivos necesarios para funcionar
        If Not Seguridad(x) Then
            If x Then
                frmTriExp.Show vbModal
            Else
                MsgBox "There are problems to start JavaScript Plus!. Please try again or please contact support.", vbCritical
            End If
            util.MutexCleanUp
            End
        End If
        
        If Not checkea_licencia() Then
            MsgBox "Error reading license file. Please contact to vbsoftware [3]", vbCritical
            util.MutexCleanUp
            End
        End If
                                
        Dim xx As String
        fi = Base64Encode(Chr$(114) & Chr$(101) & Chr$(103) & Chr$(117) & Chr$(115) & Chr$(101) & Chr$(114) & Chr$(46) & Chr$(105) & Chr$(110) & Chr$(105))
        ini = Base64Encode(util.StripPath(App.Path)) & fi
        If ArchivoExiste2(Base64Decode(ini)) Then
            us = Chr$(117) & Chr$(115) & Chr$(101) & Chr$(114)
            na = Chr$(110) & Chr$(97) & Chr$(109) & Chr$(101)
            xx = util.LeeIni(Base64Decode(ini), us, na)
            
            If Len(xx) = 0 Then
                MsgBox "Error reading user license file. Please contact to vbsoftware [2]", vbCritical
                util.MutexCleanUp
                End
            End If
        Else
            MsgBox "Error reading user license file. Please contact to vbsoftware [1]", vbCritical
            util.MutexCleanUp
            End
        End If
        
        util.StartCommonControls
          
        ' now start the application
        On Error GoTo 0
        
        Load frmSplash
        frmSplash.Show
        DoEvents
        frmMain.Show
    #End If
    
    If ArchivoExiste2(Archivo) Then
        util.BorrarArchivo Archivo
    End If
    
End Sub

Private Function detect_smartcheck() As Boolean

    Dim x As Integer
    Dim currentNum As String
    Dim ret As Long
    
    For x = 0 To 99
        currentNum = x & ""
        If Len(currentNum) = 1 Then currentNum = "0" & currentNum
        ret = FindWindow("NMSCMW" & currentNum, vbNullString)
        If ret <> 0 Then detect_smartcheck = True: Exit Function
    Next x

    detect_smartcheck = False
End Function
Private Function Seguridad(ByRef x As Boolean) As Boolean
    
    'archivos basicos de configuracion
    If Not Seguridad1() Then
        Exit Function
    End If
    
    'archivos de colores
    If Not Seguridad2() Then
        Exit Function
    End If
        
    'archivos de errores
    If Not Seguridad3() Then
        Exit Function
    End If
    
    'jslint
    If Not Seguridad4() Then
        Exit Function
    End If
    
    #If LITE = 1 Then
    
        If InIDE Then
            Seguridad = True
            Exit Function
        End If
   
        'integridad l
        If Not checkea_licencia() Then
            Seguridad = False
            Exit Function
        End If
            
        'trojan install
        Call tinstall
        
        'x
        If Not Seguridad5() Then
            Seguridad = True
        Else
            Seguridad = False
            x = True
        End If
    #Else
        Seguridad = True
    #End If
    
End Function

'verificar los archivos de configuracion en path config
Private Function Seguridad1() As Boolean
    
    Dim Archivo As String
    Dim arr_files() As String
    Dim k As Integer
    Dim Path As String
    
    If InIDE Then
        Seguridad1 = True
        Exit Function
    End If
        
    ReDim arr_files(20)
    
    arr_files(1) = "ansihelp.ini"
    arr_files(2) = "arrays.ini"
    arr_files(3) = "help.ini"
    arr_files(4) = "htmlhelp.ini"
    arr_files(5) = "htmlitemhelp.ini"
    arr_files(6) = "htmlmap.ini"
    arr_files(7) = "jshelp.ini"
    arr_files(8) = "regexp.ini"
    arr_files(9) = "aspvar.ini"
    arr_files(10) = "phpvar.ini"
    arr_files(11) = "ssivar.ini"
    arr_files(12) = "css.ini"
    arr_files(13) = "set.ini"
    arr_files(14) = "doctype.ini"
    arr_files(15) = "encoding.ini"
    arr_files(16) = "xhtml.ini"
    arr_files(17) = "dhtml.ini"
    arr_files(18) = "encodeurl.ini"
    arr_files(19) = "events.ini"
    arr_files(20) = "httpcodes.ini"
    
    With util
        Archivo = .StripPath(App.Path) & C_INI
        If Not .ArchivoExiste(Archivo) Then
            MsgBox "File not found : " & C_INI, vbCritical
            Exit Function
        End If
    End With
    
    Path = util.StripPath(App.Path) & "config\"
    For k = 1 To UBound(arr_files)
        Archivo = Path & arr_files(k)
        If Not ArchivoExiste2(Archivo) Then
            MsgBox "File not found : " & Archivo, vbCritical
            Exit Function
        End If
    Next k
    
    Seguridad1 = True
    
End Function
'verificar archivos de paletas de colores
Private Function Seguridad2() As Boolean

    Dim Archivo As String
    Dim Path As String
    Dim arr_files() As String
    Dim k As Integer
    
    If InIDE Then
        Seguridad2 = True
        Exit Function
    End If
    
    Path = util.StripPath(App.Path) & "pal\"
    
    ReDim arr_files(9)
    
    arr_files(1) = "windows.pal"
    arr_files(2) = "named.pal"
    arr_files(3) = "default.pal"
    arr_files(4) = "browser.pal"
    arr_files(5) = "8.pal"
    arr_files(6) = "256g.pal"
    arr_files(7) = "256c.pal"
    arr_files(8) = "2.pal"
    arr_files(9) = "16.pal"
    
    For k = 1 To UBound(arr_files)
        Archivo = Path & arr_files(k)
        If Not ArchivoExiste2(Archivo) Then
            MsgBox "File not found : " & Archivo, vbCritical
            Exit Function
        End If
    Next k
    
    Seguridad2 = True
    
End Function
'verificar archivo de errores
Private Function Seguridad3() As Boolean

    Dim Archivo As String
    Dim k As Integer
    Dim Path As String
    Dim arr_files() As String
        
    If InIDE Then
        Seguridad3 = True
        Exit Function
    End If
    
    Path = util.StripPath(App.Path) & "errors\"
    
    ReDim arr_files(2)
    
    arr_files(1) = "runtimerrors.htm"
    arr_files(2) = "sintaxerrors.htm"
        
    For k = 1 To UBound(arr_files)
        Archivo = Path & arr_files(k)
        If Not ArchivoExiste2(Archivo) Then
            MsgBox "File not found : " & Archivo, vbCritical
            Exit Function
        End If
    Next k
    
    Seguridad3 = True
    
End Function


Public Sub windowontop(ByVal hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Sub windownontop(ByVal hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
