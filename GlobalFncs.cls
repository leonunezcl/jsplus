VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUtilFncs"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrPASS As String

Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_BEGIN = 0
Private Const FILE_CURRENT = 1
Private Const FILE_END = 2
Private Const MAX_PATH = 260

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * 33
End Type

Private Type FILETIME
    dwLowDateTime           As Long
    dwHighDateTime          As Long
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

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Private m_Orientacion As Long
Private Const largo_cabecera = 8
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

'Constants used to make changes to the values contained in the DevMode
Private Const GW_HWNDPREV = 3
Private Const DM_MODIFY = 8
Private Const DM_IN_BUFFER = DM_MODIFY
Private Const DM_COPY = 2
Private Const DM_OUT_BUFFER = DM_COPY
Private Const DM_DUPLEX = &H1000&
Private Const DMDUP_SIMPLEX = 1
Private Const DMDUP_VERTICAL = 2
Private Const DMDUP_HORIZONTAL = 3
Private Const DM_ORIENTATION = &H1&
Private PageDirection As Integer

Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long

'Constants for NT security
Private Const cSTANDARD_RIGHTS_REQUIRED = &HF0000
Private Const cPRINTER_ACCESS_ADMINISTER = &H4
Private Const cPRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (cSTANDARD_RIGHTS_REQUIRED Or cPRINTER_ACCESS_ADMINISTER Or cPRINTER_ACCESS_USE)

Private Type PRINTER_DEFAULTS
'Note:
'  The definition of Printer_Defaults in the VB5 API viewer is incorrect.
'  Below, pDevMode has been corrected to LONG.
    pDatatype As String
    pDevMode As Long
    DesiredAccess As Long
End Type

Private Const NULLPTR = 0&
' Constants for DEVMODE
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
' Constants for DocumentProperties
'Private Const DM_MODIFY = 8
'Private Const DM_COPY = 2
'Private Const DM_IN_BUFFER = DM_MODIFY
'Private Const DM_OUT_BUFFER = DM_COPY
' Constants for dmOrientation
Private Const DMORIENT_PORTRAIT = 1
Private Const DMORIENT_LANDSCAPE = 2
' Constants for dmPrintQuality
Private Const DMRES_DRAFT = (-1)
Private Const DMRES_HIGH = (-4)
Private Const DMRES_LOW = (-2)
Private Const DMRES_MEDIUM = (-3)
' Constants for dmTTOption
Private Const DMTT_BITMAP = 1
Private Const DMTT_DOWNLOAD = 2
Private Const DMTT_DOWNLOAD_OUTLINE = 4
Private Const DMTT_SUBDEV = 3
' Constants for dmColor
Private Const DMCOLOR_COLOR = 2
Private Const DMCOLOR_MONOCHROME = 1

Private Type DEVMODE
    dmDeviceName(1 To CCHDEVICENAME) As Byte
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName(1 To CCHFORMNAME) As Byte
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
    "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long

'Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
'        "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
'        ByVal pDefault As Long) As Long

Private Declare Function DocumentProperties Lib "winspool.drv" _
        Alias "DocumentPropertiesA" (ByVal hwnd As Long, _
        ByVal hPrinter As Long, ByVal pDeviceName As String, _
        pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) _
        As Long

Private Declare Function ClosePrinter Lib "winspool.drv" _
        (ByVal hPrinter As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Private Const SHGFI_LARGEICON = &H0       'Large icon
Private Const SHGFI_SMALLICON = &H1       'Small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "COMCTL32.DLL" _
    (ByVal hIml&, ByVal I&, ByVal hdcDest&, _
    ByVal X&, ByVal Y&, ByVal Flags&) As Long


'----------------------------------------------------------
'Private variables
'----------------------------------------------------------
Private ShInfo As SHFILEINFO

Private Const SEE_MASK_INVOKEDLIST = &HC

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    '  Optional fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Enum eSpecialFolders
    eCSIDL_DESKTOP = &H0
    eCSIDL_PROGRAMS = &H2
    eCSIDL_CONTROLS = &H3
    eCSIDL_PRINTERS = &H4
    eCSIDL_PERSONAL = &H5
    eCSIDL_FAVORITES = &H6
    eCSIDL_STARTUP = &H7
    eCSIDL_RECENT = &H8
    eCSIDL_SENDTO = &H9
    eCSIDL_BITBUCKET = &HA
    eCSIDL_STARTMENU = &HB
    eCSIDL_DESKTOPDIRECTORY = &H10
    eCSIDL_DRIVES = &H11
    eCSIDL_NETWORK = &H12
    eCSIDL_NETHOOD = &H13
    eCSIDL_FONTS = &H14
    eCSIDL_TEMPLATES = &H15
End Enum

'Private Const MAX_PATH = 260
Private Type SHItemID
    cb As Long
    abID As Byte
End Type
Private Type ItemIDList
    mkid As SHItemID
End Type
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ItemIDList) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function ShellExecuteEx Lib "shell32" (lpSEI As SHELLEXECUTEINFO) As Long

Private Declare Function SHRunDialog Lib "shell32" Alias "#61" _
                            (ByVal hOwner As Long, _
                             ByVal Unknown1 As Long, _
                             ByVal Unknown2 As Long, _
                             ByVal szTitle As String, _
                             ByVal szPrompt As String, _
                             ByVal uFlags As Long) As Long
                             
'api calls to retereive the system and windows folders
Private Declare Function GetSystemDirectory _
        Lib "kernel32" _
        Alias "GetSystemDirectoryA" _
            (ByVal lpBuffer As String, _
             ByVal nSize As Long) _
             As Long
Private Declare Function GetWindowsDirectory _
        Lib "kernel32" _
        Alias "GetWindowsDirectoryA" _
            (ByVal lpBuffer As String, _
             ByVal nSize As Long) _
             As Long

'get the location of the temp directory on the system
Private Declare Function GetTempDirectory _
        Lib "kernel32" _
        Alias "GetTempPathA" _
            (ByVal lBufferLength As Long, _
             ByVal strBuffer As String) _
             As Long

'get information about the current operating system
Private Declare Function GetVersionEx _
        Lib "kernel32" _
        Alias "GetVersionExA" _
            (ByRef lpVersionInformation As OSVERSIONINFO) _
             As Long

'registry api calls

'close an open registry key
Private Declare Function RegCloseKey _
        Lib "advapi32.dll" _
            (ByVal hKey As Long) _
             As Long
             
'connect with the registry on a remote machine
Private Declare Function RegConnectRegistry _
        Lib "advapi32.dll" _
        Alias "RegConnectRegistryA" _
            (ByVal lpMachineName As String, _
             ByVal hKey As Long, _
             phkResult As Long) _
             As Long

'writes all the attributes of the specified open key
'into the registry
Private Declare Function RegFlushKey _
        Lib "advapi32.dll" _
            (ByVal hKey As Long) _
             As Long

'get the security attributes of the specified key
Private Declare Function RegGetKeySecurity _
        Lib "advapi32.dll" _
            (ByVal hKey As Long, _
             ByVal SecurityInformation As Long, _
             pSecurityDescriptor As SECURITY_DESCRIPTOR, _
             lpcbSecurityDescriptor As Long) _
             As Long

'creates a subkey under HKEY_USER or HKEY_LOCAL_MACHINE
'and stores registration information from a specified
'file into that subkey. This registration information
'is in the form of a hive. A hive is a discrete body of
'keys, subkeys, and values that is rooted at the top of
'the registry hierarchy. A hive is backed by a single
'file and .LOG file
Private Declare Function RegLoadKey _
        Lib "advapi32.dll" _
        Alias "RegLoadKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal lpFile As String) _
             As Long

'notify a specified procedure (use the AddressOf
'operator), that a key has changed
Private Declare Function RegNotifyChangeKeyValue _
        Lib "advapi32.dll" _
            (ByVal hKey As Long, _
             ByVal bWatchSubtree As Long, _
             ByVal dwNotifyFilter As Long, _
             ByVal hEvent As Long, _
             ByVal fAsynchronus As Long) _
             As Long

'open a registry key for access
Private Declare Function RegOpenKey _
        Lib "advapi32.dll" _
        Alias "RegOpenKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             phkResult As Long) _
             As Long
Private Declare Function RegOpenKeyEx _
        Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal ulOptions As Long, _
             ByVal samDesired As Long, _
             phkResult As Long) _
             As Long

'get key information
Private Declare Function RegQueryInfoKey _
        Lib "advapi32.dll" _
        Alias "RegQueryInfoKeyA" _
            (ByVal hKey As Long, _
             ByVal lpClass As String, _
             lpcbClass As Long, _
             ByVal lpReserved As Long, _
             lpcSubKeys As Long, _
             lpcbMaxSubKeyLen As Long, _
             lpcbMaxClassLen As Long, _
             lpcValues As Long, _
             lpcbMaxValueNameLen As Long, _
             lpcbMaxValueLen As Long, _
             lpcbSecurityDescriptor As Long, _
             lpftLastWriteTime As FILETIME) _
             As Long

'get value information. Note that if you declare the
'lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValue _
        Lib "advapi32.dll" _
        Alias "RegQueryValueA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal lpValue As String, _
             lpcbValue As Long) _
             As Long
Private Declare Function RegQueryValueEx _
        Lib "advapi32.dll" _
        Alias "RegQueryValueExA" _
            (ByVal hKey As Long, _
             ByVal lpValueName As String, _
             ByVal lpReserved As Long, _
             lpType As Long, _
             lpData As Any, _
             lpcbData As Long) _
             As Long

'replace one key with another
Private Declare Function RegReplaceKey _
        Lib "advapi32.dll" _
        Alias "RegReplaceKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal lpNewFile As String, _
             ByVal lpOldFile As String) _
             As Long

'reads registry information from a file and enters it
'into the registry
Private Declare Function RegRestoreKey _
        Lib "advapi32.dll" _
        Alias "RegRestoreKeyA" _
            (ByVal hKey As Long, _
             ByVal lpFile As String, _
             ByVal dwFlags As Long) _
             As Long

'saves a registry key and all its values to a file
Private Declare Function RegSaveKey _
        Lib "advapi32.dll" _
        Alias "RegSaveKeyA" _
            (ByVal hKey As Long, _
             ByVal lpFile As String, _
             lpSecurityAttributes As SECURITY_ATTRIBUTES) _
             As Long

'set the security attributes of the specified registry
'key
Private Declare Function RegSetKeySecurity _
        Lib "advapi32.dll" _
            (ByVal hKey As Long, _
             ByVal SecurityInformation As Long, _
             pSecurityDescriptor As SECURITY_DESCRIPTOR) _
             As Long

'set the information of an existing value. Note that if
'you declare the lpData parameter as String, you must
'pass it By Value.
Private Declare Function RegSetValue _
        Lib "advapi32.dll" _
        Alias "RegSetValueA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal dwType As Long, _
             ByVal lpData As String, _
             ByVal cbData As Long) _
             As Long
Private Declare Function RegSetValueEx _
        Lib "advapi32.dll" _
        Alias "RegSetValueExA" _
            (ByVal hKey As Long, _
             ByVal lpValueName As String, _
             ByVal Reserved As Long, _
             ByVal dwType As Long, _
             lpData As Any, _
             ByVal cbData As Long) _
             As Long
             
'unloads a registry key and its values from the registry
Private Declare Function RegUnLoadKey _
        Lib "advapi32.dll" _
        Alias "RegUnLoadKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String) _
             As Long

'system information api calls
Private Declare Sub GlobalMemoryStatus _
        Lib "kernel32" _
            (lpBuffer As MEMORYSTATUS)
Private Declare Function GetDiskFreeSpace _
        Lib "kernel32" _
        Alias "GetDiskFreeSpaceA" _
            (ByVal lpRootPathName As String, _
             lpSectorsPerCluster As Long, _
             lpBytesPerSector As Long, _
             lpNumberOfFreeClusters As Long, _
             lpTotalNumberOfClusters As Long) _
             As Long
Private Declare Function GetTickCount _
        Lib "kernel32" _
            () As Long

'------------------------------------------------
'                   ENUMERATORS
'------------------------------------------------
Public Enum MemType
    CPUUsage
    MemoryUsage
    TotalPhysical
    AvailablePhysical
    TotalPageFile
    AvailablePageFile
    TotalVirtual
    AvailableVirtual
    TotalDisk
    AvailableDisk
End Enum

Public Enum AccessType
    FileInput = 0
    FileOutPut = 1
    FileRandom = 2
    FileBinary = 3
    FileAppend = 4
End Enum

'registry root directory constants
'Public Enum RegistryHives
'    HKEY_CLASSES_ROOT = &H80000000
'    HKEY_CURRENT_CONFIG = &H80000005
'    HKEY_CURRENT_USER = &H80000001
'    HKEY_DYN_DATA = &H80000006
'    HKEY_LOCAL_MACHINE = &H80000002
'    HKEY_PERFORMANCE_DATA = &H80000004
'    HKEY_USERS = &H80000003
'End Enum

'registry key constants
Public Enum RegistryKeyAccess
    KEY_CREATE_LINK = &H20
    KEY_CREATE_SUB_KEY = &H4
    KEY_ENUMERATE_SUB_KEYS = &H8
    KEY_EVENT = &H1    '  Event contains key event record
    KEY_NOTIFY = &H10
    KEY_QUERY_VALUE = &H1
    KEY_SET_VALUE = &H2
    READ_CONTROL = &H20000
    STANDARD_RIGHTS_ALL = &H1F0000
    STANDARD_RIGHTS_REQUIRED = &HF0000
    Synchronize = &H100000
    STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
    STANDARD_RIGHTS_READ = (READ_CONTROL)
    STANDARD_RIGHTS_WRITE = (READ_CONTROL)
    KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not Synchronize))
    KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not Synchronize))
    KEY_EXECUTE = ((KEY_READ) And (Not Synchronize))
    KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not Synchronize))
End Enum

'registry value attributes
Public Enum RegistryKeyValues
    REG_CREATED_NEW_KEY = &H1               ' New Registry Key created
    REG_EXPAND_SZ = 2            ' Unicode nul terminated string
    REG_FULL_RESOURCE_DESCRIPTOR = 9  ' Resource list in the hardware description
    REG_LINK = 6                ' Symbolic Link (unicode)
    REG_MULTI_SZ = 7             ' Multiple Unicode strings
    REG_NONE = 0                ' No value type
    REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
    REG_NOTIFY_CHANGE_LAST_SET = &H4               ' Time stamp
    REG_NOTIFY_CHANGE_NAME = &H1               ' Create or delete (child)
    REG_NOTIFY_CHANGE_SECURITY = &H8
    REG_OPENED_EXISTING_KEY = &H2               ' Existing Key opened
    REG_OPTION_BACKUP_RESTORE = 4    ' open for backup or restore
    REG_OPTION_CREATE_LINK = 2      ' Created key is a symbolic link
    REG_OPTION_NON_VOLATILE = 0     ' Key is preserved when system is rebooted
    REG_OPTION_RESERVED = 0        ' Parameter is reserved
    REG_OPTION_VOLATILE = 1        ' Key is not preserved when system is rebooted
    REG_REFRESH_HIVE = &H2               ' Unwind changes to last flush
    REG_RESOURCE_LIST = 8          ' Resource list in the resource map
    REG_RESOURCE_REQUIREMENTS_LIST = 10
    'REG_SZ = 1                 ' Unicode nul terminated string
    REG_WHOLE_HIVE_VOLATILE = &H1               ' Restore whole hive volatile
    REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
    REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)
End Enum

Public Enum RegistryLongTypes
    REG_BINARY = 3              ' Free form binary
    REG_DWORD = 4               ' 32-bit number
    REG_DWORD_BIG_ENDIAN = 5    ' 32-bit number
    REG_DWORD_LITTLE_ENDIAN = 4 ' 32-bit number (same as REG_DWORD)
End Enum

'error codes returned
Public Enum RegistryErrorCodes
    ERROR_ACCESS_DENIED = 5&
    ERROR_INVALID_PARAMETER = 87 '  dderror
    ERROR_MORE_DATA = 234 '  dderror
    ERROR_SUCCESS = 0&
End Enum

'the shell folders like my documents, recycle bin, temp directory etc.
Public Enum ShellFoldersType
    'registry entry names
    ApplicationDataDir = 0
    TempInetFilesDir = 1
    CookiesDir = 2
    DesktopDir = 3
    FavouritesDir = 4
    FontsDir = 5
    HistoryDir = 6
    LocalAppDataDir = 7
    NetHoodDir = 8
    MyDocumentsDir = 9
    PrintHoodDir = 10
    StartProgramsDir = 11
    RecentDir = 12
    SendToDir = 13
    StartMenuDir = 14
    StartupDir = 15
    TemplatesDir = 16
    
    'these next items are not stored in the registry
    SystemDir = 17
    WindowsDir = 18
    TempDir = 19 'temperory folder is always in the Windows directory
End Enum

Public Enum StartLoginType
    RunBeforeLogin
    RunAfterLogin
End Enum

'the different nt privilages that can be set/unset
Public Enum EnumNTSettings
    'items that can be disabled on the Lock Screen
    CHANGE_PASSWORD = 0
    LOCK_WORKSTATION = 1
    REGISTRY_TOOLS = 2
    TASK_MGR = 3
    
    'the tabs on the Display Properties dialog box
    DISP_APPEARANCE_PAGE = 4
    DISP_BACKGROUND_PAGE = 5
    DISP_CPL = 6
    DISP_SCREENSAVER = 7
    DISP_SETTINGS = 8
End Enum

'------------------------------------------------
'               USER-DEFINED TYPES
'------------------------------------------------
'holds information about the current operating system that the program is
'running on
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

'the current status of physical (ram), virtual memory and the page file.
Private Type MEMORYSTATUS
        dwLength                As Long
        dwMemoryLoad            As Long
        dwTotalPhys             As Long
        dwAvailPhys             As Long
        dwTotalPageFile         As Long
        dwAvailPageFile         As Long
        dwTotalVirtual          As Long
        dwAvailVirtual          As Long
End Type

'defined structures needed
Private Type ACL
        AclRevision             As Byte
        Sbz1                    As Byte
        AclSize                 As Integer
        AceCount                As Integer
        Sbz2                    As Integer
End Type


Private Type SECURITY_ATTRIBUTES
        nLength                 As Long
        lpSecurityDescriptor    As Long
        bInheritHandle          As Long
End Type

Private Type SECURITY_DESCRIPTOR
        revision                As Byte
        Sbz1                    As Byte
        Control                 As Long
        gstrOwner               As Long
        Group                   As Long
        Sacl                    As ACL
        Dacl                    As ACL
End Type

'------------------------------------------------
'             MODULE-LEVEL CONSTANTS
'------------------------------------------------

'module constants
Private Const WIN_INFO_SUBKEY       As String = "Software\Microsoft\Windows\CurrentVersion" 'HKEY_LOCAL_MACHINE
Private Const WIN_NT_INFO_SUBKEY    As String = "Software\Microsoft\Windows NT\CurrentVersion"                              'HKEY_LOCAL_MACHINE
Private Const SHELL_FOLDERS_SUBKEY  As String = ".Default\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" 'HKEY_USERS
Private Const COUNTRY_SUBKEY        As String = ".Default\Control Panel\International" 'HKEY_USERS
Private Const NT_SETTINGS           As String = WIN_INFO_SUBKEY & "\Policies\System"                                          'HKEY_CURRENT_USER
Private Const W2K_SETTINGS          As String = WIN_INFO_SUBKEY & "\Group Policy Objects\LocalUser\Software\Microsoft\Windows\CurrentVersion\Policies\System"  'HKEY_CURRENT_USER
Private Const STARTUP_AL_SUBKEY     As String = WIN_INFO_SUBKEY & "\Run" 'run after login screen
Private Const STARTUP_BL_SUBKEY     As String = WIN_INFO_SUBKEY & "\RunServices" 'run before login screen

Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Private Type BROWSEINFO
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Const BIF_RETURNONLYFSDIRS = 1
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long

Private HandleSnapshot As Long
Private InfoProcesso As PROCESSENTRY32
Private Processi As Long

Private Const SPI_GETWORKAREA = 48

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const OFS_MAXPATHNAME = 256
Private Const OF_EXIST = &H4000

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Const IDC_WAIT = 32514&
Private Const IDC_ARROW = 32512&

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, lpCursorName As Any) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Private Const MAX_MODULE_NAME32 = 256
Private Const TH32CS_SNAPPROCESS = &H2   ' Process info required
Private Const TH32CS_SNAPTHREAD = &H4    ' Thread info required
Private Const TH32CS_SNAPMODULE = &H8    ' Module info required

' Describe a process
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

' Describe a thread
Private Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type

' Describe a module
Private Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr  As Long
    modBaseSize  As Long
    hModule  As Long
    szModule As String * MAX_MODULE_NAME32
    szExePath As String * MAX_PATH
End Type

' Functions
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal dwProcessId As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hObject As Long, p As PROCESSENTRY32) As Boolean
Private Declare Function Process32Next Lib "kernel32" (ByVal hObject As Long, p As PROCESSENTRY32) As Boolean
Private Declare Function Thread32First Lib "kernel32" (ByVal hObject As Long, p As THREADENTRY32) As Boolean
Private Declare Function Thread32Next Lib "kernel32" (ByVal hObject As Long, p As THREADENTRY32) As Boolean
Private Declare Function Module32First Lib "kernel32" (ByVal hObject As Long, p As MODULEENTRY32) As Boolean
Private Declare Function Module32Next Lib "kernel32" (ByVal hObject As Long, p As MODULEENTRY32) As Boolean

' This function is not from ToolHelp but you need it to destroy a snapshot
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const IMG_PROC = 1
Private Const IMG_APP = 2
Private Const IMG_INFO = 3
Private Const IMG_THREAD = 4
Private Const IMG_MODULE = 5

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Declare Function GetSystemMetrics Lib "user32" (ByVal Index As Long) As Long
Private X As Long, Y As Long

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const OF_READ = &H0&
Private Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const LANG_NEUTRAL = &H0

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

Dim lpFSHigh As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const ES_NUMBER = &H2000&

Private Const UniqueString          As String = "SoloUnaInstanciaDeJspadPaulinitaTeExtranoMucho"
'This variable will have the handle of the mutex
Private GMutex      As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function IsDebuggerPresent Lib "kernel32" () As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Const WAIT_OBJECT_0         As Long = &H0
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function FlatSB_SetScrollPos Lib "COMCTL32" (ByVal hwnd As Long, ByVal Code As Long, ByVal nPos As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollRange Lib "COMCTL32" (ByVal hwnd As Long, ByVal Code As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "COMCTL32" (ByVal hwnd As Long, ByVal Code As Long, ByVal fShow As Boolean) As Boolean
Private Declare Function FlatSB_GetScrollPos Lib "COMCTL32" (ByVal hwnd As Long, ByVal Code As Long) As Long

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
                            (ByVal dwFlags As Long, lpSource As Any, _
                            ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
                            ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const LANG_USER_DEFAULT = &H400&

Public Function JoinTokens(ByVal buffer As String) As String

    Dim k As Integer
    Dim c As Integer
    Dim j As Integer
    
    Dim ret As String
    Dim Buf As String
    
    'contar las ocurrencias
    For k = 1 To Len(buffer)
        If Mid$(buffer, k, 1) = "|" Then
            c = c + 1
        End If
    Next k
    
    j = 1
    For k = 1 To c
        Buf = Explode(buffer, k, "|")
        
        If Len(Trim$(Buf)) > 0 Then
            If j = 1 Then
                ret = Buf
                j = 2
            Else
                ret = ret & " " & Buf
            End If
        End If
    Next k
    
    JoinTokens = ret
    
End Function


Public Function MakeRegion(picSkin As PictureBox) As Long
    
    ' Make a windows "region" based on a given picture box'
    ' picture. This done by passing on the picture line-
    ' by-line and for each sequence of non-transparent
    ' pixels a region is created that is added to the
    ' complete region. I tried to optimize it so it's
    ' fairly fast, but some more optimizations can
    ' always be done - mainly storing the transparency
    ' data in advance, since what takes the most time is
    ' the GetPixel calls, not Create/CombineRgn
    
    Dim X As Long, Y As Long, StartLineX As Long
    Dim FullRegion As Long, LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean  ' Flags whether we are in a non-tranparent pixel sequence
    Dim hDC As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    hDC = picSkin.hDC
    PicWidth = picSkin.ScaleWidth
    PicHeight = picSkin.ScaleHeight
    
    InFirstRegion = True: InLine = False
    X = Y = StartLineX = 0
    
    Const RGN_OR = 2
    
    ' The transparent color is always the color of the
    ' top-left pixel in the picture. If you wish to
    ' bypass this constraint, you can set the tansparent
    ' color to be a fixed color (such as pink), or
    ' user-configurable
    TransparentColor = GetPixel(hDC, 0, 0)
    
    For Y = 0 To PicHeight - 1
        For X = 0 To PicWidth - 1
            
            If GetPixel(hDC, X, Y) = TransparentColor Or X = PicWidth Then
                ' We reached a transparent pixel
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, Y, X, Y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        ' Always clean up your mess
                        DeleteObject LineRegion
                    End If
                End If
            Else
                ' We reached a non-transparent pixel
                If Not InLine Then
                    InLine = True
                    StartLineX = X
                End If
            End If
        Next
    Next
    
    MakeRegion = FullRegion
End Function

Public Sub FindFiles(path As String, SearchStr As String, ByRef List1 As ListBox, ByVal quebuscar As Integer)
    'KPD-Team 1999
    'E-Mail: KPDTeam@Allapi.net
    'URL: http://www.allapi.net/

    Dim Filename As String ' Walking filename variable...
    Dim DirName As String ' SubDirectory Name
    Dim dirNames() As String ' Buffer for directory name entries
    Dim nDir As Integer ' Number of directories in this path
    Dim I As Integer ' For-loop counter...
    Dim hSearch As Long ' Search Handle
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer
    Dim k As Integer
    Dim c As Integer
    Dim misscripts()
    Dim v
    Dim ext As String
    Dim buffer As String
    
    If quebuscar = 1 Then
        ReDim misscripts(5)
        misscripts(1) = "bmp"
        misscripts(2) = "gif"
        misscripts(3) = "jpg"
        misscripts(4) = "jpeg"
        misscripts(5) = "ico"
        
        For k = 1 To 5
            buffer = buffer & "." & misscripts(k)
        Next k
        
    ElseIf quebuscar = 2 Then
        v = Util.LeeIni(IniPath, "media", "num")
        If v = "" Or Not IsNumeric(v) Then
            Exit Sub
        End If
        ReDim misscripts(v)
        For k = 1 To v
            ext = Util.Explode(Util.LeeIni(IniPath, "media", "ele" & k), 2, "|")
            If Len(ext) > 0 Then
                misscripts(c) = Mid$(ext, 3)
                buffer = buffer & Mid$(ext, 2)
                c = c + 1
            End If
        Next k
    ElseIf quebuscar = 3 Then
        v = Util.LeeIni(IniPath, "filelist", "num")
        If v = "" Or Not IsNumeric(v) Then
            Exit Sub
        End If
        ReDim misscripts(v)
        For k = 1 To v
            ext = Util.Explode(Util.LeeIni(IniPath, "filelist", "ele" & k), 2, "|")
            If Len(ext) > 0 Then
                misscripts(c) = Mid$(ext, 3)
                buffer = buffer & Mid$(ext, 2)
                c = c + 1
            End If
        Next k
    End If
    
    If VBA.Right(path, 1) <> "\" Then path = path & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    Cont = True
    hSearch = FindFirstFile(path & "*", WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
        DirName = StripNulls(WFD.cFileName)
        ' Ignore the current and encompassing directories.
        If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                dirNames(nDir) = DirName
                'DirCount = DirCount + 1
                nDir = nDir + 1
                ReDim Preserve dirNames(nDir)
            End If
        End If
        Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
    ' Walk through this directory and sum file sizes.
    hSearch = FindFirstFile(path & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            Filename = StripNulls(WFD.cFileName)
            If (Filename <> ".") And (Filename <> "..") Then
                For c = 1 To UBound(misscripts)
                    If InStr(Filename, "." & misscripts(c)) Then
                        If LCase$(VBA.Right$(Filename, 3)) = LCase(misscripts(c)) Then
                            List1.AddItem path & Filename
                            If Cont Mod 100 = 0 Then DoEvents
                            Exit For
                        End If
                    End If
                Next c
            End If
            Cont = FindNextFile(hSearch, WFD) ' Get next file
        Wend
        Cont = FindClose(hSearch)
    End If
    ' If there are sub-directories...
    If nDir > 0 Then
        ' Recursively walk into them...
        For I = 0 To nDir - 1
            Call FindFiles(path & dirNames(I) & "\", SearchStr, List1, quebuscar)
        Next I
    End If
End Sub

Public Sub FlatScrollbar(hwnd As Long)

    
End Sub

Public Sub FontStuff(ByVal titulo As String, picDraw As PictureBox)
    
    On Error GoTo GetOut
    Dim f As LOGFONT, hPrevFont As Long, hFont As Long, FontName As String
    Dim FontSize As Integer
    FontSize = 10 'Val(txtSize.Text)
    
    f.lfEscapement = 10 * 90 'Val(txtDegree.Text) 'rotation angle, in tenths
    FontName = "Tahoma" + Chr$(0) 'null terminated
    f.lfFaceName = FontName
    f.lfHeight = (FontSize * -20) / Screen.TwipsPerPixelY
    hFont = CreateFontIndirect(f)
    hPrevFont = SelectObject(picDraw.hDC, hFont)
    
    picDraw.CurrentX = 0
    'picDraw.CurrentY = 310
    
    picDraw.CurrentY = picDraw.Height - 10
    picDraw.Print titulo
    
    '  Clean up, restore original font
    hFont = SelectObject(picDraw.hDC, hPrevFont)
    DeleteObject hFont
    
    Exit Sub
GetOut:
    Exit Sub

End Sub

Public Function ChequeaCrc() As Boolean

'    On Error Resume Next
'
'    Dim OurCRC As String
'    Dim FileCRC As String * largo_cabecera
'    Dim FPos As Long
'    Dim i As Long
'
'    MyCrc.AppPath = Util.StripPath(App.path)
'    MyCrc.EXEName = App.EXEName
'    OurCRC = MyCrc.CheckCRC
'    FPos = FileLen(MyCrc.AppExe)
'    Open MyCrc.AppExe For Binary As #1
'        Get #1, FPos - 7, FileCRC   'read the last 8 bytes of the file
'    Close #1
'
'    'Descramble the hash
'    For i = 1 To largo_cabecera
'        Mid(FileCRC, i, 1) = Chr$(Asc(Mid(FileCRC, i, 1)) Xor 30)
'    Next i
'
'    'MsgBox "FileCRC : " & FileCRC
'    'MsgBox "OurCRC : " & OurCRC
'
'    'verificar si la tabla ha sido inicializada
'    Dim flag As Boolean
'    flag = True
'    For i = 1 To largo_cabecera
'        If Asc(Mid$(FileCRC, i, 1)) <> 30 Then
'            flag = False
'            Exit For
'        End If
'    Next i
'
'    'Make sure the CRC hash is present
'    If flag Then
'       MsgBox "Warning : This file has not protected. Remember this.", vbInformation
'       ChequeaCrc = True
'       Exit Function
'    End If
'
'    'Make sure the FileCRC is the same as the actual CRC
'    If Trim$(FileCRC) = Trim$(OurCRC) Then
'        ChequeaCrc = True
'    Else
'        MsgBox "This file is damaged!", vbCritical
'        ChequeaCrc = False
'    End If
'
'    Exit Function
    
End Function

Public Function IsCompiled() As Boolean

  On Error GoTo NotCompiled

   Debug.Print 1 / 0
   IsCompiled = True

NotCompiled:

End Function
Public Sub ActivarApp()
        
    Dim PrevHndl As Long
    Dim Result As Long
    
    PrevHndl = FindWindow("ThunderRT6Main", App.Title)
    
    'Get handle to previous window.
    PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
    'Restore the program.
    Result = OpenIcon(PrevHndl)
    'Activate the application.
    Result = SetForegroundWindow(PrevHndl)
    'End the application.
    End
    
End Sub

Public Function CheckAndCreateMutex() As Boolean
   
    GMutex = CreateMutex(0&, 0&, UniqueString)         'First, lets create the mutex
    
    If GMutex = 0 Then                                'Error occurred for some reason.
        MsgBox "The mechanism to ensure only one instance of this app has failed for unknown reasons.", vbCritical, "Error"
        CheckAndCreateMutex = True
    Else
    
        'Now this requires some explanation. The mutex has been created, but does not
        'belong to this specific application thread. This could be done by setting the
        'second parameter of CreateMutex function to 1, but I don't know why it doesn't
        'work in VB, though the same thing works perfectly in C++. So, another work-around
        'is here. The next function will only check if the mutex is signaled or not as the
        'second parameter is given zero. The mutex will be non-signaled if a thread owns
        'it already. Now calling once this function makes the calling thread the owner of
        'the mutex if it doesn't have an owner already.
        
        If WaitForSingleObject(GMutex, 0&) = WAIT_OBJECT_0 Then     'The mutex is signaled and
                                                                    'no other thread owns it.
                                                                    'But from now on, this thread
                                                                    'will own the mutex.
            CheckAndCreateMutex = True
        
        Else                                    'Several other things might have happened.
                                                'The mutex is may be non-signal, or it already
                                                'has an owner, or the time-out for the
                                                'function return is finished.
        
            Call CloseHandle(GMutex)            'The already owned mutex has been opened
                                                'by this thread. Well, now close the handle
                                                'of it.
            CheckAndCreateMutex = False
            
            MsgBox "Only one instance the Javascript Plus! is valid.", vbCritical
        End If
        
    End If
    
End Function


Public Function Debugger() As Boolean

    If IsDebuggerPresent <> 0 Then
        Debugger = True
    Else
        Debugger = False
    End If
    
End Function

Public Sub MutexCleanUp()

    Call ReleaseMutex(GMutex)
    
End Sub

Public Sub SetNumber(ByVal hwnd As Long)
    Dim curstyle As Long, newstyle As Long

    'retrieve the window style
    curstyle = GetWindowLong(hwnd, GWL_STYLE)

    curstyle = curstyle Or ES_NUMBER
    
    'Set the new style
    newstyle = SetWindowLong(hwnd, GWL_STYLE, curstyle)
    
End Sub

Public Sub StartCommonControls()

    ' we need to call InitCommonControls before we
   ' can use XP visual styles.  Here I'm using
   ' InitCommonControlsEx, which is the extended
   ' version provided in v4.72 upwards (you need
   ' v6.00 or higher to get XP styles)
   On Error Resume Next
   ' this will fail if Comctl not available
   '  - unlikely now though!
   Dim iccex As tagInitCommonControlsEx
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   
End Sub

Public Function StripFile(ByVal Archivo As String) As String

    Dim k As Integer
    
    For k = Len(Archivo) To 1 Step -1
        If Mid$(Archivo, k, 1) = "\" Then
            StripFile = Mid$(Archivo, k + 1)
            Exit For
        End If
    Next k
    
End Function
Public Function ApiError(ByVal elerror As Long) As String

    Dim buffer As String
    'Create a string buffer
    buffer = Space(200)
    'Format the message string
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, GetLastError, LANG_NEUTRAL, buffer, 200, ByVal 0&
    'Show the message
    ApiError = buffer

End Function


Public Function SetPathAbreviado(ByVal path As String) As String

    Dim k As Integer
    Dim c As Integer
    Dim inicio As String
    Dim final As String
    Dim ultpos As Integer
    Dim ret As String
    
    Const limite = 30
    
    ret = path
    If Len(path) > limite Then
        'sacar el ultimo posision desde slash
        For k = Len(path) To 1 Step -1
            If Mid$(path, k, 1) = "\" Then
                final = Mid$(path, k)
                Exit For
            End If
        Next k
        c = 1
        For k = 1 To Len(path)
            If Mid$(path, k, 1) = "\" Then
                If c <= limite Then
                    inicio = VBA.Left$(path, ultpos)
                    ultpos = k
                Else
                    Exit For
                End If
            End If
        Next k
        
        ret = inicio & ".." & final
    End If
    
    SetPathAbreviado = ret
    
End Function


Public Function GetInfoF(ByVal FilePath As String) As String
    
    Dim Pointer As Long, sizeofthefile As Long
    
    Pointer = lOpen(FilePath, OF_READ)
    
    'size of the file
    sizeofthefile = GetFileSize(Pointer, lpFSHigh)
    
    lclose Pointer
    
    If sizeofthefile > 1023 Then
        GetInfoF = Format(Round(sizeofthefile / 1024, 0), "###,###,###,###") & " Kb"
    Else
        GetInfoF = sizeofthefile & " Bytes"
    End If
    
End Function

Function ByteToString(ByteArray() As Byte) As String
    Dim TempStr As String
    Dim I As Integer

    For I = 1 To CCHDEVICENAME
        TempStr = TempStr & Chr(ByteArray(I))
    Next I
    ByteToString = StripNulls(TempStr)
End Function

Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = VBA.Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = Trim(OriginalStr)
End Function
Public Function GetPrinterSettings(szPrinterName As String, hDC As Long) As Boolean
    Dim hPrinter As Long
    Dim nSize As Long
    Dim pDevMode As DEVMODE
    Dim aDevMode() As Byte
    Dim TempStr As String
    Dim X As PRINTER_DEFAULTS
    
    If OpenPrinter(szPrinterName, hPrinter, X) Then
        nSize = DocumentProperties(NULLPTR, hPrinter, szPrinterName, _
                NULLPTR, NULLPTR, 0)
        ReDim aDevMode(1 To nSize)
        nSize = DocumentProperties(NULLPTR, hPrinter, szPrinterName, _
                aDevMode(1), NULLPTR, DM_OUT_BUFFER)
        Call CopyMemory(pDevMode, aDevMode(1), Len(pDevMode))

        m_Orientacion = pDevMode.dmOrientation
        
        If pDevMode.dmOrientation = DMORIENT_PORTRAIT Then
            TempStr = "PORTRAIT"
        ElseIf pDevMode.dmOrientation = DMORIENT_LANDSCAPE Then
            TempStr = "LANDSCAPE"
        Else
            TempStr = "UNDEFINED"
        End If

        Call ClosePrinter(hPrinter)
        GetPrinterSettings = True
    Else
        GetPrinterSettings = False
    End If
    
End Function
Private Sub SetOrientation(NewSetting As Long, chng As Integer, ByVal frm As Form)
    Dim PrinterHandle As Long
    Dim PrinterName As String
    Dim pd As PRINTER_DEFAULTS
    Dim MyDevMode As DEVMODE
    Dim Result As Long
    Dim Needed As Long
    Dim pFullDevMode As Long
    Dim pi2_buffer() As Long     'This is a block of memory for the Printer_Info_2 structure
        'If you need to use the Printer_Info_2 User Defined Type, the
        '  definition of Printer_Info_2 in the API viewer is incorrect.
        '  pDevMode and pSecurityDescriptor should be defined As Long.
    
    PrinterName = Printer.DeviceName
    If PrinterName = "" Then
        Exit Sub
    End If
    
    pd.pDatatype = vbNullString
    pd.pDevMode = 0&
    'Printer_Access_All is required for NT security
    pd.DesiredAccess = PRINTER_ALL_ACCESS
    
    Result = OpenPrinter(PrinterName, PrinterHandle, pd)
    
    'The first call to GetPrinter gets the size, in bytes, of the buffer needed.
    'This value is divided by 4 since each element of pi2_buffer is a long.
    Result = GetPrinter(PrinterHandle, 2, ByVal 0&, 0, Needed)
    ReDim pi2_buffer((Needed \ 4))
    Result = GetPrinter(PrinterHandle, 2, pi2_buffer(0), Needed, Needed)
    
    'The seventh element of pi2_buffer is a Pointer to a block of memory
    '  which contains the full DevMode (including the PRIVATE portion).
    pFullDevMode = pi2_buffer(7)
    
    'Copy the Public portion of FullDevMode into our DevMode structure
    Call CopyMemory(MyDevMode, ByVal pFullDevMode, Len(MyDevMode))
    
    'Make desired changes
    MyDevMode.dmDuplex = NewSetting
    MyDevMode.dmFields = DM_DUPLEX Or DM_ORIENTATION
    MyDevMode.dmOrientation = chng
    
    'Copy our DevMode structure back into FullDevMode
    Call CopyMemory(ByVal pFullDevMode, MyDevMode, Len(MyDevMode))
    
    'Copy our changes to "the PUBLIC portion of the DevMode" into "the PRIVATE portion of the DevMode"
    Result = DocumentProperties(frm.hwnd, PrinterHandle, PrinterName, ByVal pFullDevMode, ByVal pFullDevMode, DM_IN_BUFFER Or DM_OUT_BUFFER)
    
    'Update the printer's default properties (to verify, go to the Printer folder
    '  and check the properties for the printer)
    Result = SetPrinter(PrinterHandle, 2, pi2_buffer(0), 0&)
    
    Call ClosePrinter(PrinterHandle)
    
    'Note: Once "Set Printer = " is executed, anywhere in the code, after that point
    '      changes made with SetPrinter will ONLY affect the system-wide printer  --
    '      -- the changes will NOT affect the VB printer object.
    '      Therefore, it may be necessary to reset the printer object's parameters to
    '      those chosen in the devmode.
    Dim p As Printer
    For Each p In Printers
        If p.DeviceName = PrinterName Then
            Set Printer = p
            Exit For
        End If
    Next p
    Printer.Duplex = MyDevMode.dmDuplex
End Sub
Public Sub ChngPrinterOrientationLandscape(ByVal frm As Form)
    PageDirection = 2
    Call SetOrientation(DMDUP_SIMPLEX, PageDirection, frm)
End Sub
Public Sub ChngPrinterOrientationPortrait(ByVal frm As Form)

    PageDirection = 1
    Call SetOrientation(DMDUP_SIMPLEX, PageDirection, frm)
End Sub
Public Sub ResetPrinterOrientation(ByVal frm As Form)
 
    If PageDirection = 1 Then
        PageDirection = 2
    Else
        PageDirection = 1
    End If
    Call SetOrientation(DMDUP_SIMPLEX, PageDirection, frm)
End Sub
Public Function ArchivoSoloLectura(ByVal Archivo As String) As Boolean

    Const FILE_ATTRIBUTE_READONLY = &H1
    
    If (GetFileAttributes(Archivo) And FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY Then
        ArchivoSoloLectura = True
    End If

End Function


Public Function GetSpecialfolder(CSIDL As eSpecialFolders) As String
    Dim R As Long
    Dim IDL As ItemIDList
    Dim path$
    'Get the special folder
    R = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If R = 0 Then
        'Create a buffer
        path$ = Space$(512)
        'Get the path from the IDList
        R = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal path$)
        'Remove the unnecessary chr$(0)'s
        GetSpecialfolder = VBA.Left$(path, InStr(path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function

Public Sub RunDialog(ByVal hwnd As Long)
    SHRunDialog hwnd, 0, 0, vbNullString, vbNullString, 0
End Sub


Public Sub ShellExFunc(sVerb As String, sFile As String, lWin As Long)
Dim lVal As Long
'lval = ShellExecute(0, "explore", sFile, vbNullString, vbNullString, 1)
    Dim sei As SHELLEXECUTEINFO
    sei.hwnd = lWin
   ' ShellExecute lWin, sVerb, sFile, vbNullString, vbNullString, 0
    sei.lpVerb = sVerb
    If sVerb = "find" Then
        sei.lpDirectory = sFile & Chr$(0)
    Else
  '      sei.lpDirectory = sFile '& Chr$(0)
        sei.lpFile = sFile & Chr$(0)
    End If
    sei.fMask = SEE_MASK_INVOKEDLIST
    sei.cbSize = Len(sei)
    sei.nShow = 1
    ShellExecuteEx sei
End Sub

Public Sub ShellFunc(sFile As String, Optional vShow As VbAppWinStyle = 1, Optional sVerb As String = vbNullString)
    ShellExecute 0, sVerb, sFile, vbNullString, vbNullString, vShow
End Sub

Private Function AddFile(ByVal strPath As String, _
                         ByVal strFileName As String) _
                         As String
    
    'This function takes a file name and a path and will
    'put the two together to form a filepath. This is useful
    'for when the applications' path happens to be the root
    'directory.
    
    'check the last character for a backslash
    If VBA.Left(strPath, 1) = "\" Then
        'don't insert a backslash
        AddFile = strPath & strFileName
    Else
        'insert a backslash
        AddFile = strPath & "\" & strFileName
    End If
End Function

Private Function FileExists(ByVal strFilePath As String, _
                            Optional ByVal enmFlags As VbFileAttribute = vbNormal) _
                            As Boolean
    'returns True if the file exists
    
    If ((strFilePath = "") Or _
        (Dir(strFilePath, enmFlags) = "")) Then
        'invalid path/filename
        FileExists = False
    Else
        FileExists = True
    End If
End Function

Private Function HasFileAttrib(ByVal strFilePath As String, _
                               Optional ByVal enmFlags As VbFileAttribute) _
                               As Boolean
    'returns True if the file specified has the
    'appropiate type signiture, eg, a directory or is
    'read-only. If testing multiple attributes, then
    'the file MUST have all attributes to return True
    
    Dim lngErrNum As Long   'holds any error that occurred trying to access the file
    
    'make sure the file exists without upsetting any
    'stored values when the Dir function is being used
    'externally by another procedure/function
    On Error Resume Next
        'test file access
        GetAttr strFilePath
        lngErrNum = Err
    On Error GoTo 0
    
    'exit if an error occured ("#53 - File Not Found"
    'usually occurs)
    If lngErrNum > 0 Then
        HasFileAttrib = False
        Exit Function
    End If
    
    'test the file for attributes
    If ((GetAttr(strFilePath) And enmFlags) = enmFlags) Then
        HasFileAttrib = True
    Else
        HasFileAttrib = False
    End If
End Function

Private Function IsWinNT() As Boolean
    'Detect if the program is running under an NT based system (NT, 2000, XP)
    
    Const VER_PLATFORM_WIN32_NT     As Long = 2
    
    Dim osiInfo    As OSVERSIONINFO    'holds the operating system information
    Dim lngResult  As Long             'returned error value from the api call
    
    'get version information
    osiInfo.dwOSVersionInfoSize = Len(osiInfo)
    lngResult = GetVersionEx(osiInfo)
    
    'return True if the test of windows NT is positive
    IsWinNT = (osiInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Public Function ArchivoTemporal() As String

    Dim ret As String
    Dim lRet As Long
    
    ret = Space$(255)
    
    lRet = GetTempFileName(Util.StripPath(App.path), "jsplus", 0, ret)
    
    If lRet > 0 Then
        ret = VBA.Left$(ret, InStr(ret, Chr$(0)) - 1)
    End If
    
    ArchivoTemporal = ret
    
End Function

Public Function SacarBasura(ByVal Linea As String) As String

    'analizar la linea que se lee
    Linea = Replace(Linea, Chr$(8), " ")
    Linea = Replace(Linea, Chr$(9), " ")
    Linea = Replace(Linea, Chr$(10), " ")
    Linea = Replace(Linea, Chr$(11), " ")
    Linea = Replace(Linea, Chr$(13), " ")
    Linea = Replace(Linea, Chr$(0), " ")
        
    SacarBasura = Linea
    
End Function
Public Sub associate(ext As String, FileType As String, Filename As String, ByVal Descripcion As String, ByVal fMsg As Boolean)

    On Error Resume Next

    Dim b As Object
    Dim S As String
    
    Set b = CreateObject("wscript.shell")

    S = b.RegRead("HKCR\" & FileType & "\shell\open\command\")

    If ext = ".js" Then
        If InStr(S, App.EXEName) = 0 Then
            Dim Msg As String
        
            If fMsg Then
                Msg = App.Title & " it's not your javascript editor." & vbNewLine & vbNewLine
                Msg = Msg & "Set up " & App.Title & " as your default Javascript editor."
                
                If Confirma(Msg) = vbYes Then
                    b.regwrite "HKCR\" & ext & "\", FileType
                    b.regwrite "HKCR\" & FileType & "\", Descripcion
                    b.regwrite "HKCR\" & FileType & "\DefaultIcon\", Filename
                    b.regwrite "HKCR\" & FileType & "\shell\open\command\", Filename & " %L"
                    b.regdelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & ext & "\Application"
                    b.regwrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & ext & "\Application", Filename
                    b.regdelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & ext & "\OpenWithList\"
                    b.regwrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & ext & "\OpenWithList\a", Filename
                End If
            End If
        End If
    Else
        b.regwrite "HKCR\" & ext & "\", FileType
        b.regwrite "HKCR\" & FileType & "\", Descripcion '"javascript"
        b.regwrite "HKCR\" & FileType & "\DefaultIcon\", Filename
        b.regwrite "HKCR\" & FileType & "\shell\open\command\", Filename & " %L"
        b.regdelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & ext & "\Application"
        b.regwrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & ext & "\Application", Filename
        b.regdelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & ext & "\OpenWithList\"
        b.regwrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & ext & "\OpenWithList\a", Filename
    End If
    
    Set b = Nothing
    
End Sub
Public Sub BorrarArchivo(ByVal Archivo As String)
    
    Dim ret As Long
    
    ret = DeleteFile(GetShortPath(Archivo))
    
    'If ret = 0 Then
    '    MsgBox ApiError(Err.LastDllError), vbCritical
    'End If
End Sub
Public Function BrowseFolder(hwnd As Long) As String

    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BROWSEINFO

    With udtBI
        'Set the owner window
        .hwndOwner = hwnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat("C:\", "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = VBA.Left$(sPath, iNull - 1)
        End If
    End If
    
    BrowseFolder = sPath
    
End Function

Public Sub CopiarArchivo(ByVal Origen As String, ByVal Destino As String)
    CopyFile Origen, Destino, 1
End Sub

Public Function CrearDirectorio(ByVal Nombre As String) As Boolean

    Dim Security As SECURITY_ATTRIBUTES
    Dim ret&
    
    'Create a directory
    ret& = CreateDirectory(Nombre, Security)
    'If CreateDirectory returns 0, the function has failed
    If ret& = 0 Then
        CrearDirectorio = False
    Else
        CrearDirectorio = True
    End If
    
End Function

Public Sub EliminarDirectorio(ByVal Nombre As String)
    RemoveDirectory Nombre
End Sub




Public Function PropiedadesArchivo(ByVal Filename As String, OwnerhWnd As Long) As Long
        
    '     'open a file properties property page for specified file if return value
    '     '<=32 an error occurred
    '     'From: Delphi code provided by "Ian Land" (iml@dircon.co.uk)
    Dim sei As SHELLEXECUTEINFO
    Dim R As Long
     
    '     'Fill in the SHELLEXECUTEINFO structure
    With sei
        .cbSize = Len(sei)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = Filename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
       
    '     'call the API
    R = ShellExecuteEx(sei)
 
    '     'return the instance handle as a sign of success
    PropiedadesArchivo = sei.hInstApp
       
End Function
Public Function PathArchivo(ByVal Archivo As String) As String

    Dim k As Integer
    
    For k = Len(Archivo) To 1 Step -1
        If Mid$(Archivo, k, 1) = "\" Then
            PathArchivo = VBA.Left$(Archivo, k)
            Exit For
        End If
    Next k
    
End Function

Public Sub RenombrarArchivo(ByVal Antiguo As String, ByVal Nuevo As String)
    MoveFile Antiguo, Nuevo
End Sub

Public Function SysDir() As String

    Dim ret As String
    
    ret = Space$(255)
    
    Call GetSystemDirectory(ret, 255)

    SysDir = StripNulls(ret)

End Function

Public Function ValidPattern(ByVal Patron As String) As Boolean

    Dim k As Integer
    Dim j As Integer
    Dim filtro As String
    
    filtro = "!º\|@·#$%&/()=',¿¡´{}^;><`"
    
    For k = 1 To Len(filtro)
        If InStr(1, Patron, Mid$(filtro, k, 1)) Then
            ValidPattern = False
            Exit Function
        End If
    Next k
    
    ValidPattern = True
    
End Function

Public Function VBArchivoSinPath(ByVal ArchivoConPath As String) As String

    Dim k As Integer
    
    Dim ret As String
    
    ret = ""
    
    For k = Len(ArchivoConPath) To 1 Step -1
        If Mid$(ArchivoConPath, k, 1) = "\" Then
            ret = Mid$(ArchivoConPath, k + 1)
            Exit For
        End If
    Next k
    
    VBArchivoSinPath = ret
    
End Function



Public Function CalcAge(vDate1 As Date, vdate2 As Date) As String

    Dim vYears As Integer, vMonths As Integer, vDays As Integer
    
    vMonths = DateDiff("m", vDate1, vdate2)
    vDays = DateDiff("d", DateAdd("m", vMonths, vDate1), vdate2)
    If vDays < 0 Then
        vMonths = vMonths - 1
        vDays = DateDiff("d", DateAdd("m", vMonths, vDate1), vdate2)
    End If
    vYears = vMonths \ 12
    vMonths = vMonths Mod 12
    
    CalcAge = Abs(vYears)
    
End Function
Public Function KillExcel(ByVal IdProcess As Long) As Boolean

    Call TerminateProcess(IdProcess, 0)
    
End Function

Public Function LeeIni(ByVal Archivo As String, ByVal seccion As String, ByVal Llave As String) As String

    Dim ret As Long
    
    Dim buffer As String
    
    buffer = Space$(1000)
        
    ret = GetPrivateProfileString(seccion, Llave, "", buffer, Len(buffer), Archivo)
    
    buffer = Trim$(buffer)
    buffer = VBA.Left$(buffer, Len(buffer) - 1)
    
    LeeIni = buffer
    
End Function


Public Sub GrabaIni(ByVal ArchivoIni As String, ByVal seccion As String, ByVal Llave As String, ByVal valor)

    Dim ret As Long
    
    ret = WritePrivateProfileString(seccion, Llave, CStr(valor), ArchivoIni)
    
End Sub


Public Function ByteZToStr(ByRef b() As Byte) As String

    'Dim iPos As Long
    'Dim s As String
   's = StrConv(b, vbUnicode)
   'iPos = InStr(s, vbNullChar)
   'If iPos > 0 Then
   '   ByteZToStr = left$(s, iPos - 1)
   'Else
   '   ByteZToStr = s
   'End If
End Function

Public Sub Hourglass(hwnd As Long, fOn As Boolean)

    If fOn Then
        Call SetCapture(hwnd)
        Call SetCursor(LoadCursor(0, ByVal IDC_WAIT))
    Else
        Call ReleaseCapture
        Call SetCursor(LoadCursor(0, IDC_ARROW))
    End If

    DoEvents
    
End Sub
Public Function ArchivoExiste(ByVal Archivo As String) As Boolean

    Dim ret As Boolean
    Dim lRet As Long
    Dim of As OFSTRUCT
    
    ret = False
    
    lRet = OpenFile(Archivo, of, OF_EXIST)
    
    If of.nErrCode = 0 Then ret = True
    
    CloseHandle lRet
    
    ArchivoExiste = ret
    
End Function
'Change the date and time of the file
'so that no one notices whether any file
'is changed or modified using advanced search methods
Public Sub CambiarFechaArchivo(ByVal Archivo As String, ByVal Ano As Integer, ByVal Mes As Integer, ByVal Dia As Integer)
 On Error GoTo errHandle

'dim variables
Dim hFile As Long, rval As Long
Dim buff As OFSTRUCT
Dim ctime As FILETIME, mtime As FILETIME, latime As FILETIME
Dim sTime As SYSTEMTIME
Dim filen As String

Const OF_WRITE = &H1

'attempt to open file already opened for write
filen = CStr(Archivo)
hFile = OpenFile(filen, buff, OF_WRITE)

If hFile Then

    'get original file times
    rval = GetFileTime(hFile, ctime, latime, mtime)

    '-------------
    'create time
    '-------------
    
    'If chkCreated.Value = 1 Then
        
        'convert system to file time
        rval = FileTimeToLocalFileTime(ctime, ctime)
        rval = FileTimeToSystemTime(ctime, sTime)

        'Change filetimes
        sTime.wYear = Ano
        sTime.wMonth = Mes
        sTime.wDay = Dia
        'If chkCurrent.Value = 1 Then
            sTime.wHour = Hour(Time)
            sTime.wMinute = Minute(Time)
            sTime.wSecond = Second(Time)
        'Else
        '    stime.wHour = txtHour.Text
        '    stime.wMinute = txtMin.Text
        '    stime.wSecond = txtSec.Text
        'End If
        
        'reconvert it back and save it
        rval = SystemTimeToFileTime(sTime, ctime)
        rval = LocalFileTimeToFileTime(ctime, ctime)
    
    'End If
    
    
    '-------------
    'last write / modified time
    '-------------
    
    'If chkModified.Value = 1 Then
        
        'convert system to file time
        rval = FileTimeToLocalFileTime(mtime, mtime)
        rval = FileTimeToSystemTime(mtime, sTime)
    
            'Change filetimes
        sTime.wYear = Ano
        sTime.wMonth = Mes
        sTime.wDay = Dia
        'If chkCurrent.Value = 1 Then
            sTime.wHour = Hour(Time)
            sTime.wMinute = Minute(Time)
            sTime.wSecond = Second(Time)
        'Else
        '    stime.wHour = txtHour.Text
        '    stime.wMinute = txtMin.Text
        '    stime.wSecond = txtSec.Text
        'End If
        
        'reconvert it back and save it
        rval = SystemTimeToFileTime(sTime, mtime)
        rval = LocalFileTimeToFileTime(mtime, mtime)
        
    'End If
    
    '-------------
    'last accessed time
    '-------------
    
    'If chkAccessed.Value = 1 Then
    
        'convert system to file time
        rval = FileTimeToLocalFileTime(latime, latime)
        rval = FileTimeToSystemTime(latime, sTime)

        'Change filetimes
        sTime.wYear = Ano
        sTime.wMonth = Mes
        sTime.wDay = Dia

        'If chkCurrent.Value = 1 Then
            sTime.wHour = Hour(Time)
            sTime.wMinute = Minute(Time)
            sTime.wSecond = Second(Time)
        'Else
        '    stime.wHour = txtHour.Text
        '    stime.wMinute = txtMin.Text
        '    stime.wSecond = txtSec.Text
        'End If
        

        'reconvert it back and save it
        rval = SystemTimeToFileTime(sTime, latime)
        rval = LocalFileTimeToFileTime(latime, latime)
    
    'End If


    'and finally write 'em
    rval = SetFileTime(hFile, ctime, latime, mtime)

End If

'close file
rval = CloseHandle(hFile)

'alert and refresh time
'MsgBox "File Calc .exe Saved succesfully!", vbOKOnly + vbInformation, "Succesful Edit"
Exit Sub

errHandle: 'handles errors
    MsgBox "Error #" & Err.Number & ", " & Err.Description & ".", vbCritical, "ERROR!"
    Exit Sub
End Sub

Public Sub CenterForm(Form As Object, Optional MainForm As Object = Nothing)
Attribute CenterForm.VB_Description = "Centra el formulario indicado en la pantalla."
Attribute CenterForm.VB_HelpID = 1000
    Dim auxWorkArea As RECT

    If Not TypeOf Form Is Form Then
        Err.Raise 13
    End If

    If Not MainForm Is Nothing Then
        If Not TypeOf MainForm Is Form Then
            Err.Raise 13
        End If
    End If

    Call SystemParametersInfo(SPI_GETWORKAREA, 0, auxWorkArea, 0)

    If MainForm Is Nothing Then
        Form.Move IIf((auxWorkArea.Right - auxWorkArea.Left) * Screen.TwipsPerPixelX - Form.Width < 0, 0, ((auxWorkArea.Right - auxWorkArea.Left) * Screen.TwipsPerPixelX - Form.Width) \ 2) + auxWorkArea.Left * Screen.TwipsPerPixelX, IIf((auxWorkArea.Bottom - auxWorkArea.Top) * Screen.TwipsPerPixelY - Form.Height < 0, 0, ((auxWorkArea.Bottom - auxWorkArea.Top) * Screen.TwipsPerPixelY - Form.Height) \ 3) + auxWorkArea.Top * Screen.TwipsPerPixelY
    Else
        Form.Move IIf((auxWorkArea.Right - auxWorkArea.Left) * Screen.TwipsPerPixelX - Form.Width < 0, 0, ((auxWorkArea.Right - auxWorkArea.Left) * Screen.TwipsPerPixelX - Form.Width) \ 2) + auxWorkArea.Left * Screen.TwipsPerPixelX, IIf((auxWorkArea.Bottom - auxWorkArea.Top) * Screen.TwipsPerPixelY - MainForm.Height - Form.Height < 0, 0, ((auxWorkArea.Bottom - auxWorkArea.Top) * Screen.TwipsPerPixelY - MainForm.Height - Form.Height) \ 3) + MainForm.Height + auxWorkArea.Top * Screen.TwipsPerPixelY
    End If
End Sub

Public Function Explode(ElementsList As String, Index As Integer, Optional Separator As String = vbTab) As String
Attribute Explode.VB_Description = "Retorna el enésimo elemento de una lista."
Attribute Explode.VB_HelpID = 2000
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
                Element = ""
            End If

            Exit For
        End If

        Element = Mid$(SubStr2Explode, 1, InStr(SubStr2Explode, Separator) - 1)
        SubStr2Explode = Mid$(SubStr2Explode, InStr(SubStr2Explode, Separator) + 1)
    Next auxI

    Explode = Element

    On Error GoTo 0
End Function

Public Sub FillComboBox(ComboBox As Object, ElementsList As String, Optional Separator As String = vbTab)
Attribute FillComboBox.VB_Description = "Llena un objeto ComboBox con el contenido de una lista."
Attribute FillComboBox.VB_HelpID = 3000
    Dim inicio  As Integer
    Dim termino As Integer
    Dim Element As String

    If Not TypeOf ComboBox Is ComboBox Then
        Err.Raise 13
    End If

    On Error Resume Next

    ComboBox.Clear

    inicio = 1
    Do While InStr(inicio, ElementsList, Separator) > 0
        termino = InStr(inicio, ElementsList, Separator)
        Element = ""
        Element = Mid$(ElementsList, inicio, termino - inicio)
        inicio = termino + Len(Separator)
        ComboBox.AddItem Element
    Loop

    Element = ""
    Element = Mid$(ElementsList, inicio)
    If Trim$(Element) <> "" Then ComboBox.AddItem Element

    On Error GoTo 0
End Sub

Public Function Nvl(Value As Variant, Default As Variant) As Variant
Attribute Nvl.VB_Description = "Evalúa una expresión y retorna un valor por defecto en caso de que ésta sea nula. En otro caso retorna el valor de la expresión."
Attribute Nvl.VB_HelpID = 4000
    On Error Resume Next

    If VarType(Value) = vbEmpty Or VarType(Value) = vbError Or VarType(Value) = vbNull Then
        Nvl = Default
    Else
        Nvl = Value
    End If

    On Error GoTo 0
End Function

Public Sub SelectText(Control As Object)
Attribute SelectText.VB_Description = "Selecciona el texto contenido en un TextBox."
Attribute SelectText.VB_HelpID = 6000
    If TypeOf Control Is TextBox Then
        Control.SelStart = 0
        Control.SelLength = Len(Control.Text)
    End If
End Sub

Public Sub Shell(ByVal Archivo As String)
    
    ShellExecute frmMain.hwnd, "open", "excel.exe", GetShortPath(Archivo), "C:", SW_SHOWNORMAL
    
End Sub


Public Function GetShortPath(ByVal strFileName As String) As String
    
    Dim lngRes As Long, strPath As String
    'Create a buffer
    strPath = String$(165, 0)
    'retrieve the short pathname
    lngRes = GetShortPathName(strFileName, strPath, 164)
    'remove all unnecessary chr$(0)'s
    GetShortPath = VBA.Left$(strPath, lngRes)
End Function

Public Function StripPath(ByVal path As String) As String

    If VBA.Right$(path, 1) <> "\" Then
        StripPath = path & "\"
    Else
        StripPath = path
    End If
    
End Function

Public Sub UpperKey(KeyAscii As Integer, Optional ctrl As Object = Nothing)
Attribute UpperKey.VB_Description = "Cambia las letras de un TextBox a mayusculas a medida que estas se van escribiendo."
Attribute UpperKey.VB_HelpID = 7000
    On Error Resume Next

    If KeyAscii = 13 And Not ctrl Is Nothing Then
        KeyAscii = 0
        ctrl.SetFocus
    End If
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

    On Error GoTo 0
End Sub

Public Sub ValidateNumber(KeyAscii As Integer, Optional ctrl As Object = Nothing)
Attribute ValidateNumber.VB_Description = "Permite solo el ingreso de carácteres numéricos en un TextBox."
Attribute ValidateNumber.VB_HelpID = 8000
    On Error Resume Next

    If KeyAscii = 13 And Not ctrl Is Nothing Then
        KeyAscii = 0
        ctrl.SetFocus
    End If
    If (Chr$(KeyAscii) < "0" Or Chr$(KeyAscii) > "9") And KeyAscii <> 8 And KeyAscii <> 0 Then
        KeyAscii = 0
        Beep
    End If

    On Error GoTo 0
End Sub
Public Function ExcelRunning(ByRef IdProceso As Long) As Boolean

    Dim HandleProcesso As Long
    
    HandleSnapshot = CreateToolhelp32Snapshot(&H2, 0)

    InfoProcesso.dwSize = Len(InfoProcesso)

    Processi = Process32First(HandleSnapshot, InfoProcesso)

    If InStr(InfoProcesso.szExeFile, "EXCEL") Then
        IdProceso = OpenProcess(&H1, False, InfoProcesso.th32ProcessID)
        CloseHandle (HandleSnapshot)
        ExcelRunning = True
        Exit Function
    End If
    
    While Processi
        Processi = Process32Next(HandleSnapshot, InfoProcesso)
        
        If InStr(InfoProcesso.szExeFile, "EXCEL") Then
            IdProceso = OpenProcess(&H1, False, InfoProcesso.th32ProcessID)
            CloseHandle (HandleSnapshot)
            ExcelRunning = True
            Exit Function
        End If
    Wend

    CloseHandle (HandleSnapshot)
    
    ExcelRunning = False
    
End Function


Public Function VerificarResolucion(ByVal X As Long, ByVal Y As Long) As Boolean

    #If DESARROLLO = 0 Then
        If GetSystemMetrics(0) = X And GetSystemMetrics(1) = Y Then
            VerificarResolucion = True
        Else
            VerificarResolucion = False
        End If
    #Else
        VerificarResolucion = True
    #End If
    
End Function

Public Function WinDir() As String

    Dim ret As String
    
    ret = Space$(255)
    
    Call GetWindowsDirectory(ret, 255)

    WinDir = StripNulls(ret)

End Function

Public Sub WindowTop(ByVal hwnd As Long)

    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Sub



Public Property Get Orientacion() As Long
    Orientacion = m_Orientacion
End Property

Public Property Let Orientacion(ByVal pOrientacion As Long)
    m_Orientacion = pOrientacion
End Property
