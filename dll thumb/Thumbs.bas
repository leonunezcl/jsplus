Attribute VB_Name = "basThumbs"
Option Explicit

'sDefInitFileName is setup as (AppPath\AppEXEName.Ini)
'and is used as the Default Initialization Filename
Private sDefInitFileName As String

' Maximum long filename path length
Private Const MAX_PATH = 1024

'SendMessage Constants
Private Const BFFM_INITIALIZED = 1
Private Const WM_USER = &H400
Private Const BFFM_SETSELECTIONA = (WM_USER + 102)

'The Following Constants may be passed to BrowseForFolder
'as vTopFolder or vSelPath
Public Const CSIDL_DESKTOP = &H0    'DeskTop
Public Const CSIDL_PROGRAMS = &H2   'Program Groups Folder
Public Const CSIDL_CONTROLS = &H3   'Control Panel Icons Folder
Public Const CSIDL_PRINTERS = &H4   'Printers Folder
Public Const CSIDL_PERSONAL = &H5   'Documents Folder
Public Const CSIDL_FAVORITES = &H6  'Favorites Folder
Public Const CSIDL_STARTUP = &H7    'Startup Folder
Public Const CSIDL_RECENT = &H8     'Recent folder
Public Const CSIDL_SENDTO = &H9     'SendTo Folder
Public Const CSIDL_BITBUCKET = &HA  'Recycle Bin Folder
Public Const CSIDL_STARTMENU = &HB  'Start Menu Folder
Public Const CSIDL_DESKTOPDIRECTORY = &H10  'Windows\Desktop Folder
Public Const CSIDL_DRIVES = &H11    'Devices Virtual Folder (My Computer)
Public Const CSIDL_NETWORK = &H12   'Network Neighborhood Virtual Folder
Public Const CSIDL_NETHOOD = &H13   'Network Neighborhood Folder
Public Const CSIDL_FONTS = &H14     'Fonts Folder
Public Const CSIDL_TEMPLATES = &H15 'ShellNew folder

Private Type SHItemID
    cb      As Long    'Size of the ID (including cb itself)
    abID    As Byte    'The item ID (variable length)
End Type

Private Type ItemIDList
    mkid    As SHItemID
End Type

Private Type BROWSEINFO
    hOwner          As Long
    pidlRoot        As Long
    pszDisplayName  As String
    lpszTitle       As String
    ulFlags         As Long
    lpCallbackProc  As Long
    lParam          As Long
    iImage          As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'Retrieves the location of a special (system) folder.
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ItemIDList) As Long
'ParseDisplayName function should be used instead of this undocumented function.
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    
    Select Case uMsg
        Case BFFM_INITIALIZED
            ' Set the dialog's pre-selected folder using the pidl
            ' set in bi.lParam and passed in the lpData param.
            Call SendMessage(hWnd, BFFM_SETSELECTIONA, False, ByVal lpData)
    
    End Select

End Function

Public Function BrowseForFolder(hOwnerWnd As Long, Optional ByVal sInstruct As String, Optional vSelPath As Variant, Optional vTopFolder As Variant) As String

' Shows the Browse For Folder dialog
'
' hOwnerWnd     (Long)                     OwnerWindow.hWnd.
' sInstruct     (String)                   Instructions for user.
' vSelPath      (String or CSIDL Constant) Pre-select this Folder.
' vTopFolder    (String or CSIDL Constant) Set the Top folder.
'
' If successful, returns the selected folder's full path,
' returns an empty string otherwise.
'

Dim lRet As Long
Dim pidlRet As Long
Dim sPath As String * MAX_PATH
Dim lItemIDList As ItemIDList
Dim uBrowseInfo As BROWSEINFO
    
    With uBrowseInfo
        ' The desktop will own the dialog
        .hOwner = hOwnerWnd
        ' This will be the dialog's root folder.
        If IsMissing(vTopFolder) Then
            vTopFolder = CSIDL_DESKTOP
        End If
        If Len(vTopFolder) > 0 And Not IsNumeric(vTopFolder) Then
            'String Path passed in
            .pidlRoot = SHSimpleIDListFromPath(CStr(vTopFolder))
        Else
            'Long CSIDL Special Folder Constant or Nothing passed in.
            lRet = SHGetSpecialFolderLocation(ByVal hOwnerWnd, ByVal CLng(vTopFolder), lItemIDList)
            .pidlRoot = lItemIDList.mkid.cb
        End If
        ' Set the dialog's prompt string
        .lpszTitle = sInstruct
        ' Obtain and set the address of the callback function
        .lpCallbackProc = FarProc(AddressOf BrowseCallbackProc)
        ' Obtain and set the pidl of the pre-selected folder
        If IsMissing(vSelPath) Then
            'Nothing passed in
            .lParam = .pidlRoot
        ElseIf Len(vSelPath) > 0 And Not IsNumeric(vSelPath) Then
            'String Path passed in
            .lParam = SHSimpleIDListFromPath(CStr(vSelPath))
        Else
            'Long CSIDL Special Folder Constant passed in
            lRet = SHGetSpecialFolderLocation(ByVal hOwnerWnd, ByVal CLng(vSelPath), lItemIDList)
            .lParam = lItemIDList.mkid.cb
        End If
    End With
    
    ' Shows the browse dialog and doesn't return until the dialog is
    ' closed. The BrowseCallbackProc will receive all browse
    ' dialog specific messages while the dialog is open. pidlRet will
    ' contain the pidl of the selected folder if the dialog is not cancelled.
    pidlRet = SHBrowseForFolder(uBrowseInfo)
    
    If pidlRet > 0 Then
        ' Get the path from the selected folder's pidl returned
        ' from the SHBrowseForFolder call (rtns True on success,
        ' sPath must be pre-allocated!)
        If SHGetPathFromIDList(pidlRet, sPath) Then
          ' Return the path
          BrowseForFolder = VBA.Left$(sPath, InStr(sPath, vbNullChar) - 1)
        End If
        ' Free the memory the shell allocated for the pidl.
        Call CoTaskMemFree(pidlRet)
    End If
    
    ' Free the memory the shell allocated for the pre-selected folder.
    Call CoTaskMemFree(uBrowseInfo.lParam)
  
End Function

Public Function FarProc(lpProcName As Long) As Long

'Returns the value of the AddressOf operator
    
    FarProc = lpProcName

End Function

Public Function GetInitEntry(ByVal sSection As String, ByVal sKeyName As String, Optional ByVal sDefault As String = "", Optional ByVal sInitFileName As String = "") As String

'This Function Reads In a String From The Init File.
'Returns Value From Init File or sDefault If No Value Exists.
'sDefault Defaults to an Empty String ("").
'Creates and Uses sDefInitFileName (AppPath\AppEXEName.Ini)
'if sInitFileName Parameter Is Not Passed In.

Dim sBuffer As String
Dim sInitFile As String

    'If Init Filename NOT Passed In
    If Len(sInitFileName) = 0 Then
        'If Static Init FileName NOT Already Created
        If Len(sDefInitFileName) = 0 Then
            'Create Static Init FileName
            sDefInitFileName = App.Path
            If VBA.Right$(sDefInitFileName, 1) <> "\" Then
                sDefInitFileName = sDefInitFileName & "\"
            End If
            sDefInitFileName = sDefInitFileName & App.EXEName & ".ini"
        End If
        sInitFile = sDefInitFileName
    Else    'If Init Filename Passed In
        sInitFile = sInitFileName
    End If
    
    sBuffer = String$(2048, " ")
    GetInitEntry = VBA.Left$(sBuffer, GetPrivateProfileString(sSection, ByVal sKeyName, sDefault, sBuffer, Len(sBuffer), sInitFile))

End Function
Public Function SetInitEntry(ByVal sSection As String, Optional ByVal sKeyName As String, Optional ByVal sValue As String, Optional ByVal sInitFileName As String = "") As Long

'This Function Writes a String To The Init File.
'Returns WritePrivateProfileString Success or Error.
'Creates and Uses sDefInitFileName (AppPath\AppEXEName.Ini)
'if sInitFileName Parameter Is Not Passed In.

'***** CAUTION *****
'If sValue is Null then sKeyName is deleted from the Init File.
'If sKeyName is Null then sSection is deleted from the Init File.

Dim sInitFile As String

    'If Init Filename NOT Passed In
    If Len(sInitFileName) = 0 Then
        'If Static Init FileName NOT Already Created
        If Len(sDefInitFileName) = 0 Then
            'Create Static Init FileName
            sDefInitFileName = App.Path
            If VBA.Right$(sDefInitFileName, 1) <> "\" Then
                sDefInitFileName = sDefInitFileName & "\"
            End If
            sDefInitFileName = sDefInitFileName & App.EXEName & ".ini"
        End If
        sInitFile = sDefInitFileName
    Else    'If Init Filename Passed In
        sInitFile = sInitFileName
    End If
    
    If Len(sKeyName) > 0 And Len(sValue) > 0 Then
        SetInitEntry = WritePrivateProfileString(sSection, ByVal sKeyName, ByVal sValue, sInitFile)
    ElseIf Len(sKeyName) > 0 Then
        SetInitEntry = WritePrivateProfileString(sSection, ByVal sKeyName, vbNullString, sInitFile)
    Else
        SetInitEntry = WritePrivateProfileString(sSection, vbNullString, vbNullString, sInitFile)
    End If

End Function
