VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTydy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_InputFile As String
Private m_OutputFile As String
Private m_ConfigFile As String
Private m_ErrorFile As String
Private m_TidyExe As String

Const INFINITE = &HFFFF
Const STARTF_USESHOWWINDOW = &H1
Private Enum enSW
    SW_HIDE = 0
    SW_NORMAL = 1
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
End Enum
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Enum enPriority_Class
    NORMAL_PRIORITY_CLASS = &H20
    IDLE_PRIORITY_CLASS = &H40
    HIGH_PRIORITY_CLASS = &H80
End Enum
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Function DoTidy() As Boolean
    
    On Error GoTo ErrorDoTidy
    
    Dim PathConfig As String
    Dim OutPutFile As String
    Dim InputFile As String
    Dim ErrorFile As String
    Dim TidyObj As Object
    
    PathConfig = util.GetShortPath(util.PathArchivo(m_ConfigFile)) & util.VBArchivoSinPath(m_ConfigFile)
    OutPutFile = util.GetShortPath(util.PathArchivo(m_OutputFile)) & util.VBArchivoSinPath(m_OutputFile)
    InputFile = util.GetShortPath(util.PathArchivo(m_InputFile)) & util.VBArchivoSinPath(m_InputFile)
    ErrorFile = util.GetShortPath(util.PathArchivo(m_ErrorFile)) & util.VBArchivoSinPath(m_ErrorFile)
    
    Set TidyObj = CreateObject("TidyCOM.TidyObject")
    TidyObj.Options.Load PathConfig
    TidyObj.Options.ErrorFile = ErrorFile
    TidyObj.TidyToFile InputFile, OutPutFile

    Set TidyObj = Nothing
    
    If ArchivoExiste2(m_OutputFile) Or ArchivoExiste2(m_ErrorFile) Then
        DoTidy = True
    Else
        DoTidy = False
    End If
        
    Exit Function
    
ErrorDoTidy:
    MsgBox "DoTidy : " & Err & " Error : " & Error$, vbCritical
    
End Function

Public Sub ExecuteDefault()

    Dim Archivo As String
    Dim arr_tidy_config() As String
    Dim k As Integer
    
    Archivo = util.StripPath(App.Path) & "tidy\tasks.ini"
    
    get_info_section "tasks", arr_tidy_config, Archivo

    For k = 2 To UBound(arr_tidy_config)
        If Len(util.Explode(arr_tidy_config(k), 3, "|")) > 0 Then
            Run util.Explode(arr_tidy_config(k), 2, "|")
        End If
    Next k

End Sub

Public Sub Run(ByVal Archivo As String)

    Dim nFreeFile As Long
    Dim InputFile As String
    Dim k As Integer
    
    If Not frmMain.ActiveForm Is Nothing Then
        If Not frmMain.ActiveForm.Name = "frmEdit" Then Exit Sub
        HTidy.ConfigFile = util.StripPath(App.Path) & "tidy\" & Archivo
        If ArchivoExiste2(HTidy.ConfigFile) Then
            nFreeFile = FreeFile
            InputFile = util.StripPath(App.Path) & "tidy\inputfile.tidy"
            util.BorrarArchivo InputFile
            util.BorrarArchivo util.StripPath(App.Path) & "tidy\output.tidy"
            util.BorrarArchivo util.StripPath(App.Path) & "tidy\errors.tidy"
            
            Open InputFile For Output As #nFreeFile
                With frmMain.ActiveForm
                    For k = 0 To .txtCode.LineCount
                        Print #nFreeFile, .txtCode.GetLine(k)
                    Next k
                End With
            Close #nFreeFile
            
            HTidy.InputFile = InputFile
            
            HTidy.OutPutFile = util.StripPath(App.Path) & "tidy\output.tidy"
            HTidy.ErrorFile = util.StripPath(App.Path) & "tidy\errors.tidy"
            If HTidy.DoTidy Then
                DoEvents
                frmTidyOut.Show vbModal
            End If
        Else
            MsgBox "File not found : " & HTidy.ConfigFile, vbCritical
        End If
    End If
    
End Sub

Private Function SuperShell(ByVal App As String, ByVal WorkDir As String, dwMilliseconds As Long, ByVal start_size As enSW, ByVal Priority_Class As enPriority_Class) As Boolean
    
'    Dim pclass As Long
'    Dim sInfo As STARTUPINFO
'    Dim pinfo As PROCESS_INFORMATION
'    'Not used, but needed
'    Dim sec1 As SECURITY_ATTRIBUTES
'    Dim sec2 As SECURITY_ATTRIBUTES
'    'Set the structure size
'    sec1.nLength = Len(sec1)
'    sec2.nLength = Len(sec2)
'    sInfo.cb = Len(sInfo)
'    'Set the flags
'    sInfo.dwFlags = STARTF_USESHOWWINDOW
'    'Set the window's startup position
'    sInfo.wShowWindow = start_size
'    'Set the priority class
'    pclass = Priority_Class
'    'Start the program
'    If CreateProcess(vbNullString, App, sec1, sec2, False, pclass, _
'        0&, WorkDir, sInfo, pinfo) Then
'        'Wait
'        WaitForSingleObject pinfo.hProcess, dwMilliseconds
'        SuperShell = True
'    Else
'        MsgBox "Can't create process to executes tidy.", vbCritical
'        SuperShell = False
'    End If
    
    Shell ("command.com /c " & App)
    'util.Shell ("command.com /c " & App)
    'util.ShellExFunc(
    SuperShell = True
    
End Function



Public Property Get InputFile() As String
    InputFile = m_InputFile
End Property

Public Property Let InputFile(ByVal pInputFile As String)
    m_InputFile = pInputFile
End Property

Public Property Get OutPutFile() As String
    OutPutFile = m_OutputFile
End Property

Public Property Let OutPutFile(ByVal pOutputFile As String)
    m_OutputFile = pOutputFile
End Property

Public Property Get ConfigFile() As String
    ConfigFile = m_ConfigFile
End Property

Public Property Let ConfigFile(ByVal pConfigFile As String)
    m_ConfigFile = pConfigFile
End Property

Public Property Get ErrorFile() As String
    ErrorFile = m_ErrorFile
End Property

Public Property Let ErrorFile(ByVal pErrorFile As String)
    m_ErrorFile = pErrorFile
End Property

Private Sub Class_Initialize()
    m_TidyExe = util.StripPath(App.Path) & "tidy\tidy.exe"
End Sub


