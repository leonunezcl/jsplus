VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFtpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------
' Module      : CFtpClient Class Module (CFtpClient.cls)
' Type        : Internet application protocol wrapper
' Updated     : 17-OCT-2002
' Version     : 1.0.4
' Author      : Oleg Gdalevich
' Purpose     : Implements the FTP protocol client
' Notes       :
' Dependencies: CScoket.cls         class module
'               MSocketSupport.bas  code module
'               CFtpFile.cls        class module
'               CFtpFiles.cls       class module
'               CTimeOut.cls        class module
'               CTimeOutSupport.bas code module
'               CFtpError.cls       class module
' URL         : http://www.vbip.com/protocols/ftp/vb-ftp-client-library/default.asp
'------------------------------------------------------------------------------------------
' Copyright © 2002 by Oleg Gdalevich
' Visual Basic Internet Programming website (http://www.vbip.com)
'------------------------------------------------------------------------------------------
Option Explicit

Private Enum InternalStateConstants
    istFreeState    'default
    istConnect
    istGetCurrentDirectory
    istSetCurrentDirectory
    istCreateDirectory
    istRemoveDirectory
    istRenameFile
    istDeleteFile
    istDownloadFile
    istUploadFile
    istEnumFiles
    istQuitSession
End Enum

Public Enum FtpSessionStates
    ftpFreeState
    ftpClosed
    ftpConnecting
    ftpConnected
    ftpAuthentication
    ftpUserLoggedIn
    ftpChangingCurrentDirectory
    ftpDeletingFile
    ftpRemovingDirectory
    ftpCreatingDirectory
    ftpRenamingFile
    ftpEstablishingDataConnection
    ftpDataConnectionEstablished
    ftpRetrievingDirectoryInfo
    ftpDirectoryInfoRetrieved
    ftpDownloadInProgress
    ftpDownloadCompleted
    ftpUploadInProgress
    ftpUploadCompleted
    ftpQuitingSession
End Enum

Public Enum AsyncResultStatusConstants
    arStatusOk
    arStatusError
    arStatusTimeOut
    arStatusCancel
End Enum

Public Enum FtpTransferModes
    FTP_ASCII_MODE
    FTP_IMAGE_MODE
End Enum

Public Enum SessionProtocolMessageTypes
    FTP_USER_COMMAND
    FTP_SERVER_RESPONSE
    FTP_SERVER_BAD_RESPONSE
    FTP_APPLICATION_MESSAGE
End Enum

Private Const RESPONSE_CODE_LENGHT = 3
Private Const FTP_CLIENT_BASE_ERROR = 16000

Private m_varInternalState      As InternalStateConstants
Private m_strControlBuffer      As String
Private m_strNewFileName        As String
Private m_strRemoteFileName     As String
Private m_strLocalFileName      As String
Private m_blnOverWriteFile      As Boolean
Private m_intLocalFile          As Integer

Private m_strUserName           As String
Private m_strPassword           As String
Private m_varFtpServer          As Variant
Private m_lngRemotePort         As Long
Private m_strCurrentDirectory   As String
Private m_bPassiveMode          As Boolean
Private m_intTimeOut            As Integer
Private m_strFtpServerMessage   As String
Private m_strLastSentCommand    As String
Private m_TransferMode          As FtpTransferModes
Private m_lngDownloadedBytes    As Long
Private m_lngUploadedBytes      As Long
Private m_strDataSocketBuffer   As String
Private m_objFtpFiles           As CFtpFiles
Private m_strFtpFilesText       As String
Private m_lngStartPositioon     As Long
Private m_lngBytesToUpload      As Long
Private m_varFtpSessionState    As FtpSessionStates
Private m_lngLocalPort          As Long
Private m_strLastServerResponse As String
Private m_intSendBufferLenght   As Integer
Private m_blnTransferComplete   As Boolean
Private arrDataToSend()         As Byte

Private WithEvents m_objTimeOut         As CTimeout
Attribute m_objTimeOut.VB_VarHelpID = -1
Private WithEvents m_objControlSocket   As CSocket
Attribute m_objControlSocket.VB_VarHelpID = -1
Private WithEvents m_objDataSocket      As CSocket
Attribute m_objDataSocket.VB_VarHelpID = -1
Private m_objFtpError                   As CFtpError
'
Public Event OnConnect(ByVal AsyncResultStatus As AsyncResultStatusConstants)
Public Event OnCreateDirectory(ByVal AsyncResultStatus As AsyncResultStatusConstants)
Public Event OnDeleteFile(ByVal AsyncResultStatus As AsyncResultStatusConstants)
Public Event OnGetCurrentDirectory(ByVal AsyncResultStatus As AsyncResultStatusConstants)
Public Event OnSetCurrentDirectory(ByVal AsyncResultStatus As AsyncResultStatusConstants)
Public Event OnRemoveDirectory(ByVal AsyncResultStatus As AsyncResultStatusConstants)
Public Event OnRenameFile(ByVal AsyncResultStatus As AsyncResultStatusConstants)
Public Event OnDownloadFile(ByVal AsyncResultStatus As AsyncResultStatusConstants)
Public Event OnUploadFile(ByVal AsyncResultStatus As AsyncResultStatusConstants)
Public Event OnEnumFiles(ByVal AsyncResultStatus As AsyncResultStatusConstants)
Public Event OnDataTransferProgress(ByVal lngBytesTransferred As Long)
Public Event SessionProtocolMessage(ByVal strMessage As String, ByVal MessageType As SessionProtocolMessageTypes)
Public Event OnStateChange(ByVal SessionState As FtpSessionStates)
Public Event OnQuitSession(ByVal AsyncResultStatus As AsyncResultStatusConstants)
'

Public Sub ClearFileList()
    Set m_objFtpFiles = Nothing
End Sub

Public Sub Connect()
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-13-2002
'Purpose     :Starts the connection establishment process.
'Description :This is an asynchronous method. The completion event is the OnConnect.
'
'             The Connect method performs also the user authentication and
'             retrieving the initial current (working) directory on the FTP server.
'             All the FTP commands for these operations will be sent from the
'             m_objControlSocket_OnDataArrival event handler procedure as soon as
'             the corresponding server responses will be received.
'
'********************************************************************************
    '
    On Error GoTo ERROR_HANDLER
    '
    If Len(m_varFtpServer) = 0 Then
        '
        Err.Raise 380, "CFtpClient.Connect", "Invalid property value."
        '
    Else
        '
        'If no UserName provided by the user,
        'login the user as an anonymous one
        '
        If Len(m_strUserName) = 0 Then
            m_strUserName = "anonymous"
            m_strPassword = "someone@somehost.com"
        End If
        '
        'Change value of the FtpSessionState property
        m_varFtpSessionState = ftpConnecting
        RaiseEvent OnStateChange(ftpConnecting)
        '
        'Reset values of the following vars as they may store
        'something for the previous FTP session.
        m_strLastSentCommand = ""
        m_strControlBuffer = ""
        '
        'Turn the class into the "waiting for server response" state
        Call BeginAsyncMethod(istConnect)
        '
        'Call the Connect method of the CSocket class just to
        'establish a TCP connection with the server.
        m_objControlSocket.Connect m_varFtpServer, m_lngRemotePort
        '
        'If the connection is established successfully, the server sends
        'the welcome message with the response code 220, and then, from the
        'm_objControlSocket_OnDataArrival event procedure the USER FTP command
        'will be sent in order to start the user authentication process.
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    'The most possible error here - sckAlreadyConnected = 10056
    'So be sure that you have closed the previous FTP session
    'before to call the Connect method.
    '
    If m_varInternalState = istConnect Then
        Call EndAsyncMethod
    End If
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.Connect failed: " & .Description
    End With
    '
End Sub

Public Sub SetCurrentDirectory(ByVal strDirectoryPath As String)
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-13-2002
'
'Purpose     :Starts the process of changing the current (working) directory.
'             In terms of FTP it sends the CWD FTP command to the server.
'             CWD = Change Working Directory.
'
'Description :This is an asynchronous method. The completion event is the
'             OnSetCurrentDirectory.
'
'Argument    :The strDirectoryPath contains the name of the directory to change
'             to on the remote system. This can be either a fully qualified path
'             or a name relative to the current directory.
'
'********************************************************************************
    '
    On Error GoTo ERROR_HANDLER
    '
    'If the class is not busy executing another asynchronous method, we can go on.
    If m_varInternalState = istFreeState Then
        '
        If Len(strDirectoryPath) > 0 Then
            '
            'Turn the class into the "waiting for server response" state
            Call BeginAsyncMethod(istSetCurrentDirectory)
            'Send the CWD FTP command to the server
            Call SendFtpCommand("CWD", strDirectoryPath)
            '
        Else
            '
            Err.Raise 5, "CFtpClient.SetCurrentDirectory", "Invalid procedure call or argument"
            '
        End If
        '
    Else
        '
        Err.Raise FTP_CLIENT_BASE_ERROR + 7, "CFtpClient.SetCurrentDirectory", _
                  "Cannot execute the SetCurrentDirectory method. " & _
                  "Waiting fot the server response on the previous command."
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    If m_varInternalState = istSetCurrentDirectory Then
        Call EndAsyncMethod
    End If
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.SetCurrentDirectory failed: " & .Description
    End With
    '
End Sub

Public Sub GetCurrentDirectory()
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-13-2002
'
'Purpose     :Starts the process of retrieving the current (working) directory path.
'             In terms of FTP it sends the PWD FTP command to the server.
'             PWD = Print Working Directory.
'
'Description :This is an asynchronous method. The completion event is the
'             OnGetCurrentDirectory.
'
'             Since this is an asynchronous method, it doesn't return any
'             value. The retrieved directory path will be placed in the
'             CurrentDirectory property that you should read after the
'             OnGetCurrentDirectory event.
'
'********************************************************************************
    '
    On Error GoTo ERROR_HANDLER
    '
    'If the class is not busy executing another asynchronous method, we can go on.
    If m_varInternalState = istFreeState Then
        '
        'Turn the class into the "waiting for server response" state
        Call BeginAsyncMethod(istGetCurrentDirectory)
        'Send the PWD FTP command to the server
        Call SendFtpCommand("PWD", "")
        '
    Else
        '
        Err.Raise FTP_CLIENT_BASE_ERROR + 7, "CFtpClient.GetCurrentDirectory", _
                  "Cannot execute the GetCurrentDirectory method. " & _
                  "Waiting fot the server response on the previous command."
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    If m_varInternalState = istGetCurrentDirectory Then
        Call EndAsyncMethod
    End If
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.GetCurrentDirectory failed: " & .Description
    End With
    '
End Sub

Public Sub CreateDirectory(ByVal strDirectoryPath As String)
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-13-2002
'
'Purpose     :Starts the process of the creating of a new directory.
'             In terms of FTP it sends to the server the MKD FTP command with the
'             new directory path as an argument.
'             MKD = Make Directory.
'
'Description :This is an asynchronous method. The completion event is the
'             OnCreateDirectory.
'
'Argument    :The strDirectoryPath contains the name of the directory to create
'             on the remote system. This can be either a fully qualified path
'             or a name relative to the current directory.
'
'********************************************************************************
    '
    On Error GoTo ERROR_HANDLER
    '
    'If the class is not busy executing another asynchronous method, we can go on.
    If m_varInternalState = istFreeState Then
        '
        If Len(strDirectoryPath) > 0 Then
            '
            'Turn the class into the "waiting for server response" state
            Call BeginAsyncMethod(istCreateDirectory)
            'Send the MKD FTP command to the server
            Call SendFtpCommand("MKD", strDirectoryPath)
            '
        Else
            '
            Err.Raise 5, "CFtpClient.CreateDirectory", "Invalid procedure call or argument"
            '
        End If
        '
    Else
        '
        Err.Raise FTP_CLIENT_BASE_ERROR + 7, "CFtpClient.CreateDirectory", _
                  "Cannot execute the CreateDirectory method. " & _
                  "Waiting fot the server response on the previous command."
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    If m_varInternalState = istCreateDirectory Then
        Call EndAsyncMethod
    End If
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.CreateDirectory failed: " & .Description
    End With
    '
End Sub

Public Sub RemoveDirectory(ByVal strDirectoryPath As String)
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-13-2002
'
'Purpose     :Starts the process of the removing a directory.
'             In terms of FTP it sends to the server the RMD FTP command with the
'             path of the directory to remove.
'             RMD = Remove Directory.
'
'Description :This is an asynchronous method. The completion event is the
'             OnRemoveDirectory.
'
'Argument    :The strDirectoryPath contains the name of the directory to remove
'             on the remote system. This can be either a fully qualified path
'             or a name relative to the current directory.
'
'********************************************************************************
    '
    On Error GoTo ERROR_HANDLER
    '
    'If the class is not busy executing another asynchronous method, we can go on.
    If m_varInternalState = istFreeState Then
        '
        If Len(strDirectoryPath) > 0 Then
            '
            'Turn the class into the "waiting for server response" state
            Call BeginAsyncMethod(istRemoveDirectory)
            'Send the RMD FTP command to the server
            Call SendFtpCommand("RMD", strDirectoryPath)
            '
        Else
            '
            Err.Raise 5, "CFtpClient.RemoveDirectory", "Invalid procedure call or argument"
            '
        End If
        '
    Else
        '
        Err.Raise FTP_CLIENT_BASE_ERROR + 7, "CFtpClient.RemoveDirectory", _
                  "Cannot execute the RemoveDirectory method. " & _
                  "Waiting fot the server response on the previous command."
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    If m_varInternalState = istRemoveDirectory Then
        Call EndAsyncMethod
    End If
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.RemoveDirectory failed: " & .Description
    End With
    '
End Sub

Public Sub DeleteFile(ByVal strFilePath As String)
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-13-2002
'
'Purpose     :Starts the process of the deleting a file on the remote system.
'             In terms of FTP it sends to the server the DELE FTP command with the
'             path of the file to delete as an argument.
'             DELE = Delete.
'
'Description :This is an asynchronous method. The completion event is the
'             OnDeleteFile.
'
'Argument    :The strFilePath contains the path to the file to delete
'             on the remote system. This can be either a fully qualified path
'             or a name relative to the current directory.
'
'********************************************************************************
    '
    On Error GoTo ERROR_HANDLER
    '
    'If the class is not busy executing another asynchronous method, we can go on.
    If m_varInternalState = istFreeState Then
        '
        If Len(strFilePath) > 0 Then
            '
            'Turn the class into the "waiting for server response" state
            Call BeginAsyncMethod(istDeleteFile)
            'Send the DELE FTP command to the server
            Call SendFtpCommand("DELE", strFilePath)
            '
        Else
            '
            Err.Raise 5, "CFtpClient.DeleteFile", "Invalid procedure call or argument"
            '
        End If
        '
    Else
        '
        Err.Raise FTP_CLIENT_BASE_ERROR + 7, "CFtpClient.DeleteFile", _
                  "Cannot execute the DeleteFile method. " & _
                  "Waiting fot the server response on the previous command."
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    If m_varInternalState = istDeleteFile Then
        Call EndAsyncMethod
    End If
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.DeleteFile failed: " & .Description
    End With
    '
End Sub

Public Sub RenameFile(ByVal strOldFileName As String, ByVal strNewFileName As String)
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-13-2002
'
'Purpose     :Starts the process of the renaming a file on the remote system.
'             In order to rename a file on the FTP server, we need to send
'             two commands - the RNFR to specify the file to rename, and then,
'             after receiving successful response from the server, the RNTO one
'             to specify the new file name.
'             RNFR = Rename From.
'             RNTO = Rename To.
'
'Description :This is an asynchronous method. The completion event is the
'             OnRenameFile.
'
'Argument    :strOldFileName - path to the file to rename
'             strNewFileName - new name for the file
'
'             Both arguments can be either fully qualified paths
'             or names relative to the current directory.
'
'********************************************************************************
    '
    On Error GoTo ERROR_HANDLER
    '
    'If the class is not busy executing another asynchronous method, we can go on.
    If m_varInternalState = istFreeState Then
        '
        If Len(strOldFileName) > 0 And Len(strNewFileName) > 0 Then
            '
            m_strNewFileName = strNewFileName
            '
            'Turn the class into the "waiting for server response" state
            Call BeginAsyncMethod(istRenameFile)
            'Send the RNFR FTP command to the server
            Call SendFtpCommand("RNFR", strOldFileName)
            '
            'If the successful response will be received on this command,
            'the RNTO command will be sent from the m_objControlSocket_OnDataArrival
            'event handler procedure.
            '
        Else
            '
            Err.Raise 5, "CFtpClient.RenameFile", "Invalid procedure call or argument"
            '
        End If
        '
    Else
        '
        Err.Raise FTP_CLIENT_BASE_ERROR + 7, "CFtpClient.RenameFile", _
                  "Cannot execute the RenameFile method. " & _
                  "Waiting fot the server response on the previous command."
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    If m_varInternalState = istRenameFile Then
        Call EndAsyncMethod
    End If
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.RenameFile failed: " & .Description
    End With
    '
End Sub

Public Sub DownloadFile(ByVal strRemoteFilePath As String, ByVal strLocalFilePath As String, ByVal OverWrite As Boolean)
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-13-2002
'
'Purpose     :Starts the process of the downloading of a file from the server to
'             the local system. In terms of FTP it sends to the server the RETR
'             FTP command with the path of the file to download as an argument.
'
'             But before sending the RETR command the data connection must be
'             established and other options of the data transfer should be
'             defined. The client specifies the data type with the TYPE FTP
'             command and sends the REST command in the case when it needs
'             to resume the previously broken data transfer.
'
'             TYPE = Type
'             RETR = Retrieve
'             REST = Restart
'
'Description :This is an asynchronous method. The completion event is the
'             OnDeleteFile.
'
'Arguments   :The strRemoteFilePath contains the path to the file to retrieve
'             on the remote system. This can be either a fully qualified path
'             or a name relative to the current directory.
'
'             The strLocalFilePath argument specifies the path on the local
'             file system where the downloaded file will be stored.
'
'             The OverWrite argument specifies the behavior of the CFtpClient class
'             in the case when the local file with the same name is already exists.
'
'********************************************************************************
    '
    Dim strMode As String
    '
    On Error GoTo ERROR_HANDLER
    '
    'If the class is not busy executing another asynchronous method, we can go on.
    If m_varInternalState = istFreeState Then
        '
        'Both arguments must be initialized, otherwise the error
        '"Invalid procedure call or argument" occurs.
        '
        If Len(strRemoteFilePath) > 0 And Len(strLocalFilePath) > 0 Then
            '
            'Store the arguments values in the module level variables.
            m_strRemoteFileName = strRemoteFilePath
            m_strLocalFileName = strLocalFilePath
            m_blnOverWriteFile = OverWrite
            'Reset the downloaded bytes counter.
            m_lngDownloadedBytes = 0
            '
            'Read the TransferMode property of the class in order to
            'get the argument for the TYPE FTP command.
            strMode = CStr(IIf(m_TransferMode = FTP_ASCII_MODE, "A", "I"))
            '
            'Turn the class into the "waiting for server response" state
            Call BeginAsyncMethod(istDownloadFile)
            'Send the TYPE FTP command to the server
            Call SendFtpCommand("TYPE", strMode)
            '
            'If the successful response will be received on the TYPE command,
            'the PORT or PASV command will be sent from the m_objControlSocket_OnDataArrival
            'event handler procedure in order to establish the data connection.
            '
            'If restarting of data transfer is needed the REST command will be sent
            'later, from the OpenLocalFile subroutine which is called before sending
            'the RETR FTP command.
            '
        Else
            '
            Err.Raise 5, "CFtpClient.DownloadFile", "Invalid procedure call or argument"
            '
        End If
        '
    Else
        '
        Err.Raise FTP_CLIENT_BASE_ERROR + 7, "CFtpClient.DownloadFile", _
                  "Cannot execute the DownloadFile method. " & _
                  "Waiting fot the server response on the previous command."
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    If m_varInternalState = istDownloadFile Then
        Call EndAsyncMethod
    End If
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.DownloadFile failed: " & .Description
    End With
    '
End Sub

Public Sub EnumFiles()
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-13-2002
'
'Purpose     :Starts the process of the retrieving the current directory listing
'             on the remote system. To perform this operation the LIST FTP command
'             is used. The LIST command retrieves name, size, last write time and
'             other properties of every file and subdirectory located in the current
'             directory.
'
'             But before sending the LIST command the data type should be specified
'             with the TYPE FTP command and the data connection must be established
'             as directory listings are transferred only via the data connection.
'
'Description :This is an asynchronous method. The completion event is the
'             OnEnumFiles. As soon as this event has been occurred you can
'             read the current directory listing with the CurrentDirectoryFiles
'             property of the CFtpClient class. This property is an instance of
'             the CFtpFiles collection. Each item of that collection is an instance
'             of the CFtpFiles class which represents a single file or subdirectory.
'
'********************************************************************************
    '
    On Error GoTo ERROR_HANDLER
    '
    'If the class is not busy executing another asynchronous method, we can go on.
    If m_varInternalState = istFreeState Then
        '
        'Turn the class into the "waiting for server response" state
        Call BeginAsyncMethod(istEnumFiles)
        'Send the TYPE FTP command to the server
        Call SendFtpCommand("TYPE", "A")
        '
    Else
        '
        Err.Raise FTP_CLIENT_BASE_ERROR + 7, "CFtpClient.EnumFiles", _
                  "Cannot execute the EnumFiles method. " & _
                  "Waiting fot the server response on the previous command."
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    If m_varInternalState = istEnumFiles Then
        Call EndAsyncMethod
    End If
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.EnumFiles failed: " & .Description
    End With
    '
End Sub

Public Sub UploadFile(ByVal strLocalFile As String, ByVal strRemoteFile As String, ByVal lngStartPosition As Long)
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-13-2002
'
'Purpose     :Starts the process of the uploading a file from the local system
'             to the FTP server. In terms of FTP it sends to the server the STOR
'             FTP command with the path to the file to store on the remote system
'             as an argument.
'
'             But before sending the STOR command the data connection must be
'             established and other options of the data transfer should be
'             defined. The client specifies the data type with the TYPE FTP
'             command and sends the REST command in the case when it needs
'             to resume the previously broken data transfer.
'
'             TYPE = Type
'             STOR = Store
'             REST = Restart
'
'Description :This is an asynchronous method. The completion event is the
'             OnUploadFile.
'
'Arguments   :The strLocalFilePath argument specifies the path to the file
'             on the local system that is going to be uploaded.
'
'             The strRemoteFilePath contains the path to the file to store
'             on the remote system. This can be either a fully qualified path
'             or a name relative to the current directory.
'
'             If the file with the specified name already exists on the remote
'             system and the lngStartPosition argument is not 0, the CFtpClient
'             class skips the part of the local file and starts the reading
'             and transfer of the file data at the specified position.
'
'********************************************************************************
    '
    Dim strMode As String
    '
    On Error GoTo ERROR_HANDLER
    '
    'If the class is not busy executing another asynchronous method, we can go on.
    If m_varInternalState = istFreeState Then
        '
        'Both arguments must be initialized, otherwise the error
        '"Invalid procedure call or argument" occurs.
        '
        If Len(strRemoteFile) > 0 And Len(strLocalFile) > 0 Then
            '
            'Store the arguments values in the module level variables.
            m_strRemoteFileName = strRemoteFile
            m_strLocalFileName = strLocalFile
            m_lngStartPositioon = lngStartPosition
            '
            ReDim arrDataToSend(m_intSendBufferLenght - 1)
            '
            'Read the TransferMode property of the class in order to
            'get the argument for the TYPE FTP command.
            strMode = CStr(IIf(m_TransferMode = FTP_ASCII_MODE, "A", "I"))
            '
            'Turn the class into the "waiting for server response" state
            Call BeginAsyncMethod(istUploadFile)
            'Send the TYPE FTP command to the server
            Call SendFtpCommand("TYPE", strMode)
            '
        Else
            '
            Err.Raise 5, "CFtpClient.UploadFile", "Invalid procedure call or argument"
            '
        End If
        '
    Else
        '
        Err.Raise FTP_CLIENT_BASE_ERROR + 7, "CFtpClient.UploadFile", _
                  "Cannot execute the UploadFile method. " & _
                  "Waiting fot the server response on the previous command."
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    If m_varInternalState = istUploadFile Then
        Call EndAsyncMethod
    End If
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.UploadFile failed: " & .Description
    End With
    '
End Sub

Public Sub QuitSession()
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-13-2002
'
'Purpose     :Starts the process of quiting from the current FTP session.
'             In terms of FTP it sends the QUIT FTP command to the server.
'
'Description :This is an asynchronous method. The completion event is the
'             OnQuitSession.
'
'             As the QUIT command is received from the client, the FTP server
'             must send the reply code 221 and close the control connection.
'
'********************************************************************************
    '
    On Error GoTo ERROR_HANDLER
    '
    'If the class is not busy executing another asynchronous method, we can go on.
    If m_varInternalState = istFreeState Then
        '
        'Turn the class into the "waiting for server response" state
        Call BeginAsyncMethod(istQuitSession)
        'Send the QUIT FTP command to the server
        Call SendFtpCommand("QUIT", "")
        '
    Else
        '
        Err.Raise FTP_CLIENT_BASE_ERROR + 7, "CFtpClient.QuitSession", _
                  "Cannot execute the QuitSession method. " & _
                  "Waiting fot the server response on the previous command."
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    If m_varInternalState = istQuitSession Then
        Call EndAsyncMethod
    End If
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.QuitSession failed: " & .Description
    End With
    '
End Sub

Private Sub CloseDataConnection()
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-13-2002
'Purpose     :Closes the socket responsible for the data connection.
'********************************************************************************
    '
    On Error GoTo ERROR_HANDLER
    '
    Dim lngRetValue As Long
    '
    If Not m_objDataSocket.State = sckClosed Then
        '
        'Some FTP servers ignore the TCP FIN segment sent by a client
        'to the data connection. This is exactly what the CSocket does
        'when the CloseSocket method is called:
        '
        'shutdown(m_lngSocketHandle, SD_SEND)
        '
        'the CSocket class calls the closesocket Winsock API function
        'only if the reciprocal FIN TCP segment is received from the server.
        'But some servers do not send it, and continue sending data. Imagine
        'you are going to cancel 700 MB file transfer due to the slow speed
        'of the data transfer - you could wait forever.
        '
        'So, here is a litle interference in the CSocket class implementation:
        '
        lngRetValue = shutdown(m_objDataSocket.SocketHandle, SD_BOTH)
        '
        'With such a call of the shutdown Winsock API function the server
        'will have got the RST TCP segment in return on any chunk of data
        'sent once we have called this method.
        '
        m_objDataSocket.CloseSocket
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.CloseDataConnection failed: " & .Description
    End With
    '
End Sub

Public Sub CancelAsyncMethod()
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-13-2002
'
'Purpose     :Cancels any asynchronous method performed by the class.
'
'             In other words, takes the class off the waiting state - stops the
'             timer and changes a value of the m_varInternalState variable
'             with the call of the EndAsyncMethod subroutine, and then raises
'             the appropriate event.
'
'             If any data transfer is in progress, this method closes the data
'             connection calling the CloseDataConnection subroutine. If that
'             data transfer is either the downloading or uploading operation,
'             the method closes the local file using the CloseLocalFile.
'
'********************************************************************************
    '
    Dim varIntState As InternalStateConstants
    '
    On Error GoTo ERROR_HANDLER
    '
    varIntState = m_varInternalState
    '
    Call EndAsyncMethod
    '
    Select Case varIntState
        '
        Case istConnect
            RaiseEvent OnConnect(arStatusCancel)
        Case istCreateDirectory
            RaiseEvent OnCreateDirectory(arStatusCancel)
        Case istDeleteFile
            RaiseEvent OnDeleteFile(arStatusCancel)
        Case istGetCurrentDirectory
            RaiseEvent OnGetCurrentDirectory(arStatusCancel)
        Case istRemoveDirectory
            RaiseEvent OnRemoveDirectory(arStatusCancel)
        Case istRenameFile
            RaiseEvent OnRenameFile(arStatusCancel)
        Case istSetCurrentDirectory
            RaiseEvent OnSetCurrentDirectory(arStatusCancel)
        Case istEnumFiles
            m_strDataSocketBuffer = ""
            Call CloseDataConnection
            RaiseEvent OnEnumFiles(arStatusCancel)
        Case istDownloadFile
            Call CloseDataConnection
            Call CloseLocalFile
            RaiseEvent OnDownloadFile(arStatusCancel)
        Case istUploadFile
            Call CloseDataConnection
            Call CloseLocalFile
            RaiseEvent OnUploadFile(arStatusCancel)
        '
    End Select
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.CancelAsyncMethod failed: " & .Description
    End With
    '
End Sub

Private Sub SendFtpCommand(strCommand As String, strArguments As String)
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-13-2002
'Purpose     :Sends an FTP command to the FTP server.
'Arguments   :strCommand   - the FTP command to send
'             strArguments - any arguments to send with the command
'Description :This subroutine also performs a number of important tasks.
'
'             Depending on the command that is going to be sent, it changes a value
'             of the FtpSessionState property and raises the OnStateChange event.
'
'             It clears the module level buffer m_strControlBuffer that stores
'             data received via the control connection.
'
'             Stores the sent command in the module level variable m_strLastSentCommand,
'             which is perhaps the most significant variable in this class ;-)
'
'             And finally, this sub raises the SessionProtocolMessage event.
'
'********************************************************************************
    '
    Dim strDataToSend As String
    Dim varNewState   As FtpSessionStates
    '
    On Error GoTo ERROR_HANDLER
    '
    If Not m_objControlSocket.State = sckConnected Then
        Err.Raise FTP_CLIENT_BASE_ERROR + 8, "CFtpClient.SendFtpCommand", _
                  "The connection is not established."
        Exit Sub
    End If
    '
    'Clear the control connection receive buffer
    m_strControlBuffer = ""
    '
    varNewState = ftpFreeState
    '
    Select Case strCommand
        Case "USER", "PASS"
            varNewState = ftpAuthentication
        Case "CWD"
            varNewState = ftpChangingCurrentDirectory
        Case "LIST", "PWD"
            varNewState = ftpRetrievingDirectoryInfo
        Case "MKD"
            varNewState = ftpCreatingDirectory
        Case "DELE"
            varNewState = ftpDeletingFile
        Case "RETR"
            varNewState = ftpDownloadInProgress
        Case "RMD"
            varNewState = ftpRemovingDirectory
        Case "RNFR"
            varNewState = ftpRenamingFile
        Case "STOR"
            varNewState = ftpUploadInProgress
        Case "QUIT"
            varNewState = ftpQuitingSession
        Case "TYPE", "PASV", "PORT", "REST"
            Select Case m_varInternalState
                Case istDownloadFile: varNewState = ftpDownloadInProgress
                Case istUploadFile: varNewState = ftpUploadInProgress
                Case istEnumFiles: varNewState = ftpRetrievingDirectoryInfo
            End Select
    End Select
    '
    m_varFtpSessionState = varNewState
    RaiseEvent OnStateChange(varNewState)
    '
    'Build command to send to the server
    If Len(strArguments) > 0 Then
        strDataToSend = strCommand & " " & strArguments & vbCrLf
    Else
        strDataToSend = strCommand & vbCrLf
    End If
    '
    'Remember the command
    m_strLastSentCommand = strCommand
    '
    'Send the command to the server
    m_objControlSocket.SendData strDataToSend
    '
    'Let the client application know what we have sent.
    RaiseEvent SessionProtocolMessage(strCommand & " " & strArguments, FTP_USER_COMMAND)
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.SendFtpCommand failed: " & .Description
    End With
    '
End Sub

Private Sub RaiseFtpError(ByVal intReplyCode As Integer)
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-14-2002
'Purpose     :Raises the error type completion event.
'Arguments   :intReplyCode - the FTP server reply code
'Description :This procedure calls the EndAsyncMethod subroutine in order to
'             stop the timer and turn off the m_varInternalState flag variable.
'             The error information is passed to the client application with the
'             internal (CFtpError) error object the properties of which will
'             be initialized here.
'********************************************************************************
    '
    Dim strErrorDescription As String
    Dim varInternalState    As InternalStateConstants
    '
    On Error GoTo ERROR_HANDLER
    '
    varInternalState = m_varInternalState
    '
    Call EndAsyncMethod
    '
    Select Case intReplyCode
        '
        Case 120: strErrorDescription = "Service ready in nnn minutes."
        Case 332: strErrorDescription = "Need account for login."
        Case 350: strErrorDescription = "Requested file action pending further information."
        Case 421: strErrorDescription = "Service not available, closing control connection."
        Case 425: strErrorDescription = "Can 't open data connection."
        Case 426: strErrorDescription = "Connection closed; transfer aborted."
        Case 450: strErrorDescription = "Requested file action not taken."
        Case 451: strErrorDescription = "Requested action aborted: local error in processing."
        Case 452: strErrorDescription = "Requested action not taken."
        Case 500: strErrorDescription = "Syntax error, command unrecognized."
        Case 501: strErrorDescription = "Syntax error in parameters or arguments."
        Case 502: strErrorDescription = "Command not implemented."
        Case 503: strErrorDescription = "Bad sequence of commands."
        Case 504: strErrorDescription = "Command not implemented for that parameter."
        Case 530: strErrorDescription = "Not logged in."
        Case 532: strErrorDescription = "Need account for storing files."
        Case 550: strErrorDescription = "Requested action not taken."
        Case 551: strErrorDescription = "Requested action aborted: page type unknown."
        Case 552: strErrorDescription = "Requested file action aborted."
        Case 553: strErrorDescription = "Requested action not taken."
        '
    End Select
    '
    With m_objFtpError
        .Number = FTP_CLIENT_BASE_ERROR + intReplyCode
        .Description = strErrorDescription
        .Source = "FtpClientX.CFtpClient"
    End With
    '
    Call RaiseErrorCompletionEvent(varInternalState)
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.RaiseFtpError failed: " & .Description
    End With
    '
End Sub

Public Property Get DataTransferInProgress() As Boolean
Attribute DataTransferInProgress.VB_Description = "Returns True, if the data connection is opened."
    '
    Select Case m_varInternalState
        '
        Case istUploadFile, istEnumFiles, istDownloadFile
            '
            DataTransferInProgress = True
            '
    End Select
    '
End Property

Public Property Get Busy() As Boolean
Attribute Busy.VB_Description = "Indicates availability of the connection for a new operation."
    '
    If (m_varFtpSessionState = ftpClosed) Or _
       (m_varFtpSessionState = ftpFreeState) Then
        '
        Busy = False
        '
    Else
        '
        Busy = True
        '
    End If
    '
End Property

Public Property Get CurrentDirectoryFiles() As CFtpFiles
Attribute CurrentDirectoryFiles.VB_Description = "Returns an instance of the CFtpFiles collection.  Each member of the collection represents information about a single file or subdirectory in the current (working) directory."
    Set CurrentDirectoryFiles = m_objFtpFiles
End Property

Public Property Get CurrentDirectoryListing() As String
Attribute CurrentDirectoryListing.VB_Description = "Returns the current directory listing as row text data."
    CurrentDirectoryListing = m_strFtpFilesText
End Property

Public Property Get CurrentDirectory() As String
Attribute CurrentDirectory.VB_Description = "Returns path to the current (working) directory."
    CurrentDirectory = m_strCurrentDirectory
End Property

Public Property Get FtpServerMessage() As String
    FtpServerMessage = m_strFtpServerMessage
    m_strFtpServerMessage = ""
End Property

Public Property Get FtpServer() As Variant
Attribute FtpServer.VB_Description = "Sets or returns the address of FTP server."
    FtpServer = m_varFtpServer
End Property

Public Property Let FtpServer(NewValue As Variant)
    m_varFtpServer = NewValue
End Property

Public Property Get Password() As String
Attribute Password.VB_Description = "Returns or sets the password that will be sent to FTP server for the user authentication."
    Password = m_strPassword
End Property

Public Property Let Password(NewValue As String)
    m_strPassword = NewValue
End Property

Public Property Get UserName() As String
Attribute UserName.VB_Description = "Returns or sets the user name that will be sent to FTP server for the user authentication."
    UserName = m_strUserName
End Property

Public Property Let UserName(NewValue As String)
    m_strUserName = NewValue
End Property

Public Property Get PassiveMode() As Boolean
Attribute PassiveMode.VB_Description = "Sets or returns wherever the passive mode is used in order to establish the data connection."
    PassiveMode = m_bPassiveMode
End Property

Public Property Let PassiveMode(NewValue As Boolean)
    m_bPassiveMode = NewValue
End Property

Public Property Get SendBufferLenght() As Integer
    SendBufferLenght = m_intSendBufferLenght
End Property

Public Property Let SendBufferLenght(NewValue As Integer)
    '
    If NewValue < 512 Or NewValue > 8192 Then
        Err.Raise 380, "CFtpClient.SendBufferLenght", "Invalid property value."
    Else
        m_intSendBufferLenght = NewValue
    End If
    '
End Property

Public Property Get RemotePort() As Long
Attribute RemotePort.VB_Description = "The port number where you believe the FTP service is running on. The default value is 21."
    RemotePort = m_lngRemotePort
End Property

Public Property Let RemotePort(NewValue As Long)
    '
    If NewValue < 1 Or NewValue > 65535 Then
        'Err.Raise 380, "CFtpClient.RemotePort", "Invalid property value."
        m_lngRemotePort = 21
    Else
        m_lngRemotePort = NewValue
    End If
    '
End Property

Public Property Get TransferMode() As FtpTransferModes
Attribute TransferMode.VB_Description = "Returns or sets the data transfer mode."
    TransferMode = m_TransferMode
End Property

Public Property Let TransferMode(NewValue As FtpTransferModes)
    m_TransferMode = NewValue
End Property

Public Property Let TimeOut(ByVal intSeconds As Integer)
Attribute TimeOut.VB_Description = "Returns or sets the period, in seconds, to wait before a time-out expires."
    m_objTimeOut.TimeOutValue = intSeconds
End Property

Public Property Get TimeOut() As Integer
    TimeOut = m_objTimeOut.TimeOutValue
End Property

Public Property Get FtpSessionState() As FtpSessionStates
Attribute FtpSessionState.VB_Description = "Indicates the current state of the FTP session."
    FtpSessionState = m_varFtpSessionState
End Property

Public Property Get LastServerResponse() As String
    LastServerResponse = m_strLastServerResponse
End Property

Public Property Get LastFtpError() As CFtpError
    Set LastFtpError = m_objFtpError
End Property

Private Sub Class_Initialize()
    '
    Set m_objControlSocket = New CSocket
    Set m_objDataSocket = New CSocket
    Set m_objTimeOut = New CTimeout
    Set m_objFtpError = New CFtpError
    '
    m_lngRemotePort = 21            'default port number
    m_objTimeOut.TimeOutValue = 60  'default timout value
    m_TransferMode = FTP_IMAGE_MODE 'default transfer mode
    m_intSendBufferLenght = 1460
    m_varFtpSessionState = ftpClosed
    '
End Sub

Private Sub Class_Terminate()
    '
    m_objTimeOut.StopTimer
    '
    Set m_objControlSocket = Nothing
    Set m_objDataSocket = Nothing
    Set m_objTimeOut = Nothing
    Set m_objFtpError = Nothing
    '
End Sub

Private Sub m_objControlSocket_OnClose()
    '
    On Error Resume Next
    '
    'The server is going to close the control connection
    '
    'If an asynchronous method is in progress - stop it
    If Not m_varInternalState = istFreeState Then
        Call EndAsyncMethod
    End If
    '
    'Let the client application know that the connection is closed
    m_varFtpSessionState = ftpClosed
    RaiseEvent OnStateChange(ftpClosed)
    '
End Sub

Private Sub m_objControlSocket_OnConnect()
    '
    'The control connection was established successfully
    '
    'Reset the timer - don't stop as the Connect method is
    'still at work authenticating the user and retrieving
    'the initial directory name
    m_objTimeOut.Reset
    '
    'Let the client application know that we have connected
    m_varFtpSessionState = ftpConnected
    RaiseEvent OnStateChange(ftpConnected)
    '
End Sub

Private Sub m_objControlSocket_OnDataArrival(ByVal bytesTotal As Long)
    '
    'This is the heart of the class. This event handler
    'receives server responses, parses them in order to
    'retrieve reply codes, and depending on the m_varInternalState
    'value performs different tasks.
    '
    Dim strArrivedData  As String
    Dim intReplyCode    As Integer
    '
    On Error GoTo ERROR_HANDLER
    '
    If Not m_varInternalState = istFreeState Then
        m_objTimeOut.Reset
    End If
    '
    'Retrieve data from the Winsock buffer
    m_objControlSocket.GetData strArrivedData
    '
    'Initialize the LastServerResponse property
    m_strLastServerResponse = strArrivedData
    '
    'Collect incoming data in the local buffer
    m_strControlBuffer = m_strControlBuffer & strArrivedData
    '
    'try to retrieve the reply code
    intReplyCode = GetResponseCode(m_strControlBuffer)
    '
    'If the server response doesn't contain a reply code,
    'just go away from this procedure
    Debug.Print m_strLastSentCommand & " : " & intReplyCode
    If intReplyCode = 0 Then Exit Sub
    '
    'Send the session protocol message to the client app
    If intReplyCode > 399 Then
        RaiseEvent SessionProtocolMessage(Left(strArrivedData, Len(strArrivedData) - 2), FTP_SERVER_BAD_RESPONSE)
    Else
        RaiseEvent SessionProtocolMessage(Left(strArrivedData, Len(strArrivedData) - 2), FTP_SERVER_RESPONSE)
    End If
    '
    Select Case m_varInternalState
        '
        Case istFreeState
            '
            'We have no asynchronous method in progress, but the server can
            'send some data via the control connection. For example, it can
            'close the session sending the reply code 421
            '
            If intReplyCode = 421 Then
                Call CloseControlConnection
            End If
            '
        Case istQuitSession
            '
            'The QuitSession method is in progress
            '
            Select Case m_strLastSentCommand
                '
                Case "QUIT"
                    '
                    'The QUIT FTP command has been sent by the QuitSession method.
                    '
                    Select Case intReplyCode
                        Case 221
                            Call EndAsyncMethod
                            m_objControlSocket.CloseSocket
                            RaiseEvent OnQuitSession(arStatusOk)
                        Case Else
                            Call RaiseFtpError(intReplyCode)
                    End Select
                    '
            End Select
            '
        Case istConnect
            '
            'The Connect method is in progress
            '
            Select Case m_strLastSentCommand
                '
                Case ""
                    'We have not sent any command yet
                    Select Case intReplyCode
                        '
                        Case 220
                            'We have got a welcome message
                            'Start the user authentication process
                            SendFtpCommand "USER", m_strUserName
                        Case Else
                            '120, 421
                            RaiseFtpError intReplyCode
                            Call CloseControlConnection
                        '
                    End Select
                    '
                Case "USER"
                    '
                    Select Case intReplyCode
                        '
                        Case 230
                            '230 reply code means that the user is
                            'already logged in without providing the password.
                            'Start the process of the retriving of the
                            'initial directory name
                            m_varFtpSessionState = ftpUserLoggedIn
                            RaiseEvent OnStateChange(ftpUserLoggedIn)
                            Call SendFtpCommand("PWD", "")
                            '
                        Case 331
                            'The server asks for the password - send it.
                            SendFtpCommand "PASS", m_strPassword
                        Case Else
                            '332, 421, 500, 501, 530
                            'Any of these reply codes means that the
                            'authentication is failed - raise the event
                            'and close the socket with the CloseControlConnection
                            RaiseFtpError intReplyCode
                            Call CloseControlConnection
                        '
                    End Select
                    '
                Case "PASS"
                        '
                    Select Case intReplyCode
                        '
                        Case 202, 230
                            '
                            'The password has been accepted - the user is logged in!
                            m_varFtpSessionState = ftpUserLoggedIn
                            RaiseEvent OnStateChange(ftpUserLoggedIn)
                            'Start the process of the retrieving of
                            'the initial directory name
                            Call SendFtpCommand("PWD", "")
                            '
                        Case Else
                            '332, 421, 500, 501, 503, 530
                            'Password is incorrect - raise the completion event
                            'and close the socket with CloseControlConnection
                            RaiseFtpError intReplyCode
                            Call CloseControlConnection
                        '
                    End Select
                    '
                Case "PWD"
                    'Retrieving of the initial directory name
                    Select Case intReplyCode
                        '
                        Case 257
                            'The server respose should contain the directory path
                            'The GetDirFromResponse function retrieves the current
                            'directory path and initializes the CurrentDirectory property
                            If GetDirFromResponse Then
                                'The directory name was retrieved successfully
                                'End the method and raise the completion event.
                                Call EndAsyncMethod
                                RaiseEvent OnConnect(arStatusOk)
                            Else
                                'Unexpected server response. GetDirFromResponse
                                'can't parse the respose string in oreder to
                                'retrieve the directory path.
                                'Raise the error and close the socket.
                                m_objFtpError.Number = FTP_CLIENT_BASE_ERROR + 1
                                m_objFtpError.Description = "Cannot retrieve the current directory path. Unexpected server response."
                                Call CloseControlConnection
                                RaiseEvent OnConnect(arStatusError)
                            End If
                            '
                        Case Else
                            '500, 501, 502, 421, 550
                            'Something wrong with the server - it doesn't send
                            'the current directory path. Raise the completion
                            'event and close the socket.
                            RaiseFtpError intReplyCode
                            Call CloseControlConnection
                        '
                    End Select
                    '
            End Select
            '
        Case istCreateDirectory
            '
            'The CreateDirectory asynchronous method is in progress.
            'The MKD FTP command has been sent by the CreateDirectory method.
            '
            Select Case intReplyCode
                '
                Case 257
                    'The new directory was created successfully
                    'End the method and raise the completion event.
                    Call EndAsyncMethod
                    RaiseEvent OnCreateDirectory(arStatusOk)
                Case Else
                    '500, 501, 502, 421, 530, 550
                    'The server can't create the new directory.
                    RaiseFtpError intReplyCode
                '
            End Select
            '
        Case istDeleteFile
            '
            'The DeleteFie asynchronous method is in progress.
            'The DELE FTP command has been sent by the DeleteFile method.
            '
            Select Case intReplyCode
                '
                Case 250
                    'The file was deleted successfully.
                    'End the method and raise the completion event.
                    Call EndAsyncMethod
                    RaiseEvent OnDeleteFile(arStatusOk)
                Case Else
                    '450, 550, 500, 501, 502, 421, 530
                    'The server can't delete the file.
                    RaiseFtpError intReplyCode
                '
            End Select
            '
        Case istDownloadFile
            '
            'The DownloadFile asynchronous method is in progress.
            '
            Select Case m_strLastSentCommand
                '
                'The TYPE FTP command has been sent from the DownloadFile method.
                '
                Case "TYPE"
                    '
                    Select Case intReplyCode
                        '
                        Case 200
                            '
                            'If the server response on the TYPE FTP command
                            'is OK, we can go on sending the PORT or PASV
                            'FTP command in order to establish the data connection
                            '
                            If m_bPassiveMode Then
                                Call SendFtpCommand("PASV", "")
                            Else
                                Call SendPORTCommand
                            End If
                            '
                        Case Else   '500, 501, 502, 421, 530
                            Call RaiseFtpError(intReplyCode)
                        '
                    End Select
                    
                    '
                Case "REST"
                    '
                    'The REST FTP command is sent by the OpenLocalFiel subroutine.
                    '
                    Select Case intReplyCode
                        Case 350
                            '
                            'The server is able to restart the data transfer
                            'at the specified position - send the RETR command
                            'in order to start the data transfer.
                            Call SendFtpCommand("RETR", m_strRemoteFileName)
                            '
                        Case Else   '500, 501, 502, 421, 530
                            Call RaiseFtpError(intReplyCode)
                    End Select
                    '
                Case "PORT"
                    '
                    'The PORT FTP command has been sent by the SendPORTCommand subroutine.
                    '
                    Select Case intReplyCode
                        '
                        Case 200
                            '
                            'We have received a positive response form the server
                            'on the PORT command. All is ready to start the data
                            'transfer - open the local file before sending the RETR
                            'FTP command.
                            Call OpenLocalFile
                            '
                            'The OpenLocalFile subroutine checks if the restarting
                            'of data transfer is needed. If so, it sends the REST
                            'FTP command to the server, and thus the RETR command
                            'is not sent here. It will be sent a litle bit later,
                            'after receiving reponse on the REST command - see the
                            'Case "REST" code above.
                            '
                            If Not m_strLastSentCommand = "REST" Then
                                Call SendFtpCommand("RETR", m_strRemoteFileName)
                            End If
                            '
                        Case Else   '500, 501, 421, 530
                            '
                            'The server can't execute the PORT command.
                            'Close the data connection socket and raise
                            'the completion event.
                            m_objDataSocket.CloseSocket
                            Call RaiseFtpError(intReplyCode)
                            '
                    End Select
                    '
                Case "PASV"
                    '
                    Select Case intReplyCode
                        '
                        Case 227
                            '
                            'We have received a positive response on the PASV
                            'command that contains the remote socket address
                            'we need to estblish the data connection to.
                            'The EstablishPasvDataConnection is just for that.
                            '
                            Call EstablishPasvDataConnection(strArrivedData)
                            '
                        Case Else   '500, 501, 502, 421, 530
                            '
                            Call RaiseFtpError(intReplyCode)
                            '
                    End Select
                    '
                Case "RETR"
                    '
                    Select Case intReplyCode
                        '
                        Case 125, 150, 110
                            'The data connection is opened. Starting data transfer.
                            'Ignored.
                        Case 226, 250
                            'Transfer complete.
                            'See the m_objDataSocket_OnClose event procedure code.
                            '-----------------------------------
                            'Added: 17-OCT-2002
                            If m_blnTransferComplete Then
                                ProcessOnDownloadFileEvent
                            Else
                                m_blnTransferComplete = True
                            End If
                            '-----------------------------------
                            '
                        Case Else   '425, 426, 451, 450, 550, 500, 501, 421, 530
                            '
                            Call CloseDataConnection
                            Call CloseLocalFile
                            Call RaiseFtpError(intReplyCode)
                            '
                    End Select
                    '
            End Select
            '
        Case istEnumFiles
            '
            'The EnumFiles asynchronous method is in progress.
            '
            Select Case m_strLastSentCommand
                '
                'The TYPE FTP command has been sent by the EnumFiles method.
                '
                Case "TYPE"
                    '
                    Select Case intReplyCode
                            '
                            'If the server response on the TYPE FTP command
                            'is OK, we can go on sending the PORT or PASV
                            'FTP command in order to establish the data connection
                            '
                        Case 200
                            '
                            If m_bPassiveMode Then
                                Call SendFtpCommand("PASV", "")
                            Else
                                Call SendPORTCommand
                            End If
                            '
                        Case Else   '500, 501, 504, 421, 530
                            Call RaiseFtpError(intReplyCode)
                        '
                    End Select
                    '
                Case "PORT"
                    '
                    'The PORT FTP command has been sent by the SendPORTCommand subroutine.
                    '
                    Select Case intReplyCode
                        '
                        Case 200
                            '
                            'We have received a positive response form the server
                            'on the PORT command. All is ready to start the data
                            'transfer - send the LIST FTP command in order to force
                            'the server to start sending the current direcoty listing
                            'via the data connection.
                            '
                            Call SendFtpCommand("LIST", "")
                            '
                        Case Else   '500, 501, 421, 530
                            Call RaiseFtpError(intReplyCode)
                    End Select
                    '
                Case "PASV"
                    '
                    Select Case intReplyCode
                        '
                        Case 227
                            '
                            'We have received a positive response on the PASV
                            'command that contains the remote socket address
                            'we need to estblish the data connection to.
                            'The EstablishPasvDataConnection is just for that.
                            '
                            Call EstablishPasvDataConnection(strArrivedData)
                            '
                        Case Else   '500, 501, 502, 421, 530
                            Call RaiseFtpError(intReplyCode)
                    End Select
                    '
                Case "LIST"
                    '
                    Select Case intReplyCode
                        Case 125, 150
                            'The data connection is opened. Starting data transfer.
                            'Ignored.
                        Case 226, 250
                            'Transfer complete.
                            'See also the m_objDataSocket_OnClose event procedure code.
                            '
                            '---------------------------------
                            'Added: 17-OCT-2002
                            If m_blnTransferComplete Then
                                Call ProcessOnEnumFilesEvent
                            Else
                                m_blnTransferComplete = True
                            End If
                            '---------------------------------
                            '
                        Case Else   '425, 426, 451, 450, 500, 501, 421, 530
                            Call CloseDataConnection
                            Call RaiseFtpError(intReplyCode)
                    End Select
                    '
            End Select
            '
        Case istGetCurrentDirectory
            '
            'The GetCurrentDirectory asynchronous method is in progress.
            'The PWD FTP command has been sent by the GetCurrentDirectory method.
            '
            Select Case intReplyCode
                '
                Case 257
                    '
                    Call EndAsyncMethod
                    '
                    'The server respose should contain the directory path
                    'The GetDirFromResponse function retrieves the current
                    'directory path and initializes the CurrentDirectory property
                    '
                    If GetDirFromResponse Then
                        'The directory name was retrieved successfully
                        'Raise the completion event.
                        RaiseEvent OnGetCurrentDirectory(arStatusOk)
                    Else
                        'Unexpected server response. GetDirFromResponse
                        'can't parse the respose string in oreder to
                        'retrieve the directory path.
                        'Raise the error and close the socket.
                        Err.Number = FTP_CLIENT_BASE_ERROR + 1
                        Err.Description = "Cannot retrieve the current directory path. Unexpected server response."
                        RaiseEvent OnGetCurrentDirectory(arStatusError)
                    End If
                    '
                Case Else '500, 501, 502, 421, 550
                    RaiseFtpError intReplyCode
                '
            End Select
            '
        Case istRemoveDirectory
            '
            'The RemoveDirectory asynchronous method is in progress.
            'The RMD FTP command has been sent by the RemoveDirectory method.
            '
            Select Case intReplyCode
                Case 250
                    Call EndAsyncMethod
                    RaiseEvent OnRemoveDirectory(arStatusOk)
                Case Else   '500, 501, 502, 421, 530, 550
                    RaiseFtpError intReplyCode
            End Select
            '
        Case istRenameFile
            '
            'The RenameFile asynchronous method is in progress.
            '
            Select Case m_strLastSentCommand
                '
                Case "RNFR"
                    '
                    'The RNFR FTP command has been sent by the RenameFile method.
                    '
                    Select Case intReplyCode
                        Case 350
                            '
                            'The RNFR FTP command specifies the file that is
                            'going to be renamed. If we have got a positive
                            'response on this command, we can send the RNTO
                            'one to define a new name for that file.
                            '
                            Call SendFtpCommand("RNTO", m_strNewFileName)
                            '
                        Case Else   '450, 550, 500, 501, 502, 421, 530
                            RaiseFtpError intReplyCode
                    End Select
                    '
                Case "RNTO"
                    '
                    Select Case intReplyCode
                        '
                        Case 250
                            '
                            'The was renamed successfully.
                            '
                            Call EndAsyncMethod
                            RaiseEvent OnRenameFile(arStatusOk)
                            '
                        Case Else   '532, 553, 500, 501, 502, 503, 421, 530
                            RaiseFtpError intReplyCode
                        '
                    End Select
                    '
            End Select
            '
        Case istSetCurrentDirectory
            '
            'The SetCurrentDirectory asynchronous method is in progress.
            'The CWD FTP command has been sent by the SetCurrentDirectory method.
            '
            Select Case intReplyCode
                Case 250
                    Call EndAsyncMethod
                    RaiseEvent OnSetCurrentDirectory(arStatusOk)
                Case Else   '500, 501, 502, 421, 530, 550, 553
                    RaiseFtpError intReplyCode
            End Select
            '
        Case istUploadFile
            '
            'The UploadFile asynchronous method is in progress.
            '
            Select Case m_strLastSentCommand
                '
                Case "TYPE"
                    '
                    'The TYPE FTP command has been sent by the UploadFile method.
                    '
                    Select Case intReplyCode
                        '
                        Case 200
                            '
                            If m_bPassiveMode Then
                                Call SendFtpCommand("PASV", "")
                            Else
                                Call SendPORTCommand
                            End If
                            '
                        Case Else   '500, 501, 504, 421, 530
                            '
                            Call RaiseFtpError(intReplyCode)
                            '
                    End Select
                    '
                Case "PORT"
                    '
                    Select Case intReplyCode
                        '
                        Case 200
                            '
                            'The PORT command completed successfully.
                            'Now we can open the local file and send the
                            'STOR command to let the server know that we
                            'are starting the file data transfer
                            '
                            Call OpenLocalFileForUpload
                            '
                            'If the upload operation needs to be restarted,
                            'the OpenLocalFileForUpload subroutine sends
                            'the REST FTP command. In this case, the STOR
                            'command will be sent after receiving response
                            'on the REST command.
                            '
                            If Not m_strLastSentCommand = "REST" Then
                                'Send the STOR FTP command
                                Call SendFtpCommand("STOR", m_strRemoteFileName)
                            End If
                            '
                        Case Else   '500, 501, 421, 530
                            '
                            m_objDataSocket.CloseSocket
                            Call RaiseFtpError(intReplyCode)
                            '
                    End Select
                    '
                Case "PASV"
                    '
                    Select Case intReplyCode
                        '
                        Case 227
                            '
                            'We have received a positive response on the PASV
                            'command that contains the remote socket address
                            'we need to estblish the data connection to.
                            'The EstablishPasvDataConnection is just for that.
                            '
                            Call EstablishPasvDataConnection(strArrivedData)
                            '
                        Case Else   '500, 501, 502, 421, 530
                            '
                            Call RaiseFtpError(intReplyCode)
                            '
                        '
                    End Select
                    '
                Case "REST"
                    '
                    Select Case intReplyCode
                        '
                        Case 350
                            '
                            Call SendFtpCommand("STOR", m_strRemoteFileName)
                            '
                        Case Else   '500, 501, 502, 421, 530
                            '
                            Call RaiseFtpError(intReplyCode)
                            '
                    End Select
                    '
                Case "STOR"
                    '
                    Select Case intReplyCode
                        Case 125, 150, 110
                            If m_bPassiveMode Then
                                '
                                'The data connection was established
                                'start sending the file data
                                m_lngUploadedBytes = m_lngStartPositioon
                                Call UploadData
                                '
                            End If
                        Case 226, 250
                            '
                            'To let the server know that we have finished
                            'the file data transfer, we have to close the
                            'data connection socket. As soon as server
                            'reseives the FIN TCP segment it sends the
                            'transfer complete reply code. Now we can raise
                            'the completion event.
                            '
                            Call EndAsyncMethod
                            RaiseEvent OnUploadFile(arStatusOk)
                            '
                        Case Else   '425, 426, 451, 551, 552, 532, 450, 452, 553, 500, 501, 421, 530
                            '
                            Call EndAsyncMethod
                            Call CloseLocalFile
                            Call RaiseFtpError(intReplyCode)
                            '
                    End Select
                    '
            End Select
            '
    End Select
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    Dim varInternalState As InternalStateConstants
    varInternalState = m_varInternalState
    '
    Call CloseLocalFile
    Call CloseDataConnection
    Call EndAsyncMethod
    '
    With m_objFtpError
        .Number = Err.Number
        .Source = "CFtpClient.m_objControlSocket_OnDataArrival"
        .Description = Err.Description
    End With
    '
    Call RaiseErrorCompletionEvent(varInternalState)
    '
End Sub

Private Function GetDirFromResponse() As Boolean
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-15-2002
'Purpose     :This subroutine parses the server response in oreder to
'             get the current directory path.
'
'Description :Typical server response which contains the directory path looks like:
'
'             257 "dir/subdir" is current directory.
'
'             Such a response should be received on PWD client's command.
'
'********************************************************************************
    '
    'Variables for the first and second position of the
    'double-quotes characters in the response string.
    Dim intPosA As Integer, intPosB As Integer
    '
    On Error GoTo ERROR_HANDLER
    '
    m_strCurrentDirectory = ""
    '
    If Len(m_strControlBuffer) < 7 Then
        '
        'Unexpected server response, as the server should send
        'at least 7 characters: Len('257 "/"') = 7. In this
        'case the function returns False
        Exit Function
        '
    End If
    '
    'Get the position of the first <"> character
    intPosA = InStr(1, m_strControlBuffer, Chr$(34)) + 1
    'Get the position of the second <"> character
    intPosB = InStr(intPosA, m_strControlBuffer, Chr$(34))
    '
    If intPosA > 1 And intPosB > 0 Then
        '
        'If both characters are found, we can retrieve the path
        m_strCurrentDirectory = Mid$(m_strControlBuffer, intPosA, intPosB - intPosA)
        m_strCurrentDirectory = Replace$(m_strCurrentDirectory, "\", "/")
        '
        'The function returns True
        GetDirFromResponse = True
        '
    'Else
        '
        'Something wrong with the server response.
        'Err.Raise FTP_CLIENT_BASE_ERROR + 1, "FtpClientX.CFtpClient", _
        '"Cannot retrieve current directory path. Unexpected server response."
        '
    End If
    '
    Exit Function
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.GetDirFromResponse failed: " & .Description
    End With
    '
End Function

Private Sub SendPORTCommand()
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-15-2002
'Purpose     :This procedure gets free port number with the GetFreePort function,
'             and then sends the PORT FTP command to the server.

'Description :Example of the PORT command string: "PORT 64,101,203,52,3,67"
'
'             Each number in the argument string is a value of an octet(byte).
'             The leading four numbers are octets of an IP address (32-bits value).
'             The last two - octets of a port number (16-bits value).
'
'********************************************************************************
    '
    Dim strArgument     As String   'Argument for the PORT command.
    Dim lngLocalPort    As Long     'Port number the server will establish
                                    'the data connection to.
    '
    On Error Resume Next
    '
    m_varFtpSessionState = ftpEstablishingDataConnection
    RaiseEvent OnStateChange(ftpEstablishingDataConnection)
    '
    Do
        DoEvents
        '
        'Get the next port number
        lngLocalPort = GetFreePort
        '
        Err.Number = 0
        '
        If m_objDataSocket.State <> sckClosed Then m_objDataSocket.CloseSocket
        '
        'Turn the socket into the listening state
        m_objDataSocket.LocalPort = lngLocalPort
        m_objDataSocket.Listen
        '
        'If Err.Number = 0 Then Exit Do
        '
    Loop Until Err.Number = 0
    '
    On Error GoTo ERROR_HANDLER
    '
    'Build the argument string
    strArgument = Replace$(m_objControlSocket.LocalIP, ".", ",")
    strArgument = strArgument & "," & lngLocalPort \ 256 & "," & (lngLocalPort Mod 256)
    '
    'Send the PORT command to the server
    Call SendFtpCommand("PORT", strArgument)
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.SendPORTCommand failed: " & .Description
    End With
    '
End Sub

Private Sub m_objControlSocket_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-15-2002
'Purpose     :Closes the control connection socket and raises a completion event
'********************************************************************************
    '
    Dim varInternalState As InternalStateConstants
    '
    On Error GoTo ERROR_HANDLER
    '
    varInternalState = m_varInternalState
    '
    'The OnError event of CSocket class usually means that the connection
    'is not available anymore - close the socket
    Call CloseControlConnection
    '
    'Initialize properties of the internal error object
    With m_objFtpError
        .Number = Number
        .Description = Description
        .Source = "CFtpClient.m_objControlSocket_OnError"
    End With
    '
    'Raise a completion event
    Select Case varInternalState
        Case istConnect
            RaiseEvent OnConnect(arStatusError)
        Case istCreateDirectory
            RaiseEvent OnCreateDirectory(arStatusError)
        Case istDeleteFile
            RaiseEvent OnDeleteFile(arStatusError)
        Case istDownloadFile
            RaiseEvent OnDownloadFile(arStatusError)
        Case istEnumFiles
            RaiseEvent OnEnumFiles(arStatusError)
        Case istGetCurrentDirectory
            RaiseEvent OnGetCurrentDirectory(arStatusError)
        Case istRemoveDirectory
            RaiseEvent OnRemoveDirectory(arStatusError)
        Case istRenameFile
            RaiseEvent OnRenameFile(arStatusError)
        Case istSetCurrentDirectory
            RaiseEvent OnSetCurrentDirectory(arStatusError)
        Case istUploadFile
            RaiseEvent OnUploadFile(arStatusError)
    End Select
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.m_objControlSocket_OnError failed: " & .Description
    End With
    '
End Sub

Private Sub m_objDataSocket_OnClose()
    '
    'The server is going to close the data connection.
    'That means that the data transfer from the server
    'to the client is complete.
    '
    On Error GoTo ERROR_HANDLER
    '
    Select Case m_varInternalState
        '
        Case istDownloadFile
            '
            'The file downloading is complete.
            '
            '------------------------------------
            'Added: 17-OCT-2002
            If m_blnTransferComplete Then
                ProcessOnDownloadFileEvent
            Else
                m_blnTransferComplete = True
            End If
            '------------------------------------
        Case istEnumFiles
            '
            '------------------------------------
            'Added: 17-OCT-2002
            If m_blnTransferComplete Then
                Call ProcessOnEnumFilesEvent
            Else
                m_blnTransferComplete = True
            End If
            '-----------------------------------
            '
    End Select
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.m_objDataSocket_OnClose failed: " & .Description
    End With
    '
End Sub

Private Sub m_objDataSocket_OnConnect()
    '
    On Error GoTo ERROR_HANDLER
    '
    Select Case m_varInternalState
        '
        Case istEnumFiles
            '
            Call SendFtpCommand("LIST", "")
            '
        Case istDownloadFile
            '
            Call OpenLocalFile
            '
            If Not m_strLastSentCommand = "REST" Then
                Call SendFtpCommand("RETR", m_strRemoteFileName)
            End If
            '
        Case istUploadFile
            '
            Call OpenLocalFileForUpload
            '
            If Not m_strLastSentCommand = "REST" Then
                Call SendFtpCommand("STOR", m_strRemoteFileName)
            End If
            '
'            m_lngUploadedBytes = m_lngStartPositioon
'            m_lngBytesToUpload = FileLen(m_strLocalFileName)
'            '
'            m_intLocalFile = FreeFile
'            Open m_strLocalFileName For Binary As #m_intLocalFile
'            '
'            Call SendFtpCommand("STOR", m_strRemoteFileName)
            '
    End Select
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    Dim varInternalState As InternalStateConstants
    varInternalState = m_varInternalState
    '
    Call CloseLocalFile
    Call CloseDataConnection
    Call EndAsyncMethod
    '
    With m_objFtpError
        .Number = Err.Number
        .Source = "CFtpClient.m_objDataSocket_OnConnect"
        .Description = Err.Description
    End With
    '
    Call RaiseErrorCompletionEvent(varInternalState)
    '
End Sub

Private Sub m_objDataSocket_OnConnectionRequest(ByVal requestID As Long)
    '
    On Error GoTo ERROR_HANDLER
    '
    With m_objDataSocket
        '
        If .State <> sckClosed Then .CloseSocket
        '
        .Accept requestID
        '
        If m_varInternalState = istUploadFile Then
            '
            If Not .State = sckConnected Then
                Do Until .State = sckConnected Or (Not .State = sckConnecting)
                    DoEvents
                Loop
            End If
            '
            m_lngUploadedBytes = m_lngStartPositioon
            Call UploadData
            '
        End If
        '
    End With
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    Dim varInternalState As InternalStateConstants
    varInternalState = m_varInternalState
    '
    Call CloseLocalFile
    Call CloseDataConnection
    Call EndAsyncMethod
    '
    With m_objFtpError
        .Number = Err.Number
        .Source = "CFtpClient.m_objDataSocket_OnConnectionRequest"
        .Description = Err.Description
    End With
    '
    Call RaiseErrorCompletionEvent(varInternalState)
    '
End Sub

Private Sub m_objDataSocket_OnDataArrival(ByVal bytesTotal As Long)
    '
    'The data received through the data connection
    'can be either the current directory's listing or
    'the downloaded file content.
    '
    Dim strData As String
    Dim arrData() As Byte
    '
    On Error GoTo ERROR_HANDLER
    '
    If Not m_varInternalState = istFreeState Then
        m_objTimeOut.Reset
    End If
    '
    'If there are no data in the Winsock buffer we have nothing to do here
    '
    If bytesTotal > 0 Then
        '
        'Retrieve arrived data from the Winsock buffer into
        'the arrData byte array.
        m_objDataSocket.GetData arrData(), vbByte + vbArray
        '
        Select Case m_varInternalState
            '
            Case istDownloadFile
                '
                'If the received data is the file content,
                'write it in the local file that is already
                'opened from the m_objControlSocket_OnDataArrival
                'event handler procedure.
                '
                Put m_intLocalFile, , arrData()
                '
                'Raise the OnDataTransferProgress event
                m_lngDownloadedBytes = m_lngDownloadedBytes + bytesTotal
                RaiseEvent OnDataTransferProgress(m_lngDownloadedBytes)
                '
            Case istEnumFiles
                '
                'A directory/file listing is arriving.
                'Just collect this data in the m_strDataSocketBuffer
                'variable. That data will be parsed later, as soon
                'as the data connection will be closed.
                '
                'Convert the byte array to the VB String data type
                strData = StrConv(arrData(), vbUnicode)
                '
                'Append the data to the buffer
                m_strDataSocketBuffer = m_strDataSocketBuffer & strData
                '
                'Raise the OnDataTransferProgress event
                RaiseEvent OnDataTransferProgress(Len(m_strDataSocketBuffer))
                '
        End Select
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    Dim varInternalState As InternalStateConstants
    varInternalState = m_varInternalState
    '
    Call CloseLocalFile
    Call CloseDataConnection
    Call EndAsyncMethod
    '
    With m_objFtpError
        .Number = Err.Number
        .Source = "CFtpClient.m_objDataSocket_OnDataArrival"
        .Description = Err.Description
    End With
    '
    Call RaiseErrorCompletionEvent(varInternalState)
    '
End Sub

Private Sub m_objDataSocket_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '
    Dim varIntState As InternalStateConstants
    '
    On Error GoTo ERROR_HANDLER
    '
    varIntState = m_varInternalState
    Call EndAsyncMethod
    '
    With m_objFtpError
        .Number = Number
        .Description = Description
        .Source = Source
    End With
    '
    m_objDataSocket.CloseSocket
    '
    Select Case varIntState
        '
        Case istDownloadFile
            '
            Call CloseLocalFile
            RaiseEvent OnDownloadFile(arStatusError)
            '
        Case istEnumFiles
            '
            m_strDataSocketBuffer = ""
            RaiseEvent OnEnumFiles(arStatusError)
            '
        Case istUploadFile
            '
            Call CloseLocalFile
            RaiseEvent OnUploadFile(arStatusError)
        '
    End Select
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.m_objDataSocket_OnError failed: " & .Description
    End With
    '
End Sub

Private Sub m_objDataSocket_OnSendComplete()
    '
    On Error GoTo ERROR_HANDLER
    '
    m_objTimeOut.Reset
    RaiseEvent OnDataTransferProgress(m_lngUploadedBytes)
    '
    'Debug.Print "m_objDataSocket_OnSendComplete " & m_lngUploadedBytes
    Call UploadData
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    Dim varInternalState As InternalStateConstants
    varInternalState = m_varInternalState
    '
    Call CloseLocalFile
    Call CloseDataConnection
    Call EndAsyncMethod
    '
    With m_objFtpError
        .Number = Err.Number
        .Source = "CFtpClient.m_objDataSocket_OnSendComplete"
        .Description = Err.Description
    End With
    '
    Call RaiseErrorCompletionEvent(varInternalState)
    '
End Sub

Private Sub m_objTimeOut_TimeOut()
    '
    'Debug.Print "TIMEOUT"
    '
    Select Case m_varInternalState
        Case istFreeState    'default
        Case istConnect
            RaiseEvent OnConnect(arStatusTimeOut)
        Case istGetCurrentDirectory
            RaiseEvent OnGetCurrentDirectory(arStatusTimeOut)
        Case istSetCurrentDirectory
            RaiseEvent OnSetCurrentDirectory(arStatusTimeOut)
        Case istCreateDirectory
            RaiseEvent OnCreateDirectory(arStatusTimeOut)
        Case istRemoveDirectory
            RaiseEvent OnRemoveDirectory(arStatusTimeOut)
        Case istRenameFile
            RaiseEvent OnRenameFile(arStatusTimeOut)
        Case istDeleteFile
            RaiseEvent OnDeleteFile(arStatusTimeOut)
        Case istDownloadFile
            RaiseEvent OnDownloadFile(arStatusTimeOut)
        Case istUploadFile
            RaiseEvent OnUploadFile(arStatusTimeOut)
        Case istEnumFiles
            RaiseEvent OnEnumFiles(arStatusTimeOut)
        Case istQuitSession
            RaiseEvent OnQuitSession(arStatusTimeOut)
    End Select
    '
End Sub

Private Sub BeginAsyncMethod(ByVal State As InternalStateConstants)
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-15-2002
'Purpose     :Resets properties of the internal error object - m_objFtpError,
'             stores the internal state value in the m_varInternalState
'             variable, and starts the timer.
'Description :This procedure is called at the beginning of every
'             asynchronous method of the class.
'********************************************************************************
    '
    With m_objFtpError
        .Number = 0
        .Description = ""
        .Source = ""
    End With
    '
    m_varInternalState = State
    m_objTimeOut.StartTimer
    '
End Sub

Private Sub EndAsyncMethod()
'********************************************************************************
'Author      : Oleg Gdalevich
'Date/Time   : 08-15-2002
'Updated     : 17-OCT-2002
'Purpose     : Stops the timer and change the state variables values.
'Description : This procedure is called at the end of every
'              asynchronous method of the class.
'********************************************************************************
    '
    m_objTimeOut.StopTimer
    '
    If m_varInternalState = istConnect Then
        m_varFtpSessionState = ftpClosed
    ElseIf Not m_varFtpSessionState = ftpClosed Then
        m_varFtpSessionState = ftpFreeState
    End If
    '
    m_varInternalState = istFreeState
    '
    RaiseEvent OnStateChange(m_varFtpSessionState)
    '
End Sub

Private Function FileExists(strFileName As String) As Boolean
    '
    On Error GoTo ERROR_HANDLER
    '
    FileExists = (GetAttr(strFileName) And vbDirectory) = 0
    '
ERROR_HANDLER:
    '
End Function

Public Sub CloseControlConnection()
    '
    On Error GoTo ERROR_HANDLER
    '
    Call EndAsyncMethod
    '
    If Not m_objControlSocket.State = sckClosed Then
        m_objControlSocket.CloseSocket
    End If
    '
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.CloseControlConnection failed: " & .Description
    End With
    '
End Sub

Private Sub EstablishPasvDataConnection(ByVal strData As String)
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-15-2002
'Purpose     :Parses the server response in order to retrieve the remote
'             socket address to establish the data connection to.
'             Calls the Connect method of m_objDataSocket object.

'Arguments   :The server response string received on the PASV FTP command.

'Description :Example of the string passed with strData argument

'             227 Entering Passive Mode (194,220,224,2,7,189)

'********************************************************************************
    '
    Dim strTemp         As String
    Dim lngPortNumber   As Long
    Dim strIPAddress    As String
    Dim intPos          As Integer
    Dim varNums         As Variant
    '
    On Error GoTo ERROR_HANDLER
    '
    m_varFtpSessionState = ftpEstablishingDataConnection
    RaiseEvent OnStateChange(ftpEstablishingDataConnection)
    '
    'Find the first "(" character in the strData string
    intPos = InStr(1, strData, "(") + 1
    '
    'Extract numbers string into the strTemp variable
    strTemp = Mid$(strData, intPos, InStr(1, strData, ")") - intPos)
    '
    'Put these numbers into the varNums array
    varNums = Split(strTemp, ",")
    '
    'Build the IP address
    strIPAddress = varNums(0) & "." & varNums(1) & "." & varNums(2) & "." & varNums(3)
    '
    'Build the port number value
    lngPortNumber = CLng(varNums(4) * 256 + varNums(5))
    '
    'Close the data connection socket
    m_objDataSocket.CloseSocket
    '
    'Force the winsock to select a free port in
    'order to avoid the "Address in use" error
    m_objDataSocket.LocalPort = 0
    '
    'Connect the data connection socket to the address
    'specified by the FTP server in response on PASV command
    m_objDataSocket.Connect strIPAddress, lngPortNumber
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.EstablishPasvDataConnection failed: " & .Description
    End With
    '
End Sub

Private Function GetFileListX(strListing As String) As CFtpFiles
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-15-2002
'Purpose     :Parses the directory content listing received from the FTP server.
'Returns     :If the listing was parsed successfully - an instance of the
'             CFtpFiles collection class. Otherwise - Nothing.
'Arguments   :strListing - The current directory listing data.
'********************************************************************************
    '
    Dim varFiles            As Variant
    Dim varFile             As Variant
    Dim strLine             As String
    Dim objFtpFile          As CFtpFile
    Dim objFtpFiles         As New CFtpFiles
    Dim i                   As Integer
    Dim intDatePosition     As Integer
    Dim intLinkPosition     As Integer
    Dim intFileNamePosition As Integer
    Dim intFileSizePosition As Integer
    Dim varLineParts        As Variant
    '
    On Error GoTo ERROR_HANDLER
    '
    Set GetFileListX = Nothing
    '
    'If the listing begins with the phrase like "total bla bla...", remove that line.
    If LCase(Left$(strListing, 5)) = "total" Then
        strListing = Mid$(strListing, InStr(1, strListing, vbLf) + 1)
    End If
    '
    'Retrieve the first line of the listing
    'strLine = Left$(strListing, InStr(1, strListing, vbLf))

    'Put each line from the listing into an item of the varFiles array
    varFiles = Split(strListing, vbLf)
    '
    'Read the listing line by line
    For Each varFile In varFiles
        '
        'Convert the Variant data type into the String one
        strLine = CStr(varFile)
        '
        If Len(strLine) = 0 Then Exit For
        '
        'Retrieve the start position of the file/directory date string
        intDatePosition = GetDateStartPosition(strLine)
        'Get the file name start position: intDatePosition + Len("Mar 21  1998") + 1
        intFileNamePosition = intDatePosition + 13
        '
        'Remove CR symbols from the line
        strLine = Replace$(strLine, vbCr, "")
        '
        'Create a new instance of the CFtpFile class
        Set objFtpFile = New CFtpFile
        '
        With objFtpFile
            '
            If intDatePosition > 10 Then
                '
                'Here is the "UNIX FTP listing format"
                '
                'The IsDirectory property
                .IsDirectory = (Left$(strLine, 1) = "d")
                '
                'Get the last modified date
                .LastWriteTime = GetDateX(Mid$(strLine, intDatePosition, 12))
                '
                'Get the file name
                If Left$(strLine, 1) = "l" Then
                    intLinkPosition = InStr(intFileNamePosition, strLine, " -> ")
                    If intLinkPosition > 0 Then
                        .FileName = Mid$(strLine, intFileNamePosition, intLinkPosition - intFileNamePosition)
                        .FilePath = Mid$(strLine, intLinkPosition + 4)
                    Else
                        .FileName = Mid$(strLine, intFileNamePosition)
                    End If
                Else
                    .FileName = Mid$(strLine, intFileNamePosition)
                End If
                '
                'Get the file size
                intFileSizePosition = InStrRev(strLine, " ", intDatePosition - 2) + 1
                .FileSize = CLng(Mid$(strLine, intFileSizePosition, intDatePosition - intFileSizePosition - 1))
                '
                'Retrieve other file properties
                '
                'Cut the line string
                strLine = Left$(strLine, intFileSizePosition - 1)
                '
                'Replace multiple whitespaces with a single one
                For i = 15 To 2 Step -1
                    strLine = Replace$(strLine, Space(i), " ")
                Next
                '
                'Put each part of the line into an item of the varLineParts array
                varLineParts = Split(strLine, " ")
                '
                'Assign properties values
                .Permissions = CStr(varLineParts(0))
                .Owner = CStr(varLineParts(2))
                If UBound(varLineParts) = 4 Then
                    .Group = CStr(varLineParts(3))
                End If
                '
            Else
                '
                'Here is the "DOS FTP listing format"
                '
                If Len(strLine) > 0 Then
                    .FileName = Mid$(strLine, 40)
                    strLine = Left$(strLine, 38)
                    .LastWriteTime = CDate(RTrim$(Replace$(Left$(strLine, 17), "  ", " ")))
                    '
                    If InStr(1, strLine, "<DIR>") > 0 Then
                        .IsDirectory = True
                    Else
                        .FileSize = CLng(Trim$(Mid$(strLine, 18)))
                    End If
                End If
                '
            End If
            '
            If Len(strLine) > 0 Then
                objFtpFiles.Add objFtpFile, objFtpFile.FileName
            End If
            '
        End With
        '
        Set objFtpFile = Nothing
        '
    Next
    '
    Set GetFileListX = objFtpFiles
    '
ERROR_HANDLER:
    '
    Set objFtpFile = Nothing
    Set objFtpFiles = Nothing
    '
End Function

Private Sub UploadData()
'********************************************************************************
'Author      :Oleg Gdalevich
'Date/Time   :08-16-2002
'Purpose     :Reads the next chunk of the file data and sends it to the server.
'Description :If any  error occurs during execution of this procedure,
'             the UploadFile method will be stoped, and the completion
'             event OnFileUpload will be raised.
'********************************************************************************
    '
    'Dim arrDataToSend() As Byte
    Dim lngBytesToSend  As Long
    '
    On Error GoTo ERROR_HANDLER
    '
    'This procedure is called from the m_objDataSocket_OnSendComplete
    'event handler procedure. The following line is for the case when
    'that event occurres after the call of the CancelAsyncMethod.
    If Not m_varInternalState = istUploadFile Then Exit Sub
    '
    'If we have any data to send...
    If m_lngUploadedBytes < m_lngBytesToUpload Then
        '
        'Move the file pointer at the beginning of the next chunk
        Seek #m_intLocalFile, m_lngUploadedBytes + 1
        '
        'Calculate the remainder
        lngBytesToSend = m_lngBytesToUpload - m_lngUploadedBytes
        '
        'Prepare the buffer to read the file data to
        If lngBytesToSend >= m_intSendBufferLenght Then
            
            m_lngUploadedBytes = m_lngUploadedBytes + m_intSendBufferLenght
        Else
            ReDim arrDataToSend(lngBytesToSend - 1)
            m_lngUploadedBytes = m_lngUploadedBytes + lngBytesToSend
        End If
        '
        'Read the chunk of data from the file
        Get #m_intLocalFile, , arrDataToSend()
        '
        'Send the data chunk to the server
        m_objDataSocket.SendData arrDataToSend()
        'Debug.Print "UploadData: m_lngUploadedBytes = " & UBound(arrDataToSend())
        '
    Else
        '
        'Sending file operation is complete
        'We just have to close the file and data connection.
        Call CloseLocalFile
        m_objDataSocket.CloseSocket
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.UploadData failed: " & .Description
    End With
    '
End Sub

Private Function GetResponseCode(strResponse As String) As Integer
'********************************************************************************
'Author    :Oleg Gdalevich
'Date/Time :03.15.01
'Purpose   :Parses FTP server response to retrieve a response code
'Returns   :Integer value that contains the response code; or 0, if
'           there is no response code in the server reply
'Arguments :String that contains the FTP server response
'********************************************************************************
    '
    Dim varResponseLines    As Variant
    Dim strLine             As String
    Dim varLine             As Variant
    Dim strWelcomeMessage   As String
    '
    On Error GoTo ERROR_HANDLER
    '
TRY_AGAIN:
    '
    'The FTP server response code is always a 3-digit value
    'The module level constant RESPONSE_CODE_LENGHT = 3
    If Len(strResponse) > (RESPONSE_CODE_LENGHT - 1) Then
        '
        'Put each line of the server response into
        'the varResponseLines array item
        varResponseLines = Split(strResponse, vbCrLf)
        '
        'Analyze each line of the server response
        For Each varLine In varResponseLines
            '
            'If the line begins with expression like "XXX-" (X is a numeric value)
            'then this line is just a comment
            If IsNumeric(Left(varLine, 3)) And Mid(varLine, 4, 1) = "-" Then
                '
                'Cut the comment text from the line
                strWelcomeMessage = strWelcomeMessage & Mid$(CStr(varLine), 5) & vbCrLf
                '
            'If the line begins with expression like "XXX "
            'we have got "a true" response
            ElseIf IsNumeric(Left(varLine, 3)) And Mid(varLine, 4, 1) = " " Then
                '
                GetResponseCode = CInt(Left(varLine, 3))
                '
            End If
            '
        Next
        '
        If Len(strWelcomeMessage) > 2 And GetResponseCode = 0 Then
            DoEvents
            strResponse = m_strControlBuffer
            '
            If Left$(strResponse, 4) = "221-" Then
                '
                'Just let the programm to go on
                '
                'Some servers like to send the goodbye string like this one:
                '221-You have transferred 0 bytes in 0 files.
                '
                'for example the Intel FTP server - ftp.intel.com
                'Its welcome message with the FTP software:
                '220 ftp1 FTP server (Version wu-2.6.1(1) Thu Nov 29 13:38:22 PST 2001) ready.
                '
            Else
                '
                If m_objDataSocket.State = sckConnected Then
                    GoTo TRY_AGAIN
                End If
                '
            End If
            '
        End If
        '
        If Len(strWelcomeMessage) > 2 Then
            '
            'If we have got some comments from the server
            'store it in the WelcomeMessage property of the class
            '
            strWelcomeMessage = Left$(strWelcomeMessage, Len(strWelcomeMessage) - 2)
            m_strFtpServerMessage = strWelcomeMessage
            '
        End If
        '
    End If
    '
    Exit Function
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.GetResponseCode failed: " & .Description
    End With
    '
End Function

Private Function GetDateX(ByVal strDateString As String) As Date
    '
    'The strDateString could be in the following formats:
    '
    'Jul  9 11:25
    'OR
    'Mar 21  1998
    '
    Dim strDay As String
    Dim strMonth As String
    Dim strYear As String
    Dim strTime As String
    Dim intMonth As Integer
    '
    On Error GoTo ERROR_HANDLER
    '
    strDateString = Replace$(strDateString, "  ", " ")
    '
    strMonth = LCase$(Left$(strDateString, 3))
    strDay = RTrim$(Mid$(strDateString, 5, 2))
    strYear = LTrim$(Right$(strDateString, 5))
    '
    If InStr(1, strYear, ":") > 0 Then
        strTime = strYear
        strYear = CStr(Year(Date))
    Else
        strTime = "00:00"
    End If
    '
    Select Case strMonth
        Case "jan": intMonth = 1
        Case "feb": intMonth = 2
        Case "mar": intMonth = 3
        Case "apr": intMonth = 4
        Case "may": intMonth = 5
        Case "jun": intMonth = 6
        Case "jul": intMonth = 7
        Case "aug": intMonth = 8
        Case "sep": intMonth = 9
        Case "oct": intMonth = 10
        Case "nov": intMonth = 11
        Case "dec": intMonth = 12
    End Select
    
    'GetDate = DateSerial(CInt(strYear), intMonth, CInt(vDay))
    GetDateX = CDate(intMonth & "-" & strDay & "-" & strYear & " " & strTime)
    '
    Exit Function
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.GetFreePort failed: " & .Description
    End With
    '
End Function

Private Function GetFreePort() As Long
'********************************************************************************
'Author    :Oleg Gdalevich
'Date/Time :03.15.01
'Purpose   :Just to get the next port number
'Return    :Port number
'********************************************************************************
    '
    Static lngPort As Long
    '
    On Error GoTo ERROR_HANDLER
    '
    If lngPort = 0 Then
        '
        'Get random number in a range: 20,000-40,000
        Randomize
        lngPort = CLng((20001) * Rnd + 20000)
        '
    Else
        lngPort = lngPort + 1
    End If
    '
    GetFreePort = lngPort
    '
    Exit Function
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.GetFreePort failed: " & .Description
    End With
    '
End Function

Private Sub OpenLocalFileForUpload()
    '
    'This function has been added 12-SEP-2002
    '
    On Error GoTo ERROR_HANDLER
    '
    'These values are used by the UploadData procedure
    m_lngUploadedBytes = m_lngStartPositioon
    m_lngBytesToUpload = FileLen(m_strLocalFileName)
    '
    'Open the local file
    m_intLocalFile = FreeFile
    Open m_strLocalFileName For Binary As #m_intLocalFile
    '
    If m_lngStartPositioon > 0 Then
        Call SendFtpCommand("REST", CStr(m_lngStartPositioon))
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.OpenLocalFileForUpload failed: " & .Description
    End With
    '
End Sub

Private Sub OpenLocalFile()
    '
    Dim lngRestartPosition As Long
    '
    On Error GoTo ERROR_HANDLER
    '
    'If the local file already exists, we have to go
    'by one of the two ways. If the user has requested
    'to overwrite the existing file, we just call the
    'Kill statement in order to delete the file.
    'Otherwise we have to append missing data from the
    'remote file.
    '
    If FileExists(m_strLocalFileName) Then
        '
        If m_blnOverWriteFile Then
            'Delete the local file
            Kill m_strLocalFileName
        Else
            'Get the restart position
            lngRestartPosition = FileLen(m_strLocalFileName)
        End If
        '
    End If
    '
    'Open the local file to write the downloaded data to.
    m_intLocalFile = FreeFile
    Open m_strLocalFileName For Binary As #m_intLocalFile
    '
    If lngRestartPosition > 0 Then
        '
        m_lngDownloadedBytes = lngRestartPosition
        Seek #m_intLocalFile, lngRestartPosition + 1
        Call SendFtpCommand("REST", CStr(lngRestartPosition))
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.OpenLocalFile failed: " & .Description
    End With
    '
End Sub

Private Sub CloseLocalFile()
    '
    On Error GoTo ERROR_HANDLER
    '
    If m_intLocalFile <> 0 Then
        Close #m_intLocalFile
    End If
    '
    m_intLocalFile = 0
    m_strLocalFileName = ""
    m_strRemoteFileName = ""
    m_lngDownloadedBytes = 0
    m_lngUploadedBytes = 0
    m_lngBytesToUpload = 0
    m_lngStartPositioon = 0
    m_blnOverWriteFile = False
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.CloseLocalFile failed: " & .Description
    End With
    '
End Sub

Private Function GetDateStartPosition(ByVal strListingLine As String) As Integer
    '
    'drwxr-xr-x   4 ftpuser  ftpusers       512 Jul  2  2001 kylix
    '                                          <------------>
    'drwxr-xr-x   3 ftpuser  ftpusers       512 Jul 19 11:25 optimizeit
    '                                          <------------>
    '
    Dim intStartPosition    As Integer
    Dim varMonthsArray      As Variant
    Dim strStringToSearch   As String
    Dim i                   As Integer
    '
    On Error GoTo ERROR_HANDLER
    '
    varMonthsArray = Array("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
    strListingLine = LCase$(strListingLine)
    '
    For i = 0 To 11
        '
        strStringToSearch = " " & varMonthsArray(i) & " "
        intStartPosition = InStr(1, strListingLine, strStringToSearch)
        '
        If intStartPosition > 0 Then
            If IsNumeric(Mid$(strListingLine, intStartPosition - 1, 1)) And _
               IsNumeric(Mid$(strListingLine, intStartPosition + 9, 1)) Then
               '
               GetDateStartPosition = intStartPosition + 1
               Exit For
               '
            End If
        End If
        '
    Next i
    '
    Exit Function
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.GetDateStartPosition failed. " & .Description
    End With
    '
End Function

Private Sub RaiseErrorCompletionEvent(ByVal varInternalState As InternalStateConstants)
    '
    Select Case varInternalState
        Case istFreeState    'default
        Case istConnect
            RaiseEvent OnConnect(arStatusError)
        Case istGetCurrentDirectory
            RaiseEvent OnGetCurrentDirectory(arStatusError)
        Case istSetCurrentDirectory
            RaiseEvent OnSetCurrentDirectory(arStatusError)
        Case istCreateDirectory
            RaiseEvent OnCreateDirectory(arStatusError)
        Case istRemoveDirectory
            RaiseEvent OnRemoveDirectory(arStatusError)
        Case istRenameFile
            RaiseEvent OnRenameFile(arStatusError)
        Case istDeleteFile
            RaiseEvent OnDeleteFile(arStatusError)
        Case istDownloadFile
            RaiseEvent OnDownloadFile(arStatusError)
        Case istUploadFile
            RaiseEvent OnUploadFile(arStatusError)
        Case istEnumFiles
            RaiseEvent OnEnumFiles(arStatusError)
        Case istQuitSession
            RaiseEvent OnQuitSession(arStatusError)
    End Select
    '
End Sub

Private Sub ProcessOnEnumFilesEvent()
'---------------------------------------------------------------------------------------
' Procedure : ProcessOnEnumFilesEvent
' Created   : 17-OCT-2002
' Updated   :
' Author    : Oleg Gdalevich
' Purpose   : 1) Re-creates the m_objFtpFiles collection
'             2) Calls the GetFileListX function to parse the listing
'             3) Moves the retrieved row listing data from the m_strDataSocketBuffer
'                variable into the m_strFtpFilesText one that is a local storage
'                for the CurrentDirectoryListing property
'             4) Raises the OnEnumFiles event
' Notes     : This code is moved into a separated subroutine because it could be called
'             from two places: m_objControlSocket_OnDataArrival and m_objDataSocket_OnClose
'---------------------------------------------------------------------------------------
    '
    On Error GoTo ERROR_HANDLER
    '
    m_blnTransferComplete = False
    '
    'Receiving the current directory listing is complete.
    '
    Call EndAsyncMethod
    '
    'Destroy the m_objFtpFiles  collection object
    If Not m_objFtpFiles Is Nothing Then
        Set m_objFtpFiles = Nothing
    End If
    '
    'Create a new instance of the CFtpFiles class
    Set m_objFtpFiles = New CFtpFiles
    '
    'Initialize the CurrentDirectoryListing property
    m_strFtpFilesText = m_strDataSocketBuffer
    'Initialize the CurrentDirectoryFiles property
    Set m_objFtpFiles = GetFileListX(m_strDataSocketBuffer)
    'Clear the data connection buffer
    m_strDataSocketBuffer = ""
    '
    If Not m_objFtpFiles Is Nothing Then
        '
        RaiseEvent OnEnumFiles(arStatusOk)
        '
    Else
        '
        'If m_objFtpFiles collection object is Nothing,
        'it means that the received directory listing
        'is in unknown format (for the CFtpClient) class.
        '
        With m_objFtpError
            .Number = FTP_CLIENT_BASE_ERROR + 2
            .Description = "Unknown listing format."
            .Source = "CFtpClient.m_objDataSocket_OnClose"
        End With
        '
        RaiseEvent OnEnumFiles(arStatusError)
        '
    End If
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.ProcessOnEnumFilesEvent failed. " & .Description
    End With
    '
End Sub

Private Sub ProcessOnDownloadFileEvent()
'---------------------------------------------------------------------------------------
' Procedure : ProcessOnDownloadFileEvent
' Created   : 17-OCT-2002
' Updated   :
' Author    : Oleg Gdalevich
' Purpose   : Just to call the CloseLocalFile procedure and then raise the OnDownloadFile
'             event.
' Notes     : This code is moved into a separated subroutine because it could be called
'             from two places: m_objControlSocket_OnDataArrival and m_objDataSocket_OnClose
'---------------------------------------------------------------------------------------
    '
    On Error GoTo ERROR_HANDLER
    '
    m_blnTransferComplete = False
    '
    Call EndAsyncMethod
    Call CloseLocalFile
    RaiseEvent OnDownloadFile(arStatusOk)
    '
    Exit Sub
    '
ERROR_HANDLER:
    '
    With Err
        .Raise .Number, .Source, "CFtpClient.ProcessOnDownloadFileEvent failed. " & .Description
    End With
    '
End Sub
