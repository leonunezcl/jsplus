VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFtpFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'------------------------------------------------------------------------------------------
' Module      : CFtpFile Class Module (CFtpFile.cls)
' Type        : Data holder
' Updated     : 17-OCT-2002
' Version     : 1.0.0
' Author      : Oleg Gdalevich
' Purpose     : Holds FTP file(directory) information
' Notes       : This module is used by CFtpClient and CFtpFiles class modules
' Dependencies: No
' URL         : http://www.vbip.com/protocols/ftp/vb-ftp-client-library/default.asp
'------------------------------------------------------------------------------------------
' Copyright © 2002 by Oleg Gdalevich
' Visual Basic Internet Programming website (http://www.vbip.com)
'------------------------------------------------------------------------------------------
Option Explicit

Private m_strFileName       As String
Private m_dtLastWriteTime   As Date
Private m_lngFileSize       As Long
Private m_blnIsDirectory    As Boolean
Private m_strFilePath       As String
Private m_strPermissions    As String
Private m_strOwner          As String
Private m_strGroup          As String

Public Property Get FileName() As String
    FileName = m_strFileName
End Property

Public Property Let FileName(ByVal strFileName As String)
    m_strFileName = strFileName
End Property

Public Property Get LastWriteTime() As Date
    LastWriteTime = m_dtLastWriteTime
End Property

Public Property Let LastWriteTime(ByVal dtLastWriteTime As Date)
    m_dtLastWriteTime = dtLastWriteTime
End Property

Public Property Get FileSize() As Long
    FileSize = m_lngFileSize
End Property

Public Property Let FileSize(ByVal lngFileSize As Long)
    m_lngFileSize = lngFileSize
End Property

Public Property Get IsDirectory() As Boolean
    IsDirectory = m_blnIsDirectory
End Property

Public Property Let IsDirectory(ByVal blnIsDirectory As Boolean)
    m_blnIsDirectory = blnIsDirectory
End Property

Public Property Get FilePath() As String
    FilePath = m_strFilePath
End Property

Public Property Let FilePath(ByVal strFilePath As String)
    m_strFilePath = strFilePath
End Property

Public Property Get Permissions() As String
    Permissions = m_strPermissions
End Property

Public Property Let Permissions(ByVal strPermissions As String)
    m_strPermissions = strPermissions
End Property

Public Property Get Owner() As String
    Owner = m_strOwner
End Property

Public Property Let Owner(ByVal strOwner As String)
    m_strOwner = strOwner
End Property

Public Property Get Group() As String
    Group = m_strGroup
End Property

Public Property Let Group(ByVal strGroup As String)
    m_strGroup = strGroup
End Property

