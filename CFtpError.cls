VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFtpError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------
' Module      : CFtpError Class Module (CFtpError.cls)
' Type        : Data holder
' Updated     : 03-OCT-2002
' Version     : 1.0.0
' Author      : Oleg Gdalevich
' Purpose     : Holds FTP error information
' Notes       : This module is used by CFtpClient class module
' Dependencies: No
' URL         : http://www.vbip.com/protocols/ftp/vb-ftp-client-library/default.asp
'------------------------------------------------------------------------------------------
' Copyright © 2002 by Oleg Gdalevich
' Visual Basic Internet Programming website (http://www.vbip.com)
'------------------------------------------------------------------------------------------
Option Explicit

Private m_lngNumber      As Long
Private m_strDescription As String
Private m_strSource      As String

Public Property Get Number() As Long
    Number = m_lngNumber
End Property

Public Property Let Number(ByVal lngNewValue As Long)
    m_lngNumber = lngNewValue
End Property

Public Property Get Description() As String
    Description = m_strDescription
End Property

Public Property Let Description(ByVal strNewValue As String)
    m_strDescription = strNewValue
End Property

Public Property Get Source() As String
    Source = m_strSource
End Property

Public Property Let Source(ByVal strNewValue As String)
    m_strSource = strNewValue
End Property
