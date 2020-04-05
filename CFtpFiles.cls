VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFtpFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"CFtpFile"
Attribute VB_Ext_KEY = "Member0" ,"CFtpFile"
'------------------------------------------------------------------------------------------
' Module      : CFtpFiles Class Module (CFtpFiles.cls)
' Type        : Collection class
' Updated     : 17-OCT-2002
' Version     : 1.0.0
' Author      : Oleg Gdalevich
' Purpose     : Collection of the CFtpFile objects
' Notes       : Used by the CFtpClient class
' Dependencies: CFtpFile class module
' URL         : http://www.vbip.com/protocols/ftp/vb-ftp-client-library/default.asp
'------------------------------------------------------------------------------------------
' Copyright © 2002 by Oleg Gdalevich
' Visual Basic Internet Programming website (http://www.vbip.com)
'------------------------------------------------------------------------------------------
Option Explicit

Private m_colFtpFiles As Collection

Public Function Add(oFtpFile As CFtpFile, vKey As Variant) As CFtpFile
    '
    m_colFtpFiles.Add oFtpFile
    Set Add = oFtpFile
    '
End Function

Public Property Get Item(vntIndexKey As Variant) As CFtpFile
Attribute Item.VB_UserMemId = 0
    Set Item = m_colFtpFiles(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = m_colFtpFiles.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    m_colFtpFiles.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colFtpFiles.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set m_colFtpFiles = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colFtpFiles = Nothing
End Sub
