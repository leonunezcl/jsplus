VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSearchFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_FileName As String
Private m_Lines As New Collection

Public Sub AddLine(pLinea As cLinea)
    
    m_Lines.Add pLinea
    
End Sub


Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal pFileName As String)
    m_FileName = pFileName
End Property
Public Property Get Lines() As Collection
    Set Lines = m_Lines
End Property


