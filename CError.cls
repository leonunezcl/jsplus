VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_StrError As String
Private m_Linea As Integer

Public Property Get Linea() As Integer
    Linea = m_Linea
End Property

Public Property Let Linea(ByVal pLinea As Integer)
    m_Linea = pLinea
End Property

Public Property Get StrError() As String
    StrError = m_StrError
End Property

Public Property Let StrError(ByVal pStrError As String)
    m_StrError = pStrError
End Property
