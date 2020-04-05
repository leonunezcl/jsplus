VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFuncion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Nombre As String
Private m_Parametros As String
Private m_Tag As String
Private m_FullName As String
Public Property Get Nombre() As String
    Nombre = m_Nombre
End Property

Public Property Let Nombre(ByVal pNombre As String)
    m_Nombre = pNombre
End Property

Public Property Get Parametros() As String
    Parametros = m_Parametros
End Property

Public Property Let Parametros(ByVal pParametros As String)
    m_Parametros = pParametros
End Property



Public Property Get Tag() As String
    Tag = m_Tag
End Property

Public Property Let Tag(ByVal pTag As String)
    m_Tag = pTag
End Property

Public Property Get FullName() As String
    FullName = m_FullName
End Property

Public Property Let FullName(ByVal pFullName As String)
    m_FullName = pFullName
End Property
