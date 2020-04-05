VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_SelEvento As String
Private m_Idxobj As Integer
Private m_Texto As String

Public Property Get SelEvento() As String
    SelEvento = m_SelEvento
End Property

Public Property Let SelEvento(ByVal pSelEvento As String)
    m_SelEvento = pSelEvento
End Property

Public Property Let IdxObj(ByVal pIdxObj As Integer)
    m_Idxobj = pIdxObj
End Property

Public Property Get IdxObj() As Integer
    IdxObj = m_Idxobj
End Property


Public Property Get Texto() As String
    Texto = m_Texto
End Property

Public Property Let Texto(ByVal pTexto As String)
    m_Texto = pTexto
End Property

