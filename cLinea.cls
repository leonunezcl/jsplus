VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Linea As String
Private m_Columna As Integer
Private m_Fila As Integer

Public Property Get Linea() As String
    Linea = m_Linea
End Property

Public Property Let Linea(ByVal pLinea As String)
    m_Linea = pLinea
End Property

Public Property Get Columna() As Integer
    Columna = m_Columna
End Property

Public Property Let Columna(ByVal pColumna As Integer)
    m_Columna = pColumna
End Property

Public Property Get Fila() As Integer
    Fila = m_Fila
End Property

Public Property Let Fila(ByVal pFila As Integer)
    m_Fila = pFila
End Property
