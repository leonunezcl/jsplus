VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Eventos As New Collection
Private m_Evento As New cEvento


Public Sub Add()
    m_Eventos.Add m_Evento
End Sub

Public Sub Attach(cboEvents As ComboBox)

    cboEvents.Clear
    Dim k As Integer
    'With CEventos
        For k = 1 To m_Eventos.Count
            Set m_Evento = New cEvento
            Set m_Evento = m_Eventos(k)
            cboEvents.AddItem m_Evento.SelEvento & "=" & Chr$(34) & m_Evento.Texto & Chr$(34)
            Set m_Evento = Nothing
        Next k
    'End With
        
End Sub

Public Sub Clear()

    Dim k As Integer
    
    For k = m_Eventos.Count To 1 Step -1
        Set m_Evento = New cEvento
        Set m_Evento = m_Eventos.Item(k)
        Set m_Evento = Nothing
        m_Eventos.Remove k
    Next k
    
End Sub

Public Property Get Eventos() As Collection
    Set Eventos = m_Eventos
End Property

Public Property Set Eventos(vNewValue As Collection)
    Set m_Eventos = vNewValue
End Property

Public Property Get Evento() As cEvento
    Set Evento = m_Evento
End Property

Public Property Set Evento(vNewValue As cEvento)
    Set m_Evento = vNewValue
End Property
Public Function Output() As String

    Dim src As New cStringBuilder
    Dim k As Integer
    
    src.Append " "
    
    For k = 1 To m_Eventos.Count
        Set m_Evento = New cEvento
        Set m_Evento = m_Eventos.Item(k)
        src.Append " " & m_Evento.SelEvento & "=" & Chr$(34) & m_Evento.Texto & Chr$(34)
        Set m_Evento = Nothing
    Next k
    
    Output = src.ToString
    
    Set src = Nothing
    
End Function

Public Sub Remove(ByVal strEvento As String)

    Dim k As Integer
    
    strEvento = Explode(strEvento, 1, "=")
    For k = 1 To m_Eventos.Count
        Set m_Evento = New cEvento
        Set m_Evento = m_Eventos.Item(k)
        If m_Evento.SelEvento = strEvento Then
            m_Eventos.Remove k
            Set m_Evento = Nothing
            Exit For
        End If
        Set m_Evento = Nothing
    Next k
    
End Sub


