VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSearchFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_File As cSearchFile
Private m_Files As New Collection
'Private m_Count As Long

Public Sub AddLine(ByVal Filename As String, ByVal linea As String, _
                   ByVal Fila As Integer, ByVal Columna As Integer)

    Dim k As Integer
    Dim pLinea As New cLinea
    
    For k = 1 To m_Files.Count
        Set m_File = New cSearchFile
        Set m_File = m_Files.Item(k)
        If m_File.Filename = Filename Then
            pLinea.linea = linea
            pLinea.Fila = Fila
            pLinea.Columna = Columna
            m_File.AddLine pLinea
            Exit For
        End If
    Next k
    
End Sub


Public Property Get Files() As Collection
    Set Files = m_Files
End Property
Public Property Set Files(ByVal pFiles As Collection)
    Set m_Files = pFiles
End Property





Public Sub Clear()
    
    Dim k As Integer
    
    For k = m_Files.Count To 1 Step -1
        m_Files.Remove k
    Next k
    
End Sub
Public Property Get Count() As Long
    Count = m_Files.Count
End Property


Public Sub Add(pFile As cSearchFile)
        
    'agregar la funcion a la coleccion
    m_Files.Add pFile
                
End Sub

Public Function Exists(ByVal Filename As String) As Boolean

    Dim k As Integer
    
    For k = 1 To m_Files.Count
        Set m_File = New cSearchFile
        Set m_File = m_Files.Item(k)
        If m_File.Filename = Filename Then
            Exists = True
            Exit Function
        End If
    Next k
    
End Function
