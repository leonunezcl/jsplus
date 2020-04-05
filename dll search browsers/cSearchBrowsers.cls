VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSearchBrowsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_iexplorer As String
Private m_firefox As String
Private m_netscape As String
Private m_opera As String

Public Property Get Iexplorer() As String
    Iexplorer = m_explorer
End Property

Public Property Let Iexplorer(ByVal pExplorer As String)
    m_explorer = pExplorer
End Property

Public Property Get Firefox() As String
    Firefox = m_firefox
End Property

Public Property Let Firefox(ByVal pFirefox As String)
    m_firefox = pFirefox
End Property

Public Property Get Netscape() As String
    Netscape = m_netscape
End Property

Public Property Let Netscape(ByVal pNetscape As String)
    m_netscape = pNetscape
End Property

Public Property Get Opera() As String
    Opera = m_opera
End Property

Public Property Let Opera(ByVal pOpera As String)
    m_opera = pOpera
End Property
Public Function Start() As Boolean

    Dim ret As Boolean
    
    If MsgBox("Would you like JavaScript Plus! to search for your browsers automatically") = vbYes Then
    
    End If
    
    Start = ret
    
End Function


