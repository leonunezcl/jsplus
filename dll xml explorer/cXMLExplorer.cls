VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXmlExplorer"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_FileName As String
Private m_LangPath As String
Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal pFileName As String)
    m_FileName = pFileName
End Property
Public Sub StartExplorer()

    ListaLangs.Load m_LangPath
    
    frmViewTree.FileName = m_FileName
    frmViewTree.Show vbModal
End Sub



Public Property Get LangPath() As String
    LangPath = m_LangPath
End Property

Public Property Let LangPath(ByVal pLangPath As String)
    m_LangPath = pLangPath
End Property

