VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPlugin"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Autor As String
Private m_Caption As String
Private m_Description As String
Private m_Version As String
Private m_ClassId As String
Private m_IconMenu As StdPicture
Private m_MenuId As String
Private m_FileName As String
Public Property Get Autor() As String
    Autor = m_Autor
End Property

Public Property Let Autor(ByVal pAutor As String)
    m_Autor = pAutor
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal pCaption As String)
    m_Caption = pCaption
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(ByVal pDescription As String)
    m_Description = pDescription
End Property

Public Property Get Version() As String
    Version = m_Version
End Property

Public Property Let Version(ByVal pVersion As String)
    m_Version = pVersion
End Property

Public Property Get IconMenu() As StdPicture
    Set IconMenu = m_IconMenu
End Property

Public Property Set IconMenu(ByVal pIconMenu As StdPicture)
    Set m_IconMenu = pIconMenu
End Property

Public Property Get ClassId() As String
    ClassId = m_ClassId
End Property

Public Property Let ClassId(ByVal pClassId As String)
    m_ClassId = pClassId
End Property


Public Property Get MenuId() As String
    MenuId = m_MenuId
End Property

Public Property Let MenuId(ByVal pMenuId As String)
    m_MenuId = pMenuId
End Property

Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal pFileName As String)
    m_FileName = pFileName
End Property
