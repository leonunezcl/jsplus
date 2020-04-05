VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGetFiles"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public OutputString As New cStringBuilder
Private m_ReturnString As Boolean
Private m_Autor As String
Private m_Caption As String
Private m_Description As String
Private m_Version As String
Private m_ClassId As String
Private m_ErrNumber As String
Private m_ErrMessage As String
Public Function Start() As Boolean

     On Error GoTo StartError

     Dim frm As New frmPlugin

     frm.Show vbModal

     If Len(glbOutputString.ToString) > 0 Then
         OutputString.Append glbOutputString.ToString
         m_ReturnString = True
     End If

     Set frm = Nothing

     Start = True

     Exit Function
StartError:
m_ErrNumber = Err.Number
m_ErrMessage = Err.Description
End Function

Private Sub Class_Initialize()
    m_Autor = "Luis Nunez"
    m_Caption = "VBSoftware - Get files from folder"
    m_Description = "This plugin gets files from any folder"
    m_Version = App.Major & "." & App.Minor & "." & App.Revision
    m_ClassId = "cGetFiles"
End Sub
Public Property Get ClassId() As String
     ClassId = m_ClassId
End Property
Public Property Get ReturnString() As Boolean
    ReturnString = m_ReturnString
End Property
Public Property Get Autor() As String
    Autor = m_Autor
End Property
Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Get Description() As String
    Description = m_Description
End Property
Public Property Get Version() As String
    Version = m_Version
End Property
Public Property Get ErrNumber() As Long
    ErrNumber = m_ErrNumber
End Property
Public Property Let ErrNumber(ByVal pErrNumber As Long)
    m_ErrNumber = pErrNumber
End Property
Public Property Get ErrMessage() As String
    ErrMessage = m_ErrMessage
End Property
Public Property Let ErrMessage(ByVal pErrMessage As String)
    m_ErrMessage = pErrMessage
End Property

