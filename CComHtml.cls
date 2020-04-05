VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Class As String
Private m_Name As String
Private m_Id As String
Private m_Style As String
Private m_Title As String

Public Property Get Class() As String
    Class = m_Class
End Property

Public Property Let Class(ByVal vNewValue As String)
    m_Class = vNewValue
End Property

Public Sub Clear()
    m_Class = vbNullString
    m_Name = vbNullString
    m_Id = vbNullString
    m_Style = vbNullString
    m_Title = vbNullString
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    m_Name = vNewValue
End Property

Public Property Get ID() As String
    ID = m_Id
End Property

Public Property Let ID(ByVal vNewValue As String)
    m_Id = vNewValue
End Property

Public Function Output() As String

    Dim src As New cStringBuilder
    
    src.Append " "
    If Len(m_Class) > 0 Then
        src.Append " class=" & Chr$(34) & m_Class & Chr$(34)
    End If
    
    If Len(m_Name) > 0 Then
        src.Append " name=" & Chr$(34) & m_Name & Chr$(34)
    End If
    
    If Len(m_Id) > 0 Then
        src.Append " id=" & Chr$(34) & m_Id & Chr$(34)
    End If
    
    If Len(m_Style) > 0 Then
        src.Append " style=" & Chr$(34) & m_Style & Chr$(34)
    End If
    
    If Len(m_Title) > 0 Then
        src.Append " title=" & Chr$(34) & m_Title & Chr$(34)
    End If
    
    Output = src.ToString
    
    Set src = Nothing
    
End Function

Public Function strClass() As String
    strClass = "class=" & Chr$(34) & m_Class & Chr$(34)
End Function

Public Function strId() As String
    strId = "id=" & Chr$(34) & m_Id & Chr$(34)
End Function

Public Function strName() As String
    strName = "name=" & Chr$(34) & m_Name & Chr$(34)
End Function

Public Function strStyle() As String
    strStyle = "style=" & Chr$(34) & m_Style & Chr$(34)
End Function

Public Property Get Style() As String
    Style = m_Style
End Property

Public Property Let Style(ByVal vNewValue As String)
    m_Style = vNewValue
End Property

Public Function strTitle() As String
    strTitle = "title=" & Chr$(34) & m_Title & Chr$(34)
End Function

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(ByVal vNewValue As String)
    m_Title = vNewValue
End Property
