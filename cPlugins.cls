VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Plugin As cPlugin
Private m_Plugins As New Collection
'Private m_Count As Integer
Public Sub Add(Plugin As cPlugin)

    Dim k As Integer
    
    k = m_Plugins.Count + 1
    
    m_Plugins.Add Plugin, "plug" & k
    
End Sub

Public Sub Clear()
    
    Dim k As Integer
    
    For k = m_Plugins.Count To 1 Step -1
        m_Plugins.Remove k
    Next k
    
End Sub

Public Sub Remove(ByVal Caption As String)

    On Error GoTo ErrorPlugin
    
    Dim src As New cStringBuilder
    Dim Plugin As cPlugin
    Dim k As Integer
    'Dim Msg As String
    
    For k = 1 To m_Plugins.Count
        Set Plugin = New cPlugin
        Set Plugin = m_Plugins.Item(k)
        If Plugin.Caption = Caption Then
            util.BorrarArchivo Plugin.Filename
            Exit For
        End If
        Set Plugin = Nothing
    Next k
    
    Set src = Nothing
    Set Plugin = Nothing
    
    Exit Sub
    
ErrorPlugin:
    MsgBox "cPlugins.Remove " & Err & " " & Error$, vbCritical
            
End Sub

Public Sub Run(ByVal Caption As String)
    
    On Error GoTo ErrorPlugin
    
    Dim obj As Object
    Dim src As New cStringBuilder
    Dim Plugin As cPlugin
    Dim k As Integer
    Dim Msg As String
    
    For k = 1 To m_Plugins.Count
        Set Plugin = New cPlugin
        Set Plugin = m_Plugins.Item(k)
        If Plugin.Caption = Caption Then
            'ejecutar el plugin
            Set obj = CreateObject("plugin." & Plugin.ClassId)
    
            If obj.Start() Then
                If obj.returnstring Then
                    If Not frmMain.ActiveForm Is Nothing Then
                        If frmMain.ActiveForm.Name = "frmEdit" Then
                            util.Hourglass frmMain.hwnd, True
                            frmMain.ActiveForm.Insertar obj.OutputString.ToString
                            util.Hourglass frmMain.hwnd, False
                        End If
                    End If
                End If
            Else
                Msg = "Plugin Engine Error" & vbNewLine
                Msg = Msg & vbNewLine
                Msg = Msg & "Error : " & obj.errnumber & vbNewLine
                Msg = Msg & "Description :" & obj.errmessage & vbNewLine
                MsgBox Msg, vbCritical
            End If
            
            Exit For
        End If
        Set Plugin = Nothing
    Next k
    
    Set src = Nothing
    Set obj = Nothing
    Set Plugin = Nothing
    
    Exit Sub
    
ErrorPlugin:
    MsgBox "cPlugins.Run " & Err & " " & Error$, vbCritical
    
End Sub



Public Sub Load()

    On Error GoTo ErrorPlugin
    
    'Dim V As Variant
    'Dim inifile As String
    Dim ele As String
    Dim obj As Object
    Dim Plugin As cPlugin
    Dim k As Integer
    Dim s As String
    'Dim c As Integer
    
    Clear
    
    s = Dir(App.path & "\plugins\")
    k = 1
    Do Until s = ""
        If LCase$(VBA.Right$(s, 4)) = ".dll" Then
            ele = "c" & Left$(s, InStr(1, s, ".") - 1)
            On Error Resume Next
            Set obj = CreateObject("plugin." & ele)
            If Err = 0 Then
                If Len(obj.Caption) > 0 Then
                    Set Plugin = New cPlugin
                    Plugin.Autor = obj.Autor
                    Plugin.Caption = obj.Caption
                    Plugin.Description = obj.Description
                    Plugin.Version = obj.Version
                    Plugin.ClassId = obj.ClassId
                    Plugin.MenuId = "mnuPlugins(" & k & ")"
                    Plugin.Filename = util.StripPath(App.path) & "plugins\" & s
                    Add Plugin
                    k = k + 1
                    Set Plugin = Nothing
                End If
            End If
            Set obj = Nothing
            Err = 0
        End If
        s = Dir()
    Loop
        
    Exit Sub
ErrorPlugin:
    MsgBox "cPlugins.Load " & Err & " " & Error$, vbCritical
    
End Sub

Public Property Get Plugins() As Collection
    Set Plugins = m_Plugins
End Property

Public Property Set Plugins(ByVal pPlugins As Collection)
    Set m_Plugins = pPlugins
End Property

Public Property Get Count() As Integer
    Count = m_Plugins.Count
End Property

