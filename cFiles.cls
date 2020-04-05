VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_File As cFile
Private m_Files As New Collection
'Private m_Count As Long
Private m_Id As Integer
Private m_Glosa As String
Private m_FileName As String




Public Sub ClearFileName(frm As Form)

   Dim k As Integer
   Dim pFile As cFile
    
   For k = 1 To m_Files.count
      Set pFile = New cFile
      Set pFile = m_Files.ITem(k)
       
      If pFile.IdDoc = CInt(frm.Tag) Then
         pFile.filename = vbNullString
         Exit For
      End If
      Set pFile = Nothing
   Next k
    
End Sub

Public Function CloseAll() As Boolean

    Dim frm As Form
    
    For Each frm In Forms
        If TypeName(frm) = "frmEdit" Then
            If Not CloseFile(frm) Then
                CloseAll = False
                Exit Function
            Else
                Unload frm
            End If
        End If
    Next
    
    CloseAll = True
    
    Set frm = Nothing
    
End Function

Public Function CloseFile(frm As Form) As Boolean

    Dim k As Integer
    'Dim ret As Boolean
    Dim pFile As cFile
    
    For k = 1 To m_Files.count
        Set pFile = New cFile
        Set pFile = m_Files.ITem(k)
        If pFile.IdDoc = CInt(frm.Tag) Then
            CloseFile = pFile.CloseFile(frm)
            Set pFile = Nothing
            Exit Function
        End If
        Set pFile = Nothing
    Next k
    
    CloseFile = True
    
End Function

Public Property Let filename(ByVal pFileName As String)
    m_FileName = pFileName
End Property

Public Property Get Files() As Collection
    Set Files = m_Files
End Property


Public Property Set Files(ByVal pFiles As Collection)
    Set m_Files = pFiles
End Property


Public Sub Add(pFile As cFile)
        
    'agregar la funcion a la coleccion
    m_Files.Add pFile
                
End Sub
Public Sub Clear()
    
    Dim k As Integer
    
    For k = m_Files.count To 1 Step -1
        m_Files.Remove k
    Next k
    
End Sub



Public Property Get count() As Long
    count = m_Files.count
End Property

Public Function GetFileName(frm As Form) As String

   Dim k As Integer
   Dim pFile As cFile
   Dim ret As String
   
   For k = 1 To m_Files.count
      Set pFile = New cFile
      Set pFile = m_Files.ITem(k)
       
      If pFile.IdDoc = CInt(frm.Tag) Then
         ret = pFile.filename
         GoTo salir
      End If
      Set pFile = Nothing
   Next k
   
salir:
   GetFileName = ret
   
End Function

Public Function GetOpenFile() As Boolean

    Dim sFileName As String
        
    m_FileName = vbNullString
    
    If Cdlg.VBGetOpenFileName(sFileName, , , , , , glosa, , , , , frmMain.hwnd) Then
        GetOpenFile = True
    Else
        GetOpenFile = False
    End If
        
    m_FileName = ObtenerNombreLargoArchivo(sFileName)
    
End Function

Public Function GetId() As Integer

    GetId = m_Id
    
    m_Id = m_Id + 1
    
End Function


Public Function GetSaveFile(ByRef TheFileName As String) As Boolean
        
   Dim strTempFileName As String
   Dim fCancel As Boolean
   
   strTempFileName = TheFileName
   
volver:
    If Cdlg.VBGetSaveFileName(TheFileName, , , glosa, , LastPath, , "", frmMain.hwnd) Then
        If InStr(TheFileName, ".") = 0 Then
            Dim ret As Integer
            
            Dim default_ext As String
        
            default_ext = LCase$(util.LeeIni(IniPath, "startup", "default_ext"))
    
            If default_ext = "" Then
                default_ext = "js"
                Call util.GrabaIni(IniPath, "startup", "default_ext", default_ext)
            End If
            
            ret = MsgBox("There is no a file extension in the file name" & vbNewLine & "Do you want to add the '" & default_ext & "' extension automatically?", vbYesNoCancel + vbQuestion + vbDefaultButton1)
            If ret = 2 Then GoTo volver
            If ret = 6 Then
                TheFileName = TheFileName & "." & default_ext
            End If
        End If
    Else
        fCancel = True
        TheFileName = strTempFileName
    End If
   
    'If Len(TheFileName) > 0 Then
    If Not fCancel Then
        GetSaveFile = True
    Else
        GetSaveFile = False
    End If
    
End Function

Public Function IsOpen() As Boolean
    
    Dim k As Integer
    Dim ret As Boolean
    Dim pFile As cFile
    
    For k = 1 To m_Files.count
        Set pFile = New cFile
        Set pFile = m_Files.ITem(k)
        If LCase$(pFile.filename) = LCase$(m_FileName) Then
            ret = True
            Exit For
        End If
        Set pFile = Nothing
    Next k
            
    IsOpen = ret
    
End Function

Public Sub Remove(ByVal IdDoc As Integer)

    Dim k As Integer
    'Dim ret As Boolean
    Dim pFile As cFile
    
    For k = 1 To m_Files.count
        Set pFile = New cFile
        Set pFile = m_Files.ITem(k)
        If pFile.IdDoc = IdDoc Then
            m_Files.Remove k
            Exit For
        End If
        Set pFile = Nothing
    Next k
    
End Sub

Public Sub Save(frm As Form, ByVal fSaveAs As Boolean, Optional ByVal ClearFileName As Boolean = False)

    Dim k As Integer
    'Dim ret As Boolean
    Dim pFile As cFile
    
    For k = 1 To m_Files.count
        Set pFile = New cFile
        Set pFile = m_Files.ITem(k)
        
        'Debug.Print pFile.Caption
        If pFile.IdDoc = CInt(frm.Tag) Then
            pFile.SaveFile frm, fSaveAs
            
            If ClearFileName Then
               pFile.filename = vbNullString
            End If
            Exit For
        End If
        Set pFile = Nothing
    Next k
    
End Sub

Public Sub SaveFtp(frm As Form, ByVal fSaveAs As Boolean, _
                 ByVal filename As String, ByVal TempFile As String, ByVal SiteName As String, _
                 ByVal url As String, ByVal User As String, ByVal pwd As String, _
                 ByVal Port As Integer, ByVal RemoteFolder As String, ByVal Passive As Integer, _
                 Optional ByVal SaveDialog As Boolean = True)

    Dim k As Integer
    'Dim ret As Boolean
    Dim pFile As cFile
    Dim IdFound As Boolean
    
    For k = 1 To m_Files.count
        Set pFile = New cFile
        Set pFile = m_Files.ITem(k)
        If pFile.IdDoc = CInt(frm.Tag) Then
            pFile.Ftp = True
            pFile.filename = filename
            pFile.TempFile = TempFile
            pFile.SiteName = SiteName
            pFile.url = url
            pFile.User = User
            pFile.pwd = pwd
            pFile.Port = Port
            pFile.Passive = Passive
            pFile.RemoteFolder = RemoteFolder
            pFile.SaveFile frm, fSaveAs, SaveDialog
            IdFound = True
            Exit For
        End If
        Set pFile = Nothing
    Next k
    
End Sub


Public Sub SaveAll(frm As Form, ByVal Path As String)
    
    Dim k As Integer
    'Dim ret As Boolean
    Dim pFile As cFile
    
    For k = 1 To m_Files.count
        Set pFile = New cFile
        Set pFile = m_Files.ITem(k)
        
        'Debug.Print pFile.Caption
        If pFile.IdDoc = CInt(frm.Tag) Then
            pFile.SaveFile2 frm, Path
            Exit For
        End If
        Set pFile = Nothing
    Next k
    
End Sub


Private Sub Class_Initialize()
    m_Id = 1
    m_Glosa = strGlosa()
End Sub



Public Property Get glosa() As String
    glosa = m_Glosa
End Property


Public Property Get filename() As String
    filename = m_FileName
End Property

