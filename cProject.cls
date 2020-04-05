VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cproject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Nombre As String
Private m_FileName As String
Private m_LastPath As String
Private m_File As cFile
Private m_Files As New Collection

Public Property Set Files(ByVal pFiles As Collection)
    Set m_Files = pFiles
End Property



Public Property Get Files() As Collection
    Set Files = m_Files
End Property



Public Sub CloseProject()
    
    Dim k As Integer
    
    Dim Msg As String
        
    If m_Files.count > 0 Then
        Msg = "Save files before close this project"
        If Confirma(Msg) = vbYes Then
            Call SaveProject
        End If
    End If
    
    For k = m_Files.count To 1 Step -1
        m_Files.Remove k
    Next k
    
    m_Nombre = vbNullString
    m_FileName = vbNullString
    frmMain.picStatus.Panels(7).Text = m_Nombre
    
End Sub

Public Property Get count() As Long
    count = m_Files.count
End Property

Public Sub AddFile(pFile As cFile)
    
    Dim plFile As cFile
    Dim k As Integer
    
    With m_Files
        For k = 1 To .count
            Set plFile = New cFile
            Set plFile = .ITem(k)
            If plFile.filename = pFile.filename Then
                MsgBox "File : " & pFile.filename & " already exists.", vbCritical
                Exit Sub
            End If
            Set plFile = Nothing
        Next k
    End With
    
    
    'agregar la funcion a la coleccion
    m_Files.Add pFile
    
End Sub


Public Sub DeleteFile(ByVal filename As String)

    Dim k As Integer
    'Dim ret As Boolean
    Dim pFile As cFile
    
    For k = 1 To m_Files.count
        Set pFile = New cFile
        Set pFile = m_Files.ITem(k)
        If pFile.filename = filename Then
            m_Files.Remove k
            Exit For
        End If
        Set pFile = Nothing
    Next k
    
End Sub


Public Sub DeleteProject()

End Sub

Public Sub ManageProject()
    If Len(m_Nombre) > 0 Then
        frmProMan.Show vbModal
    Else
        MsgBox "There is no a current project active", vbCritical
    End If
End Sub

Public Sub NewProject()

    Dim Msg As String
    Dim glosa As String
    
    If m_Files.count > 0 Then
        Msg = "Save files before create a new project"
        If Confirma(Msg) = vbYes Then
            Call SaveProject
        End If
    End If
    
    'cerrar todos los archivos del projecto anterior
    Call CloseProject
    
    m_Nombre = InputStr("Name", "New Project")
    
    If Len(m_Nombre) = 0 Then
        Exit Sub
    End If
    
    If Len(m_LastPath) = 0 Then
        m_LastPath = App.Path
    End If
    
    'guardar el nombre del projecto
    glosa = "JavaScript Plus Project Files (*.jsplusprj)|*.jsplusprj|"
    glosa = glosa & "All Files (*.*)|*.*"
    
    If Not Cdlg.VBGetSaveFileName(m_FileName, , , glosa, , m_LastPath, "Save Project As ...", "jsplusprj") Then
        m_Nombre = vbNullString
        Exit Sub
    End If

    m_LastPath = util.PathArchivo(m_FileName)
    
    util.GrabaIni m_FileName, "project", "name", m_Nombre
    
    frmMain.picStatus.Panels(7).Text = "Project : " & m_Nombre
    
    Call ManageProject
    
End Sub

Public Sub OpenProject()

    Dim k As Integer
    Dim pFile As cFile
            
    Dim glosa As String
    Dim linea As String
    Dim Num
    
    If Len(m_Nombre) > 0 Then
        Call SaveProject
    End If
    
    Call CloseProject
    
    If Len(m_LastPath) = 0 Then
        m_LastPath = App.Path
    End If
    
    glosa = "JavaScript Plus Project Files (*.jsplusprj)|*.jsplusprj|"
    glosa = glosa & "All Files (*.*)|*.*"
        
    If Not Cdlg.VBGetOpenFileName(m_FileName, , , , , , glosa, , m_LastPath, "Open Project ...", "jsplusprj", frmMain.hwnd) Then
        m_Nombre = vbNullString
        Exit Sub
    End If
    
    m_Nombre = util.LeeIni(m_FileName, "project", "name")
    
    If Len(m_Nombre) = 0 Then
        MsgBox "This files is not a valid JavaScript Plus! project file", vbCritical
        Exit Sub
    End If
    
    'cargar los archivos
    Num = Trim$(util.LeeIni(m_FileName, "files", "num"))
    
    If Len(Num) = 0 Then
        m_Nombre = vbNullString
        Exit Sub
    ElseIf Not IsNumeric(Num) Then
        m_Nombre = vbNullString
        Exit Sub
    End If
        
    frmMain.picStatus.Panels(7).Text = "Project : " & m_Nombre
    
    For k = 1 To Num
        'leer archivos del projecto
        linea = util.LeeIni(m_FileName, "files", "file" & k)
        If Len(linea) > 0 Then
            Set pFile = New cFile
            pFile.filename = linea
            
            If ArchivoExiste2(linea) Then
                AddFile pFile
                frmMain.opeEdit linea
            Else
                MsgBox "Warning : " & linea & " don't exists", vbCritical
            End If
        End If
    Next k
    
End Sub

Public Sub SaveProject()

    Dim k As Integer
    Dim pFile As cFile
    Dim glosa As String
    
    If Len(m_Nombre) = 0 Then
        MsgBox "There is no a current project active", vbCritical
        Exit Sub
    End If
    
    If Len(m_LastPath) = 0 Then
        m_LastPath = App.Path
    End If
    
    glosa = "JavaScript Plus Project Files (*.jsplusprj)|*.jsplusprj|"
    glosa = glosa & "All Files (*.*)|*.*"
    
    If Len(m_FileName) = 0 Then
        If Not Cdlg.VBGetSaveFileName(m_FileName, , , glosa, , m_LastPath, "Save Project As ...", "jsplusprj") Then
            Exit Sub
        End If
    End If
        
    util.BorrarArchivo m_FileName
    
    If m_Files.count > 0 Then
        util.GrabaIni m_FileName, "project", "name", m_Nombre
        For k = 1 To m_Files.count
            Set pFile = New cFile
            Set pFile = m_Files.ITem(k)
                
            Call util.GrabaIni(m_FileName, "files", "file" & k, pFile.filename)
        Next k
        
        Call util.GrabaIni(m_FileName, "files", "num", m_Files.count)
        
        'MsgBox "All Ok", vbCritical
    End If
    
End Sub

Public Property Get nombre() As String
    nombre = m_Nombre
End Property

Public Property Let nombre(ByVal pNombre As String)
    m_Nombre = pNombre
End Property
