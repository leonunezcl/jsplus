VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCodeExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_FileName As String
Private m_Images As New Collection
Private m_Links As New Collection
Private m_Scripts As New Collection
Private m_StyleSheets As New Collection
Private m_Code() As String

Private Sub Clear()
    
    Dim k As Integer
    
    For k = m_Images.count To 1 Step -1
        m_Images.Remove k
    Next k
    
    For k = m_Links.count To 1 Step -1
        m_Links.Remove k
    Next k
    
    For k = m_Scripts.count To 1 Step -1
        m_Scripts.Remove k
    Next k
    
    For k = m_StyleSheets.count To 1 Step -1
        m_StyleSheets.Remove k
    Next k
    
    ReDim m_Code(0)
    
End Sub


Public Function Explore() As Boolean

    On Error GoTo ErrorExplore
    
    Dim k As Integer
    Dim nFreeFile As Long
    Dim buffer As String
    Dim ext As String
    Dim x As Long
    Dim IniFileTree As String
    Dim CodeTemp() As String
    Dim arr_ext() As String
    Dim fok As Boolean
    
    If m_FileName = "" Then Exit Function
    If InStr(m_FileName, ".") = 0 Then Exit Function
    
    ext = LCase$(GetFileExtension(m_FileName))
    
    If Len(ext) = 0 Then Exit Function
    
    IniFileTree = StripPath(App.Path) & "treeext.ini"
   
    get_info_section "tree", arr_ext, IniFileTree
   
    If UBound(arr_ext) = 0 Then
        Call util.GrabaIni(IniFileTree, "tree", "ext1", "htm")
        Call util.GrabaIni(IniFileTree, "tree", "ext2", "html")
        Call util.GrabaIni(IniFileTree, "tree", "ext3", "asp")
        Call util.GrabaIni(IniFileTree, "tree", "ext4", "aspx")
        get_info_section "tree", arr_ext, IniFileTree
    End If
    
    For k = 1 To UBound(arr_ext)
        If ext = LCase$(arr_ext(k)) Then
            fok = True
            Exit For
        End If
    Next k
    
    If Not fok Then Exit Function
    
    Call Clear
    
    nFreeFile = FreeFile
    k = 1
    
    'obtener el archivo
    Open m_FileName For Input As nFreeFile
        Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, buffer
            
            'Debug.Print buffer
            If InStr(buffer, Chr$(10)) Then
                CodeTemp = Split(buffer, Chr$(10))
            
                For x = 0 To UBound(CodeTemp)
                    ReDim Preserve m_Code(k)
                    m_Code(k) = CodeTemp(x)
                    k = k + 1
                Next x
            Else
                ReDim Preserve m_Code(k)
                m_Code(k) = buffer
                k = k + 1
            End If
        Loop
    Close nFreeFile
    
    'obtenido el archivo obtener el codigo a usar eliminar aquellos tags que no son necesarios
    Call procesar_tags
    
    Explore = True
    
    Exit Function
    
ErrorExplore:
    MsgBox "cCodeExplorer.Explore : " & Err & " " & Error$, vbCritical
    
End Function


Public Property Get filename() As String
    filename = m_FileName
End Property

Public Property Let filename(ByVal pFileName As String)
    m_FileName = pFileName
End Property

'solo obtener los tags correctos
Private Sub get_line_tags(ByVal linea As String, ByRef arr_tags() As String)

    Dim k As Integer
    Dim linea_aux As String
    Dim linea_tmp As String
    Dim inicio As Integer
    Dim fin As Integer
    Dim Tag As String
    
    ReDim arr_tags(0)
    
    linea_tmp = LCase$(Trim$(util.SacarBasura(linea)))
        
    'no procesar comentarios
    If (Left$(linea_tmp, 4) = "<!--" Or Left$(linea_tmp, 4) = "-->") Then
        Exit Sub
    ElseIf (Left$(linea_tmp, 9) = "<comment>" Or Left$(linea_tmp, 10) = "</comment>") Then
        Exit Sub
    ElseIf (Left$(linea_tmp, 14) = "document.write" Or Left$(linea_tmp, 16) = "document.writeln") Then
        Exit Sub
    ElseIf InStr(linea_tmp, "document.write") Or InStr(linea_tmp, "document.writeln") Then
        Exit Sub
    End If
    
    linea_aux = Trim$(util.SacarBasura(linea))
    
    k = 1
    Do
        inicio = InStr(1, linea_aux, "<")
        If inicio = 0 Then Exit Do
        If inicio <> 1 Then
            linea_aux = Mid$(linea_aux, inicio)
            inicio = InStr(1, linea_aux, "<")
            If inicio = 0 Then Exit Do
        End If
        
        fin = InStr(inicio + 1, linea_aux, ">")
        If fin = 0 Then Exit Do
        
        Tag = Mid$(linea_aux, inicio, fin)
                
        If Tag = "" Then Exit Do
        ReDim Preserve arr_tags(k)
        arr_tags(k) = Tag
        k = k + 1
        
        linea_aux = Trim$(Mid$(linea_aux, fin + 1))
    Loop
    
End Sub

Private Function get_token(ByVal linea As String, ByVal Token As String) As String

    Dim inicio As Integer
    Dim fin As Integer
    Dim linea_aux As String
    Dim k As Integer
    Dim j As Integer
    Dim ret As String
    
    linea_aux = Trim$(Mid$(linea, InStr(linea, Token)))
    linea_aux = Mid$(linea_aux, InStr(1, linea_aux, "=") + 1)
    linea_aux = Replace(linea_aux, Chr$(34), "")
            
    If InStr(1, linea_aux, Chr$(34)) > 0 Then
        inicio = InStr(1, linea_aux, Chr$(34))
        If inicio > 0 Then
            fin = InStr(inicio + 1, linea_aux, Chr$(34))
            If fin > 0 Then
                ret = Mid$(linea_aux, inicio + 1, fin - 2)
            End If
        End If
    Else
        If InStr(1, linea_aux, "/>") Then
            ret = Left$(linea_aux, InStr(1, linea_aux, ">") - 2)
        ElseIf InStr(1, linea_aux, ">") Then
            ret = Left$(linea_aux, InStr(1, linea_aux, ">") - 1)
        End If
    End If
    
    ret = Trim$(ret)
        
    If InStr(1, ret, " ") > 0 Then
        For k = 1 To Len(ret)
            If Mid(ret, k, 1) = "." Then
                For j = k + 1 To Len(ret)
                    If Mid$(ret, j, 1) = " " Then
                        ret = Left$(ret, j - 1)
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next k
    End If
        
    get_token = ret
    
End Function

Private Function get_token2(ByVal linea As String, ByVal Token As String) As String

    Dim linea_aux As String
    
    If Token = " href=" Then
        If InStr(linea, " href=") > 0 Then
            linea_aux = Trim$(Mid$(linea, InStr(linea, " href=")))
            linea_aux = Mid$(linea_aux, 7)
            If InStr(linea_aux, Chr$(34)) Then
                linea_aux = Left$(linea_aux, InStr(linea_aux, Chr$(34)) - 1)
            ElseIf InStr(linea_aux, "'") Then
                linea_aux = Left$(linea_aux, InStr(linea_aux, "'") - 1)
            End If
        End If
    ElseIf Token = " src=" Then
        If InStr(linea, " src=") > 0 Then
            linea_aux = Trim$(Mid$(linea, InStr(linea, " src=")))
            linea_aux = Mid$(linea_aux, 6)
            If InStr(linea_aux, Chr$(34)) Then
                linea_aux = Left$(linea_aux, InStr(linea_aux, Chr$(34)) - 1)
            ElseIf InStr(linea_aux, "'") Then
                linea_aux = Left$(linea_aux, InStr(linea_aux, "'") - 1)
            End If
        End If
    End If
    
    get_token2 = Trim$(linea_aux)
    
End Function

Public Property Get Images() As Collection
    Set Images = m_Images
End Property

Public Property Set Images(ByVal pImages As Collection)
    Set m_Images = pImages
End Property

Private Sub procesar_tags()

    Dim k As Long
    Dim j As Long
    Dim C As Long
    Dim arr_tags() As String
    Dim attr As String
    
    C = 1
    
    For k = 1 To UBound(m_Code)
        get_line_tags m_Code(k), arr_tags
                
        For j = 1 To UBound(arr_tags)
            'seleccionar los tags validos
            
            If InStr(LCase$(arr_tags(j)), "<script") Then
                If InStr(LCase$(arr_tags(j)), " src=") Then
                    attr = get_token2(arr_tags(j), " src=")
                    If Len(attr) > 0 Then m_Scripts.Add attr
                End If
            ElseIf InStr(LCase$(arr_tags(j)), "<link ") Then
                If InStr(LCase$(arr_tags(j)), " href=") Then
                    attr = get_token2(arr_tags(j), " href=")
                    If Len(attr) > 0 Then m_StyleSheets.Add attr
                End If
            ElseIf InStr(LCase$(arr_tags(j)), "<a ") Then
                If InStr(LCase$(arr_tags(j)), " href=") Then
                    attr = get_token2(arr_tags(j), " href=")
                    If (Left$(attr, 1) = "#" Or Left$(attr, 1) = "+" Or Left$(attr, 1) = "'") Then
                    
                    Else
                        If Len(attr) > 0 Then m_Links.Add attr
                    End If
                End If
            ElseIf InStr(LCase$(arr_tags(j)), "<img ") Then
                If InStr(LCase$(arr_tags(j)), " src=") Then
                    attr = get_token2(arr_tags(j), " src=")
                    If Len(attr) > 0 Then
                        If InStr(attr, "+") = 0 Then
                            m_Images.Add attr
                        End If
                    End If
                End If
            End If
            C = C + 1
        Next j
    Next k

End Sub

Public Property Get Links() As Collection
    Set Links = m_Links
End Property

Public Property Set Links(ByVal pLinks As Collection)
    Set m_Links = pLinks
End Property

Public Property Get Scripts() As Collection
    Set Scripts = m_Scripts
End Property

Public Property Set Scripts(ByVal pScripts As Collection)
    Set m_Scripts = pScripts
End Property

Public Property Get StyleSheets() As Collection
    Set StyleSheets = m_StyleSheets
End Property

Public Property Let StyleSheets(ByVal pStyleSheets As Collection)
    Set m_StyleSheets = pStyleSheets
End Property

Public Property Get ImagesCount() As Integer
    ImagesCount = m_Images.count
End Property


Public Property Get LinksCount() As Integer
    LinksCount = m_Links.count
End Property


Public Property Get ScriptsCount() As Integer
    ScriptsCount = m_Scripts.count
End Property


Public Property Get StyleSheetsCount() As Integer
    StyleSheetsCount = m_StyleSheets.count
End Property

