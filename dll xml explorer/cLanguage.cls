VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type eColors
    bookmark As String
    bookmark_bk As String
    resalte_linea  As String
    hdividerline  As String
    vdividerline  As String
    comentarios  As String
    comentarios_bk  As String
    comentarios_font  As String
    palabras_clave  As String
    palabras_clave_bk  As String
    palabras_clave_font  As String
    numeros_de_linea  As String
    numeros_de_linea_bk  As String
    numeros_de_linea_font  As String
    numeros  As String
    numeros_bk  As String
    numeros_font  As String
    operadores  As String
    operadores_bk  As String
    operadores_font  As String
    scopekeyword  As String
    scopekeyword_bk  As String
    scopekeyword_font  As String
    cadenas  As String
    cadenas_bk  As String
    cadenas_font  As String
    texto  As String
    texto_bk  As String
    texto_font  As String
    fondo_ventana  As String
    fontname  As String
    FontSize  As String
    FontBold  As String
    FontItalic  As String
    TabSize  As String
    tag_name  As String
    tag_name_bk  As String
    tag_name_font  As String
    tag_element  As String
    tag_element_bk  As String
    tag_element_font  As String
    tag_entity  As String
    tag_entity_bk  As String
    tag_entity_font  As String
    tag_text  As String
    tag_text_bk  As String
    tag_text_font  As String
End Type

Private Type eLenguajes
    nombre As String
    arr_ext() As String
    defcolor As eColors
End Type

Private m_arr_lista() As eLenguajes

Private Function file_ext(ByVal FileName As String) As String

    Dim k As Integer
    Dim ret As String
    
    If InStr(FileName, ".") Then
        For k = Len(FileName) To 1 Step -1
            If Mid$(FileName, k, 1) = "." Then
                ret = LCase$(Mid$(FileName, k + 1))
                Exit For
            End If
        Next k
    End If
    
    file_ext = ret
    
End Function

Public Function IsValidExt(ByVal File As String) As Boolean

    Dim k As Integer
    Dim j As Integer
    Dim ext As String
    Dim ret As Boolean
    
    ext = util.VBArchivoSinPath(File)
    If InStr(ext, ".") Then
        ext = Mid$(ext, InStr(1, ext, ".") + 1)
        For k = 1 To UBound(m_arr_lista)
            For j = 1 To UBound(m_arr_lista(k).arr_ext)
                If m_arr_lista(k).arr_ext(j) = ext Then
                    ret = True
                    GoTo salir
                End If
            Next j
        Next k
    End If
    
salir:
    IsValidExt = ret
    
End Function

Public Sub Load(ByVal Path As String)

    Dim s As String
    Dim k As Integer
    Dim objTempLang As New Language
    Dim src As New cStringBuilder
    Dim c As Integer
    
    Dim arr_data() As String
    
    c = 1
    
    CSGlobals.UnregisterAllLanguages
    
    s = Dir(Path)
    
    Do Until s = ""
        If VBA.Right$(s, 3) = "def" Then
            
            ReDim Preserve m_arr_lista(c)
            ReDim m_arr_lista(c).arr_ext(0)
                        
            s = Path & s
            m_arr_lista(c).nombre = util.LeeIni(s, "Name", "ele")
            
            'leer informacion del archivo de lenguaje
            get_info_section "extension", arr_data, s
            For k = 1 To UBound(arr_data)
                ReDim Preserve m_arr_lista(c).arr_ext(k)
                m_arr_lista(c).arr_ext(k) = LCase$(arr_data(k))
            Next k
            
            objTempLang.CaseSensitive = util.LeeIni(s, "CaseSensitive", "ele")
            objTempLang.EscapeChar = util.LeeIni(s, "EscapeChar", "ele")
                        
            'keywords
            get_info_section "Keywords", arr_data, s
            For k = 1 To UBound(arr_data)
                If Len(arr_data(k)) > 0 Then
                    src.Append arr_data(k) & vbLf
                End If
            Next k
            
            objTempLang.Keywords = ""
            If Len(src.ToString) > 1 Then
                objTempLang.Keywords = Left$(src.ToString, Len(src.ToString) - 1)
            End If
            Set src = Nothing
            
            'multi line comments 1
            get_info_section "MultiLineComments1", arr_data, s
            For k = 1 To UBound(arr_data)
                If Len(arr_data(k)) > 0 Then
                    src.Append arr_data(k) & vbLf
                End If
            Next k
            
            objTempLang.MultiLineComments1 = ""
            If Len(src.ToString) > 1 Then
                objTempLang.MultiLineComments1 = Left$(src.ToString, Len(src.ToString) - 1)
            End If
            Set src = Nothing
                        
            'multi line comments 2
            get_info_section "MultiLineComments2", arr_data, s
            For k = 1 To UBound(arr_data)
                If Len(arr_data(k)) > 0 Then
                    src.Append arr_data(k) & vbLf
                End If
            Next k
            
            objTempLang.MultiLineComments2 = ""
            If Len(src.ToString) > 1 Then
                objTempLang.MultiLineComments2 = Left$(src.ToString, Len(src.ToString) - 1)
            End If
            Set src = Nothing
            
            'operators
            get_info_section "Operators", arr_data, s
            For k = 1 To UBound(arr_data)
                If Len(arr_data(k)) > 0 Then
                    src.Append arr_data(k) & vbLf
                End If
            Next k
            
            objTempLang.Operators = ""
            If Len(src.ToString) > 1 Then
                objTempLang.Operators = Left$(src.ToString, Len(src.ToString) - 1)
            End If
            Set src = Nothing
            
            objTempLang.ScopeKeywords1 = util.LeeIni(s, "ScopeKeywords1", "ele")
            objTempLang.ScopeKeywords2 = util.LeeIni(s, "ScopeKeywords2", "ele")
            objTempLang.SingleLineComments = util.LeeIni(s, "SingleLineComments", "ele")
            
            'string delims
            get_info_section "StringDelims", arr_data, s
            For k = 1 To UBound(arr_data)
                If Len(arr_data(k)) > 0 Then
                    src.Append arr_data(k) & vbLf
                End If
            Next k
            objTempLang.StringDelims = ""
            If Len(src.ToString) > 1 Then
                objTempLang.StringDelims = Left$(src.ToString, Len(src.ToString) - 1)
            End If
            Set src = Nothing
                        
            objTempLang.Style = util.LeeIni(s, "Style", "ele")
            
            'tag atributes names
            get_info_section "TagAttributeNames", arr_data, s
            For k = 1 To UBound(arr_data)
                If Len(arr_data(k)) > 0 Then
                    src.Append arr_data(k) & vbLf
                End If
            Next k
            
            objTempLang.TagAttributeNames = ""
            If Len(src.ToString) > 1 Then
                objTempLang.TagAttributeNames = Left$(src.ToString, Len(src.ToString) - 1)
            End If
            Set src = Nothing
            
            'tag element names
            get_info_section "TagElementNames", arr_data, s
            For k = 1 To UBound(arr_data)
                If Len(arr_data(k)) > 0 Then
                    src.Append arr_data(k) & vbLf
                End If
            Next k
            
            objTempLang.TagElementNames = ""
            If Len(src.ToString) > 1 Then
                objTempLang.TagElementNames = Left$(src.ToString, Len(src.ToString) - 1)
            End If
            Set src = Nothing
            
            'TagEntities
            get_info_section "TagEntities", arr_data, s
            For k = 1 To UBound(arr_data)
                If Len(arr_data(k)) > 0 Then
                    src.Append arr_data(k) & vbLf
                End If
            Next k
            
            objTempLang.TagEntities = ""
            If Len(src.ToString) > 1 Then
                objTempLang.TagEntities = Left$(src.ToString, Len(src.ToString) - 1)
            End If
            Set src = Nothing
            
            objTempLang.TerminatorChar = util.LeeIni(s, "TerminatorChar", "ele")
            
            CSGlobals.RegisterLanguage m_arr_lista(c).nombre, objTempLang
            
            Set objTempLang = Nothing
            
            'ahora leer la configuracion de colores para el lenguaje seleccionado
            m_arr_lista(c).defcolor.bookmark = util.LeeIni(s, "colors", "bookmark")
            m_arr_lista(c).defcolor.bookmark_bk = util.LeeIni(s, "colors", "bookmark_bk")
            m_arr_lista(c).defcolor.cadenas = util.LeeIni(s, "colors", "cadenas")
            m_arr_lista(c).defcolor.cadenas_bk = util.LeeIni(s, "colors", "cadenas_bk")
            m_arr_lista(c).defcolor.cadenas_font = util.LeeIni(s, "colors", "cadenas_font")
            m_arr_lista(c).defcolor.comentarios = util.LeeIni(s, "colors", "comentarios")
            m_arr_lista(c).defcolor.comentarios_bk = util.LeeIni(s, "colors", "comentarios_bk")
            m_arr_lista(c).defcolor.comentarios_font = util.LeeIni(s, "colors", "comentarios_font")
            m_arr_lista(c).defcolor.fondo_ventana = util.LeeIni(s, "colors", "fondo_ventana")
            m_arr_lista(c).defcolor.FontBold = util.LeeIni(s, "colors", "FontBold")
            m_arr_lista(c).defcolor.FontItalic = util.LeeIni(s, "colors", "FontItalic")
            m_arr_lista(c).defcolor.fontname = util.LeeIni(s, "colors", "fontname")
            m_arr_lista(c).defcolor.FontSize = util.LeeIni(s, "colors", "FontSize")
            m_arr_lista(c).defcolor.hdividerline = util.LeeIni(s, "colors", "hdividerline")
            m_arr_lista(c).defcolor.numeros = util.LeeIni(s, "colors", "numeros")
            m_arr_lista(c).defcolor.numeros_bk = util.LeeIni(s, "colors", "numeros_bk")
            m_arr_lista(c).defcolor.numeros_de_linea = util.LeeIni(s, "colors", "numeros_de_linea")
            m_arr_lista(c).defcolor.numeros_de_linea_bk = util.LeeIni(s, "colors", "numeros_de_linea_bk")
            m_arr_lista(c).defcolor.numeros_de_linea_font = util.LeeIni(s, "colors", "numeros_de_linea_font")
            m_arr_lista(c).defcolor.numeros_font = util.LeeIni(s, "colors", "numeros_font")
            m_arr_lista(c).defcolor.operadores = util.LeeIni(s, "colors", "operadores")
            m_arr_lista(c).defcolor.operadores_bk = util.LeeIni(s, "colors", "operadores_bk")
            m_arr_lista(c).defcolor.operadores_font = util.LeeIni(s, "colors", "operadores_font")
            m_arr_lista(c).defcolor.palabras_clave = util.LeeIni(s, "colors", "palabras_clave")
            m_arr_lista(c).defcolor.palabras_clave_bk = util.LeeIni(s, "colors", "palabras_clave_bk")
            m_arr_lista(c).defcolor.palabras_clave_font = util.LeeIni(s, "colors", "palabras_clave_font")
            m_arr_lista(c).defcolor.resalte_linea = util.LeeIni(s, "colors", "resalte_linea")
            m_arr_lista(c).defcolor.scopekeyword = util.LeeIni(s, "colors", "scopekeyword")
            m_arr_lista(c).defcolor.scopekeyword_bk = util.LeeIni(s, "colors", "scopekeyword_bk")
            m_arr_lista(c).defcolor.scopekeyword_font = util.LeeIni(s, "colors", "scopekeyword_font")
            m_arr_lista(c).defcolor.TabSize = util.LeeIni(s, "colors", "TabSize")
            m_arr_lista(c).defcolor.tag_element = util.LeeIni(s, "colors", "tag_element")
            m_arr_lista(c).defcolor.tag_element_bk = util.LeeIni(s, "colors", "tag_element_bk")
            m_arr_lista(c).defcolor.tag_element_font = util.LeeIni(s, "colors", "tag_element_font")
            m_arr_lista(c).defcolor.tag_entity = util.LeeIni(s, "colors", "tag_entity")
            m_arr_lista(c).defcolor.tag_entity_bk = util.LeeIni(s, "colors", "tag_entity_bk")
            m_arr_lista(c).defcolor.tag_entity_font = util.LeeIni(s, "colors", "tag_entity_font")
            m_arr_lista(c).defcolor.tag_name = util.LeeIni(s, "colors", "tag_name")
            m_arr_lista(c).defcolor.tag_name_bk = util.LeeIni(s, "colors", "tag_name_bk")
            m_arr_lista(c).defcolor.tag_name_font = util.LeeIni(s, "colors", "tag_name_font")
            m_arr_lista(c).defcolor.tag_text = util.LeeIni(s, "colors", "tag_text")
            m_arr_lista(c).defcolor.tag_text_bk = util.LeeIni(s, "colors", "tag_text_bk")
            m_arr_lista(c).defcolor.tag_text_font = util.LeeIni(s, "colors", "tag_text_font")
            m_arr_lista(c).defcolor.texto = util.LeeIni(s, "colors", "texto")
            m_arr_lista(c).defcolor.texto_bk = util.LeeIni(s, "colors", "texto_bk")
            m_arr_lista(c).defcolor.texto_font = util.LeeIni(s, "colors", "texto_font")
            m_arr_lista(c).defcolor.vdividerline = util.LeeIni(s, "colors", "vdividerline")
            
            c = c + 1
        End If
        s = Dir
    Loop
  
End Sub
Public Sub SetLang(ByVal FileName As String, txtCode As CodeSense)

    Dim ext As String
    Dim k As Integer
    Dim j As Integer
    
    If InStr(FileName, ".") Then
        ext = file_ext(FileName)
        If ext = "tidy" Then ext = "htm"
    Else
        Dim default_ext As String
    
        default_ext = LCase$(util.LeeIni(IniPath, "startup", "default_ext"))
        If Len(default_ext) = 0 Then
            ext = "js"
        Else
            ext = default_ext
        End If
    End If
    
    For k = 1 To UBound(m_arr_lista)
        For j = 1 To UBound(m_arr_lista(k).arr_ext)
            If m_arr_lista(k).arr_ext(j) = ext Then
                
                'lenguaje
                txtCode.Language = m_arr_lista(k).nombre
                
                'resalte de linea
                Call txtCode.SetColor(cmClrHighlightedLine, m_arr_lista(k).defcolor.resalte_linea)
                
                'comentarios
                Call txtCode.SetColor(cmClrComment, m_arr_lista(k).defcolor.comentarios)
                'Call txtCode.SetColor(cmClrCommentBk, m_arr_lista(k).defcolor.comentarios_bk)
                Call txtCode.SetFontStyle(cmStyComment, cmFontNormal)
                
                'keywords
                Call txtCode.SetColor(cmClrKeyword, m_arr_lista(k).defcolor.palabras_clave)
                'Call txtCode.SetColor(cmClrKeywordBk, m_arr_lista(k).defcolor.palabras_clave_bk)
                Call txtCode.SetFontStyle(cmStyKeyword, cmFontNormal)
                
                'line numbers
                Call txtCode.SetColor(cmClrLineNumber, m_arr_lista(k).defcolor.numeros_de_linea)
                Call txtCode.SetColor(cmClrLineNumberBk, m_arr_lista(k).defcolor.numeros_de_linea_bk)
                Call txtCode.SetFontStyle(cmStyLineNumber, cmFontNormal)
                
                'numeros
                Call txtCode.SetColor(cmClrNumber, m_arr_lista(k).defcolor.numeros)
                'Call txtCode.SetColor(cmClrNumberBk, m_arr_lista(k).defcolor.numeros_bk)
                Call txtCode.SetFontStyle(cmStyNumber, cmFontNormal)
                
                'operadores
                Call txtCode.SetColor(cmClrOperator, m_arr_lista(k).defcolor.operadores)
                'Call txtCode.SetColor(cmClrOperatorBk, m_arr_lista(k).defcolor.operadores_bk)
                Call txtCode.SetFontStyle(cmStyOperator, cmFontNormal)
                
                'alcance
                Call txtCode.SetColor(cmClrScopeKeyword, m_arr_lista(k).defcolor.scopekeyword)
                'Call txtCode.SetColor(cmClrScopeKeywordBk, m_arr_lista(k).defcolor.scopekeyword_bk)
                Call txtCode.SetFontStyle(cmStyScopeKeyword, cmFontNormal)
                
                'strings
                Call txtCode.SetColor(cmClrString, m_arr_lista(k).defcolor.cadenas)
                'Call txtCode.SetColor(cmClrStringBk, m_arr_lista(k).defcolor.cadenas_bk)
                Call txtCode.SetFontStyle(cmStyString, cmFontNormal)
                
                'texto
                Call txtCode.SetColor(cmClrText, m_arr_lista(k).defcolor.texto)
                'Call txtCode.SetColor(cmClrTextBk, m_arr_lista(k).defcolor.texto_bk)
                Call txtCode.SetFontStyle(cmStyText, cmFontNormal)
                
                'tag html
                Call txtCode.SetColor(cmClrTagAttributeName, m_arr_lista(k).defcolor.tag_name)
                'Call txtCode.SetColor(cmClrTagAttributeNameBk, m_arr_lista(k).defcolor.tag_name_bk)
                Call txtCode.SetFontStyle(cmStyTagAttributeName, cmFontNormal)
                
                'tag element
                Call txtCode.SetColor(cmClrTagElementName, m_arr_lista(k).defcolor.tag_element)
                'Call txtCode.SetColor(cmClrTagElementNameBk, m_arr_lista(k).defcolor.tag_element_bk)
                Call txtCode.SetFontStyle(cmStyTagElementName, cmFontNormal)
                
                'tag entity
                Call txtCode.SetColor(cmClrTagEntity, m_arr_lista(k).defcolor.tag_entity)
                'Call txtCode.SetColor(cmClrTagEntityBk, m_arr_lista(k).defcolor.tag_entity_bk)
                Call txtCode.SetFontStyle(cmStyTagEntity, cmFontNormal)
                
                'tag text
                Call txtCode.SetColor(cmClrTagText, m_arr_lista(k).defcolor.tag_text)
                'Call txtCode.SetColor(cmClrTagTextBk, m_arr_lista(k).defcolor.tag_text_bk)
                Call txtCode.SetFontStyle(cmStyTagText, cmFontNormal)
                
                'fondo de la ventana
                Call txtCode.SetColor(cmClrWindow, m_arr_lista(k).defcolor.fondo_ventana)
                '******
                
                txtCode.Font.Name = util.Nvl(m_arr_lista(k).defcolor.fontname, "courier new")
                txtCode.Font.Size = util.Nvl(m_arr_lista(k).defcolor.FontSize, "10")
                txtCode.Font.Bold = util.Nvl(m_arr_lista(k).defcolor.FontBold, "0")
                txtCode.Font.Italic = util.Nvl(m_arr_lista(k).defcolor.FontItalic, "0")
                txtCode.TabSize = util.Nvl(m_arr_lista(k).defcolor.TabSize, "3")

                GoTo salir
            End If
        Next j
    Next k
salir:
    
End Sub


