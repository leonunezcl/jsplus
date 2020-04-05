VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Funciones As New Collection

Public Sub Add(m_Funcion As CFuncion)
    
    Dim j As Integer
    
    j = m_Funciones.Count + 1
    
    'agregar la funcion a la coleccion
    m_Funciones.Add m_Funcion, "k" & j
                
End Sub


Public Sub Clear()
    
    Dim k As Integer
    
    For k = m_Funciones.Count To 1 Step -1
        m_Funciones.Remove k
    Next k
    
End Sub

Private Function ContarBloques(ByVal archivo As String, ByVal c As Integer, _
                               ByRef reterror As Boolean, ByRef linea_sgte As Integer) As Integer

    Dim nfreefile As Long
    Dim linea As String
    Dim LineaAux As String
    Dim ret As Integer
    Dim k As Integer
    Dim j As Integer
    Dim I As Integer
    Dim p1 As Integer
    Dim P2 As Integer
    Dim f_block_comment As Boolean
    Dim desde As Integer
    
    nfreefile = FreeFile
    j = 1
    k = 1
    desde = c
    Open archivo For Input As #nfreefile
        Do While Not EOF(nfreefile)
            Line Input #nfreefile, linea
            
            'If k > 370 Then
            '    Debug.Print "stop!"
            'End If
            
            'empezar a leer desde la linea detectada
            If (k >= desde) Then
                'limpiar la linea de caracteres extraños
                linea = Trim$(Util.SacarBasura(linea))
                
                'que no este en blanco
                If Len(linea) = 0 Then GoTo SgteLinea
                
                'analizar si se esta leyendo un comentario o se esta en un bloque de comentario
                If EsLineaComentario(linea, f_block_comment) Then GoTo SgteLinea
            
                'verificar si hubo un comienzo de lectura de funcion
                If f_block_comment Then GoTo SgteLinea
            
                'que no sea un comentario simple
                If VBA.Left$(linea, 2) = "//" Then GoTo SgteLinea
                
                'verificar si es un comienzo de comentario multilinea pero de una sola linea
                If VBA.Left$(linea, 2) = "/*" And VBA.Right$(linea, 2) = "*/" Then GoTo SgteLinea
                
                'verificar si es un comentario multilinea en mas de 1 linea
                If VBA.Left$(linea, 2) = "/*" Then
                    'verificar si ya hubo un comienzo de bloque
                    If Not f_block_comment Then
                        'hay un bloque de comentario
                        f_block_comment = True
                        GoTo SgteLinea
                    End If
                ElseIf VBA.Left$(linea, 2) = "*/" Then  'fin de linea solo
                    'verificar si hubo un comienzo de lectura de funcion
                    If f_block_comment Then
                        f_block_comment = False
                        GoTo SgteLinea
                    End If
                ElseIf VBA.Right$(linea, 2) = "*/" Then  'fin de linea solo
                    'verificar si hubo un comienzo de lectura de funcion
                    If f_block_comment Then
                        f_block_comment = False
                        GoTo SgteLinea
                    End If
                Else
                    If f_block_comment Then
                        GoTo SgteLinea
                    End If
                End If
                
                'verificar si se llego entonces a la siguiente funcion
                If j > 1 Then
                    If IsFunction(linea) Then
                        'verificar si la funcion comienza en el termino de la anterior
                        'esto marca el fin de la rutina
                        If InStr(1, linea, "}") > 0 Then
                            'marcar un incremento + a los marcadores de bloque
                            ret = ret + 1
                            
                            'numero de linea donde comienza la siguiente funcion
                            linea_sgte = k - 1
                        ElseIf InStr(1, linea, "{") > 0 Then
                            'numero de linea donde comienza la siguiente funcion
                            linea_sgte = k - 1
                        Else
                            'numero de linea donde comienza la siguiente funcion
                            linea_sgte = k - 1
                        End If
                        
                        reterror = False
                        GoTo SalirLectura
                    End If
                Else
                    j = 2
                End If
                
                'extraer solo el codigo necesario
                '1.- extraer los bloques de cadena. Tanto con ' y con ""
                '2 extraer las palabras reservadas
                linea = CortaComentario(linea)
                
                'ahora contar los inicio y fin de bloques {} que existan en la linea
                For I = 1 To Len(linea)
                    If Mid$(linea, I, 1) = "{" Then ret = ret + 1
                    If Mid$(linea, I, 1) = "}" Then ret = ret + 1
                Next I
            End If
SgteLinea:
            k = k + 1
        Loop
SalirLectura:
    Close #nfreefile
  
    ContarBloques = ret
    
    Exit Function
ErrorContarBloques:
    MsgBox "ContarBloques : " & Err & " " & Error$, vbCritical
    reterror = True
    Exit Function
End Function

Private Function EsLineaComentario(ByVal linea As String, ByRef f_block_comment) As Boolean

    'que no sea un comentario simple
    If VBA.Left$(linea, 2) = "//" Then
        EsLineaComentario = True
        Exit Function
    End If
    
    'verificar si es un comienzo de comentario multilinea pero de una sola linea
    If VBA.Left$(linea, 2) = "/*" And VBA.Right$(linea, 2) = "*/" Then
        EsLineaComentario = True
        Exit Function
    End If
    
    'verificar si es un comentario multilinea en mas de 1 linea
    If VBA.Left$(linea, 2) = "/*" Then
        'verificar si ya hubo un comienzo de bloque
        If Not f_block_comment Then
            'hay un bloque de comentario
            f_block_comment = True
            EsLineaComentario = True
            Exit Function
        End If
    ElseIf VBA.Left$(linea, 2) = "*/" Then  'fin de linea solo
        'verificar si hubo un comienzo de lectura de funcion
        If f_block_comment Then
            f_block_comment = False
            EsLineaComentario = True
            Exit Function
        End If
    ElseIf VBA.Right$(linea, 2) = "*/" Then  'fin de linea solo
        'verificar si hubo un comienzo de lectura de funcion
        If f_block_comment Then
            EsLineaComentario = True
            Exit Function
        End If
    End If
            
    EsLineaComentario = False
        
End Function

Private Function ExtraeNombreFuncion(ByVal linea As String) As String

    Dim ret As String
    Dim POS As Integer
    
    linea = Replace(linea, ",", "")
    
    If VBA.Left$(linea, 9) = "function " Then   'la funcion empieza en el comienzo de la linea
        linea = Mid$(linea, 10)
        
        If InStr(1, linea, "(") > 0 Then
            ret = VBA.Left$(linea, InStr(1, linea, "(") - 1)
        End If
    Else
        'la funcion empieza en otro lado
        POS = InStr(linea, "function ")
        
        If POS > 0 Then
            If InStr(linea, "function ") Then
                linea = Mid$(linea, InStr(linea, "function ") + 10)
            End If
            
            If InStr(1, linea, "(") > 0 Then
                linea = VBA.Left$(linea, InStr(1, linea, "(") + 1)
            End If
        End If
    End If
    
    ExtraeNombreFuncion = ret
    
End Function

Private Function ExtraeParametros(ByVal linea As String) As String

    Dim ret As String
    
    'buscar comienzo de parametro
    If InStr(1, linea, "(") Then
        linea = Mid$(linea, InStr(1, linea, "(") + 1)
        If InStr(1, linea, ")") Then
            linea = Replace(VBA.Left$(linea, InStr(1, linea, ")") - 1), " ", "")
            ret = linea
        End If
    End If
    
    ExtraeParametros = ret
    
End Function

Private Function IsFunction(ByVal linea As String) As Boolean

    'valida que no sea una linea en blanco y que no se diga por
    'valida una linea que venga con comentario
    linea = LCase$(linea)
    
    If Not IsValidLine(linea) Then
        Exit Function
    ElseIf InStr(linea, "function") Then
        IsFunction = True
    Else
        Exit Function
    End If
    
End Function

'extraer comentario de la derecha
Private Function CortaComentario(ByRef linea As String) As String

    Dim j As Integer
    Dim c As Integer
    Dim p As Integer
    Dim p1 As Long
    Dim P2 As Long
    Dim lSearch As String
    
    linea = Trim$(linea)
    
    'extraer comentareos
    If InStr(linea, "'") <> 0 Then
        If IsNotInQuote(linea, "'") Then
            'remove the comment from the line
            linea = VBA.Left$(linea, InStr(linea, "'") - 1)
        End If
    End If
            
    'extraer comilla doble
    lSearch = linea
    If InStr(1, lSearch, Chr$(34)) > 0 Then
        Do
            p1 = InStr(1, lSearch, Chr$(34))
            
            If p1 > 0 Then
                'buscar la otra posicion
                P2 = InStr(p1 + 1, lSearch, Chr$(34))
                If P2 > 0 Then
                    lSearch = VBA.Left$(lSearch, p1 - 1) & Mid$(lSearch, P2 + 1)
                Else
                    linea = lSearch
                    Exit Do
                End If
            Else
                linea = lSearch
                Exit Do
            End If
        Loop
    End If
            
    'valida el fin de comentario
    If InStr(linea, "'") > 0 Then
        For j = Len(linea) To 1 Step -1
            If Mid(linea, j, 1) = "'" Then
                c = c + 1
                p = j
                Exit For
            End If
        Next j
        
        'hay comentareo al lado ?
        If c = 1 Then
            linea = VBA.Left$(linea, j - 1)
        End If
    End If
        
    CortaComentario = linea
    
End Function
Private Function IsNotInQuote(ByVal strText As String, ByVal strWords As String) As Boolean
    'This function will tell you if the specified text is in quotes within
    'the second string. It does this by counting the number of quotation
    'marks before the specified strWords. If the number is even, then the
    'strWords are not in qototes, otherwise they are.
    
    'the quotation mark, " , is ASCII character 34
    
    Dim lngGotPos As Long
    Dim lngCounter As Long
    Dim lngNextPos As Long
    
    'find where the position of strWords in strText
    lngGotPos = InStr(1, strText, strWords)
    If lngGotPos = 0 Then
        IsNotInQuote = True
        Exit Function
    End If
    
    'start counting the number of quotation marks
    lngNextPos = 0
    Do
        lngNextPos = InStr(lngNextPos + 1, strText, Chr(34))
        
        If (lngNextPos <> 0) And (lngNextPos < lngGotPos) Then
            'quote found, add to total
            lngCounter = lngCounter + 1
        End If
    Loop Until (lngNextPos = 0) Or (lngNextPos >= lngGotPos)
    
    'no quotes at all found
    If lngCounter = 0 Then
        IsNotInQuote = True
        Exit Function
    End If
    
    'if the number of quotes is even, then return true, else return false
    If lngCounter Mod 2 = 0 Then
        IsNotInQuote = True
    End If
End Function

Private Function IsValidLine(ByVal linea As String) As Boolean

    Dim ret As Boolean
    
    linea = Trim$(Util.SacarBasura(linea))
    
    If Len(linea) > 0 Then
        If VBA.Left$(linea, 2) <> "//" And VBA.Left$(linea, 2) <> "/*" And VBA.Left$(linea, 2) <> "*/" Then
            ret = True
        End If
    End If
                    
    IsValidLine = ret
    
End Function

Public Function Read(ByVal TheFile As String) As Boolean

    Dim nfreefile As Long
    Dim linea As String
    Dim f_block_comment As Boolean
    Dim f_start_function As Boolean
    Dim f_block_start As Boolean
    Dim function_name As Boolean
    Dim c As Integer
    Dim j As Integer
    Dim linea_sgte As Integer
    Dim backup_file As String
    Dim work_file As String
    Dim block_count As Integer
    Dim ret_error As Boolean
    Dim fun_error As Boolean
    Dim nombre_funcion As String
    
    Dim m_Funcion As CFuncion
    Dim m_Error As CError
    
    'copiar la funcion a un archivo de respaldo
    backup_file = Util.PathArchivo(TheFile) & Util.VBArchivoSinPath(TheFile) & ".bak"
    Util.CopiarArchivo TheFile, backup_file
    
    'copiar la funcion a un archivo de trabajo
    work_file = Util.PathArchivo(TheFile) & Util.VBArchivoSinPath(TheFile) & ".wrk"
    Util.CopiarArchivo TheFile, work_file
    
    'contador de funciones
    j = 1
    c = 1
    nfreefile = FreeFile
    Open TheFile For Input Shared As #nfreefile
        Do While Not EOF(nfreefile)
            Line Input #nfreefile, linea
                        
            'Debug.Print "linea : " & linea & " numero : " & c
                        
            'limpiar la linea de caracteres extraños
            linea = Trim$(Util.SacarBasura(linea))
            
            'que no este en blanco
            If Len(linea) = 0 Then GoTo SgteLinea
                                    
            'analizar si se esta leyendo un comentario o se esta en un bloque de comentario
            If EsLineaComentario(linea, f_block_comment) Then GoTo SgteLinea
            
            'verificar si hubo un comienzo de lectura de funcion
            If f_block_comment Then GoTo SgteLinea
            
            'verificar si se comenzo el inicio de una funcion
            If (c >= linea_sgte) Then
                'verificar si se esta leyendo una funcion
                If Not IsFunction(linea) Then GoTo SgteLinea
                
                'comenzar a leer una funcion
                f_start_function = True
                linea_sgte = 0
                ret_error = False
                
                'verificar integridad de la funcion contando los bloques dentro de la funcion
                block_count = ContarBloques(work_file, c, ret_error, linea_sgte)
                
                'hubo error al leer la funcion
                If ret_error Then GoTo SalirLectura
                
                'verificar integridad del bloque
                fun_error = False
                If (block_count Mod 2) > 0 Then
                    'error en la integridad de la funcion. La funcion tiene errores de sintaxis
                    fun_error = True
                End If
                
                'configuracion de la funcion
                nombre_funcion = ExtraeNombreFuncion(linea)
                
                If Len(nombre_funcion) = 0 Then GoTo SgteLinea
                
                Set m_Funcion = New CFuncion
                m_Funcion.Id = j
                If Len(nombre_funcion) > 0 Then
                    m_Funcion.Nombre = nombre_funcion
                    m_Funcion.Parametros = ExtraeParametros(linea)
                Else
                    m_Funcion.Nombre = "error_extract_name_function"
                End If
                                
                'agregar la funcion a la coleccion
                Call Add(m_Funcion)
                
                'si hay errores acumular los errores
                If fun_error Then
                    Set m_Error = New CError
                                        
                    'agregar el detalle del error
                    m_Error.StrError = "This function have problem at analizing the blocks {}. Seem have consistency problem."
                    m_Error.linea = c
                    
                    m_Funcion.AddError m_Error
                    
                    Set m_Error = Nothing
                End If
                
                Set m_Funcion = Nothing
                j = j + 1
                'si hay error en la lectura de los bloques entonces agregar a la lista de errores
            End If
SgteLinea:
        c = c + 1
        Loop
SalirLectura:
    Close #nfreefile
    
    'hacer una segunda pasada para ver las posibles funciones que se quedaron afuera despues del primer analisis
    
    'borrar el archivo de trabajo
    Util.BorrarArchivo work_file
    
    Read = True
    
    Exit Function
    
ErrorRead:
    MsgBox "Read : " & Err & " " & Error$, vbCritical
    Read = False
    Exit Function
    
End Function

Public Property Get Funciones() As Collection
    Set Funciones = m_Funciones
End Property

Public Property Set Funciones(ByVal pFunciones As Collection)
    Set m_Funciones = pFunciones
End Property


Public Function Search(ByVal strFuncion As String, ByRef strParam As String) As Boolean

    Dim k As Integer
    Dim m_Funcion As CFuncion
    Dim ret As Boolean
    
    strParam = ""
    For k = 1 To m_Funciones.Count
        Set m_Funcion = New CFuncion
        Set m_Funcion = m_Funciones.Item(k)
        
        If m_Funcion.Nombre = strFuncion Then
            strParam = m_Funcion.Parametros
            ret = True
            Exit For
        End If
        
        Set m_Funcion = Nothing
    Next k
    
    Set m_Funcion = Nothing
    
    Search = ret
    
End Function

Private Sub Class_Terminate()
    Clear
End Sub

