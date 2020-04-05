VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExpFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_FileName As String
Private m_Funciones As New Collection

Private Type m_InfoFun
    Name As String
    Tag As String
End Type
Private arr_fun_info() As m_InfoFun

Private m_Code() As String

Public Function Explore() As Boolean

    On Error GoTo ErrorExplore
    
    Dim k As Long
    Dim nFreeFile As Long
    Dim buffer As String
    Dim x As Long
    Dim CodeTemp() As String
    
    If m_FileName = "" Then Exit Function
        
    Call Clear
    
    nFreeFile = FreeFile
    k = 1
    
    'obtener el archivo
    Open m_FileName For Input As nFreeFile
        Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, buffer
            
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
    If UBound(m_Code) > 0 Then
        Call procesar_tags
    End If
    
    Explore = True
    
    Exit Function
    
ErrorExplore:
    MsgBox "cExpFunctions.Explore : " & Err & " " & Error$, vbCritical
    
End Function



Private Function isfunction(ByVal linea As String, ByRef inicio As Integer, ByRef proto As Boolean) As Boolean

    Dim ret As Boolean
    Dim tmp As String
    Dim j As Integer
    Dim j2 As Integer
    
    proto = True
    If InStr(linea, "=function()") Then
        inicio = InStr(linea, "=function()")
        ret = True
    ElseIf InStr(linea, "= function ()") Then
        inicio = InStr(linea, "= function ()")
        ret = True
    ElseIf InStr(linea, "=function ") Then
        inicio = InStr(linea, "=function ")
        ret = True
    ElseIf InStr(linea, "= function ") Then
        inicio = InStr(linea, "= function ")
        ret = True
    ElseIf InStr(linea, "=function(") Then
        inicio = InStr(linea, "=function(")
        ret = True
    ElseIf InStr(linea, "function ") Then
        inicio = InStr(linea, "function ")
        tmp = Mid$(linea, inicio + Len("function "))
        For j = 1 To Len(tmp)
            If Mid$(tmp, j, 1) = "(" Then
                For j2 = j + 1 To Len(tmp)
                    If Mid$(tmp, j2, 1) = ")" Then
                        proto = False
                        ret = True
                        GoTo salir
                    End If
                Next j2
            End If
        Next j
    ElseIf InStr(linea, " = function(") Then
        inicio = InStr(linea, " = function(")
        ret = True
    ElseIf InStr(linea, " = function (") Then
        inicio = InStr(linea, " = function (")
        ret = True
    End If
    
salir:
    isfunction = ret
    
End Function

Private Function left_valido(ByVal linea_aux, ByVal inicio As Integer) As Boolean

    Dim k As Integer
    
    Const chars = ",(?:]º!·#$%&/?¿¡`^-+/\_;<*@ª|"""
    
    linea_aux = Replace(linea_aux, " ", "")
    inicio = InStr(linea_aux, "function")
    For k = 1 To Len(chars)
        If Mid$(linea_aux, inicio - 1, 1) = Mid$(chars, k, 1) Then
            left_valido = False
            Exit Function
        End If
        
        If inicio > 2 Then
            If Mid$(linea_aux, inicio - 2, 1) = Mid$(chars, k, 1) Then
                left_valido = False
                Exit Function
            End If
        End If
    Next k

    left_valido = True
    
End Function

Private Sub procesar_tags()

    Dim k As Long
    Dim j As Long
    Dim C As Long
    'Dim attr As String
    Dim funcion As New CFuncion
    
    ReDim arr_code(0)
    C = 1
    
    For k = 1 To UBound(m_Code)
        get_line_tags m_Code(k)
        For j = 1 To UBound(arr_fun_info())
            Set funcion = New CFuncion
            funcion.FullName = arr_fun_info(j).Name
            funcion.nombre = ExtraeNombreFuncion(arr_fun_info(j).Name)
            funcion.Parametros = ExtraeParametros(arr_fun_info(j).Name)
            funcion.Tag = arr_fun_info(j).Tag
            m_Funciones.Add funcion, "k" & C
            Set funcion = Nothing
            C = C + 1
        Next j
    Next k

End Sub


Private Function ExtraeParametros(ByVal linea As String) As String

    Dim ret As String
    
    'buscar comienzo de parametro
    If InStr(1, linea, "(") Then
        linea = Mid$(linea, InStr(1, linea, "(") + 1)
        If InStr(1, linea, ")") Then
            linea = VBA.Left$(linea, InStr(1, linea, ")") - 1)
            ret = linea
        End If
    End If
    
    ExtraeParametros = ret
    
End Function

Private Function ExtraeNombreFuncion(ByVal linea As String) As String

    Dim ret As String
    'Dim pos As Integer
    
    If InStr(linea, "(") Then
        ret = VBA.Left$(linea, InStr(1, linea, "(") - 1)
    End If
    
    ExtraeNombreFuncion = ret
    
End Function


'solo obtener los tags correctos
Private Sub get_line_tags(ByVal linea As String)

    On Error GoTo Errorget_line_tags
    
    Dim k As Integer
    Dim j As Integer
    'Dim c As Integer
    Dim linea_aux As String
    Dim inicio As Integer
    Dim fin As Integer
    'Dim figual As Integer
    Dim Tag As String
    Dim sName As String
    'Dim fanonima As Boolean
    'Dim param As String
    Dim linea_izq As String
    Dim linea_der As String
    'Dim toleft As String
    'Dim toright As String
    Dim proto As Boolean
    
    'Dim fisfun As Boolean
    
    ReDim arr_fun_info(0)
    
    linea_aux = Trim$(util.SacarBasura(linea))
    
    If InStr(linea_aux, "var reverser") Then
        Debug.Print "stop1"
    End If
    
    'no procesar comentarios
    If (Left$(linea_aux, 4) = "<!--" Or Left$(linea_aux, 4) = "-->") Then
        Exit Sub
    ElseIf (Left$(linea_aux, 9) = LCase$("<comment>") Or Left$(LCase$(linea_aux), 10) = "</comment>") Then
        Exit Sub
    ElseIf Left$(linea_aux, 2) = "//" Then
        Exit Sub
    ElseIf (Left$(linea_aux, 2) = "/*" Or Left$(linea_aux, 2) = "*/") Then
        Exit Sub
    End If
    
    k = 1
    Do
        'verificar si hay algo relatado con alguna funcion
        If Not isfunction(linea_aux, inicio, proto) Then
            Exit Do
        End If
                
        'luego verificar que hay hacia la izquierda
        'debe haber un blanco o debe haber un = (funcion anonima)
        'hacia la derecha lo mismo blanco o parentesis derecho
        
        'extraer parametros ->derecha
        linea_izq = linea_aux
        linea_der = linea_aux
        
        'extraer nombre de la funcion. Aqui puedo determinar si la funcion
        'es con constructor o anonima
        If Not proto Then    'constructor
            'una funcion normal sin prototipo
            
            Tag = Mid$(linea_izq, inicio)
            
            'nombre de la funcion
            linea_der = Mid$(linea_der, inicio + 9)
            
            inicio = InStr(linea_der, "(")
            
            If inicio = 0 Then Exit Do
            
            fin = InStr(linea_der, ")")
            
            If fin = 0 Then Exit Do
            
            'nombre de la funcion con los parametros
            sName = Left$(linea_der, fin)
            
            'moverse desde el ultimo parentesis encontrado + 1 posicion
            linea_aux = Trim$(Mid$(linea_der, fin + 1))
        Else                                    'anonima
            'verificar hacia la derecha este bien construido
            inicio = InStr(LCase$(linea_izq), "(")
            
            If inicio = 0 Then Exit Do
            
            fin = InStr(LCase$(linea_izq), ")")
            
            If fin = 0 Then Exit Do
            
            'moverse desde el ultimo parentesis encontrado + 1 posicion
            linea_aux = Trim$(Mid$(linea_der, fin + 1))
            
            'todo okey entonces extraer desde el = hacia el 2 blanco hacia la izquierda
            inicio = InStr(LCase$(linea_izq), "=")
            If inicio = 0 Then Exit Do
            
            inicio = InStr(LCase$(linea_izq), ".")
            If inicio = 0 Then Exit Do
            
            'nombre de la funcion con los parametros para la busqueda
            Tag = Left$(linea_izq, fin)
            
            'ahora el nombre de la funcion anonima
            linea_izq = Mid$(Tag, InStr(1, Tag, ".") + 1)
            linea_der = linea_izq
            
            If InStr(1, linea_izq, "=") > 0 Then
                linea_izq = Left$(linea_izq, InStr(1, linea_izq, "=") - 1)
                If InStr(linea_izq, ".") Then
                    For j = Len(linea_izq) To 1 Step -1
                        If Mid$(linea_izq, j, 1) = "." Then
                            linea_izq = Mid$(linea_izq, j + 1)
                            Exit For
                        End If
                    Next j
                End If
                linea_der = Mid$(linea_der, InStr(1, linea_der, "("))
                sName = linea_izq & linea_der
            Else
                'nombre de la funcion
                inicio = InStr(LCase$(linea_der), "function")
                If inicio = 0 Then Exit Do
        
                linea_der = Mid$(linea_der, inicio + 9)
                
                inicio = InStr(linea_der, "(")
                
                If inicio = 0 Then Exit Do
                
                fin = InStr(linea_der, ")")
                
                If fin = 0 Then Exit Do
                
                'nombre de la funcion con los parametros
                sName = linea_der
                Tag = Tag
                
                'moverse desde el ultimo parentesis encontrado + 1 posicion
                linea_aux = Trim$(Mid$(linea_der, fin + 1))
            End If
        End If
               
        If Len(Trim$(sName)) = 0 Then Exit Do
        If sName = ")" Then Exit Do
        If sName = "()" Then Exit Do
        If Left$(sName, 1) = "(" Then Exit Do
        If InStr(sName, "+") Or InStr(sName, "[") Or InStr(sName, "]") Or InStr(sName, "*") Or InStr(sName, "?") Then
            Debug.Print sName
        Else
            If InStr(Tag, "(") Then
                Dim arr_blancos() As String
                Dim linea_paso As String
                
                linea_paso = LCase$(Tag)
                If InStr(linea_paso, "function ") Or InStr(linea_paso, "function(") Then
                    If InStr(linea_paso, "function ") Then
                        linea_paso = Mid$(linea_paso, InStr(linea_paso, "function "))
                        If InStr(linea_paso, "(") Then
                            linea_paso = Left$(linea_paso, InStr(1, linea_paso, "(") - 1)
                            If Right$(linea_paso, 1) = " " Then linea_paso = Left$(linea_paso, Len(linea_paso) - 1)
                            arr_blancos = Split(linea_paso, " ")
                        Else
                            Debug.Print linea_paso
                        End If
                    Else
                        ReDim arr_blancos(0)
                    End If
                    
                    If UBound(arr_blancos) <= 1 Then
                        ReDim Preserve arr_fun_info(k)
                        
                        If Left$(sName, 4) = "var " Then
                            sName = Mid$(sName, 5)
                        End If
                        
                        arr_fun_info(k).Name = sName
                        arr_fun_info(k).Tag = Tag
                        k = k + 1
                    Else
                        Debug.Print Tag
                    End If
                Else
                    Debug.Print Tag
                End If
            Else
                Debug.Print Tag
            End If
        End If
        
        'Debug.Print arr_fun_info(k).Name; "->"; arr_fun_info(k).tag
        
seguir:
    Loop
    
    Exit Sub
Errorget_line_tags:
    Debug.Print Err & " " & Error$
    
End Sub



Public Sub Clear()
    
    Dim k As Integer
    
    For k = m_Funciones.count To 1 Step -1
        m_Funciones.Remove k
    Next k
    
    ReDim arr_fun_info(0)
    ReDim m_Code(0)
    
End Sub



Public Property Set Funciones(ByVal pFunciones As Collection)
    Set m_Funciones = pFunciones
End Property


Public Property Get Funciones() As Collection
    Set Funciones = m_Funciones
End Property


Public Property Get filename() As String
    filename = m_FileName
End Property

Public Property Let filename(ByVal pFileName As String)
    m_FileName = pFileName
End Property


