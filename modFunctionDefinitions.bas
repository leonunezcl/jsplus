Attribute VB_Name = "Module5"
'Option Base 1
Option Explicit

'Type used for AutoComplete
Type ObjectDescription
    strMembers() As String
    intMemberType() As MemberTypes
End Type

'The different member types of an object -
'Const, Enum, Function and Property (variable)
Enum MemberTypes
    memConst
    memEvent
    memFunction
    memProperty
    memCollection
    memObject
End Enum

'Type used to hold function definitions (the little
'tooltip that pops up)

Type FunctionDescription
    strDef() As String
End Type

'Function List and Function Description variables
'Function list holds key numbers for descriptions (a lookup table),
'eg. if colFuncList("MsgBox") = 4, then
' udtFuncDesc(4).strDef(1) is the first definition
'for MsgBox(key no. 4). Understand?
Global colFuncList As New Collection
Global udtFuncDesc() As FunctionDescription

'Object List and Object info vars
Global colObjList As New Collection
Public udtObjInfo() As ObjectDescription
Public udtObjetos() As String
Public udtJsFuncs() As FunctionDescription

'Dim intOldCount

Public Function nombre_funcion(ByVal funcion As String) As String

    Dim ret As String
    
    If InStr(funcion, "(") Then
        ret = VBA.Left$(funcion, InStr(1, funcion, "(") - 1)
    Else
        ret = funcion
    End If
    
    nombre_funcion = ret
    
End Function
Function FuncDefinedJs(ByVal strFunc As String) As Boolean
    
    Dim k As Integer
    Dim j As Integer
    Dim funcion As String
    Dim ret As Boolean
    'Dim found As Boolean
    
    intCurrentWordItemJs = 0
    
    strFunc = LCase$(strFunc)
    For k = 1 To UBound(udtJsFuncs)
        For j = 1 To UBound(udtJsFuncs(k).strDef)
            funcion = nombre_funcion(LCase$(udtJsFuncs(k).strDef(j)))
            'If InStr(funcion, "format") Then
            '    Debug.Print "stop!"
            'End If
            If funcion = strFunc Then
                intCurrentWordItemJs = k
                ret = True
                GoTo salir
            End If
        Next j
    Next k
    
salir:
    FuncDefinedJs = ret

End Function

'AddFunc ---
' PURPOSE: Add a function into definition array
' INPUTS:
'  strFuncName - Function Name
'  ParamArray strFuncDefs - Function parameters
' RETURNS: New index number
' EXAMPLE: AddFunc("test", "test1", "foo, bar")
Function AddFunc(strFuncName As String, bolAddToAutoComplete As Boolean, ByVal strFuncDefs As String) As Integer

    On Error GoTo errorfuncdup

    Dim intNewCount As Integer
    Dim intTemp As Integer
    Dim MyStrFuncDefs()
    Dim C As Integer
    
    If Len(strFuncName) = 0 Then
        Exit Function
    End If

    intNewCount = colFuncList.count + 1
    colFuncList.Add intNewCount, Trim$(strFuncName)

    'if we have to add to autocomplete list
    'If bolAddToAutoComplete = True Then
    '    intTemp = UBound(udtObjInfo(1).strMembers) + 1
    '    ReDim Preserve udtObjInfo(1).strMembers(intTemp)
    '    ReDim Preserve udtObjInfo(1).intMemberType(intTemp)
    '    udtObjInfo(1).strMembers(intOldCount + intNewCount) = strFuncName
    '    udtObjInfo(1).intMemberType(intOldCount + intNewCount) = memFunction
    'End If

    'contar parametros
    C = 1
    Do
        If Len(util.Explode(strFuncDefs, C + 1, "|")) = 0 Then
            Exit Do
        End If
        C = C + 1
    Loop
    
    ReDim MyStrFuncDefs(0)
    ReDim MyStrFuncDefs(C)
    
    For C = 1 To UBound(MyStrFuncDefs)
        MyStrFuncDefs(C) = util.Explode(strFuncDefs, C, "|")
    Next C
    
    'resize definition array to hold the number of
    'definitions passed in
    'ReDim udtFuncDesc(intNewCount)
    'ReDim udtFuncDesc(intNewCount).strDef(0)
    ReDim udtFuncDesc(intNewCount).strDef(UBound(MyStrFuncDefs))
    For intTemp = LBound(MyStrFuncDefs) To UBound(MyStrFuncDefs)
        udtFuncDesc(intNewCount).strDef(intTemp) = strFuncName & "(" & MyStrFuncDefs(intTemp) & ")"
    Next intTemp

    'return new index
    AddFunc = intNewCount
Exit Function
errorfuncdup:
    'MsgBox "AddFunc : " & Err & " " & Error$, vbCritical
    Err = 0
End Function

Public Sub AgregaFuncionJs(ByVal funcion As String)

    Dim k As Integer
    Dim j As Integer
    Dim funori As String
    Dim funtemp As String
    Dim parori As String
    Dim parfun As String
    Dim found As Boolean
    
    funtemp = nombre_funcion(funcion)
    
    For k = 1 To UBound(udtJsFuncs)
        funori = nombre_funcion(udtJsFuncs(k).strDef(1))
        If funori = funtemp Then
            parfun = Mid$(funcion, InStr(1, funcion, "(") + 1)
            If parfun = ")" Then
                parfun = vbNullString
            Else
                parfun = VBA.Left$(parfun, InStr(1, parfun, ")") - 1)
            End If
            
            For j = 1 To UBound(udtJsFuncs(k).strDef)
                parori = Mid$(udtJsFuncs(k).strDef(j), InStr(1, udtJsFuncs(k).strDef(j), "(") + 1)
                If parori = ")" Then
                    parori = vbNullString
                Else
                    parori = VBA.Left$(parori, InStr(1, parori, ")") - 1)
                End If
                
                If parori = parfun Then
                    found = True
                    Exit For
                End If
            Next j
            
            If Not found Then
                j = UBound(udtJsFuncs(k).strDef) + 1
                ReDim Preserve udtJsFuncs(k).strDef(j)
                udtJsFuncs(k).strDef(j) = funcion
                Exit Sub
            End If
        End If
    Next k
    
    k = UBound(udtJsFuncs()) + 1
    
    ReDim Preserve udtJsFuncs(k)
    ReDim Preserve udtJsFuncs(k).strDef(1)
    udtJsFuncs(k).strDef(1) = funcion
            
End Sub

'FuncDefined ---
' PURPOSE: Find out if a function is defined
' INPUTS:
'  strFunc - The Function name
' RETURNS: Boolean stating whether function is defined
' EXAMPLE: bolTemp = FuncDefined("MsgBox")
'
' I won't comment this as anyone should understand
' how it works ;)
Function FuncDefined(ByVal strFunc As String) As Boolean

    Dim intTemp As Integer
    Dim k As Integer
    Dim funcion As String
    
    On Error GoTo nofunc

    intTemp = colFuncList(strFunc)
    FuncDefined = True
    Exit Function
    Exit Function
nofunc:
FuncDefined = False
End Function

'AddObject (similar to AddFunc)---
' PURPOSE: Add Object into AutoComplete array
' INPUTS:
'  strObjName - Name of object
'  ParamArray strObjMembers - Members of this object,
'   prefixed by P(roperty), F(unction), C(onst), or E(num)
' RETURNS: New index number
' EXAMPLE: AddObject("testobj", "Ftestfunction", "Etestenum")
Function AddObject(strObjName As String, strObjMembers As Collection) As Integer
Dim intNewCount As Integer
Dim intTemp As Integer

If strObjMembers.count = 0 Then
    Exit Function
End If

'Find new index
intNewCount = colObjList.count + 1
'Add object to lookup table
colObjList.Add intNewCount, strObjName

'resize array of members
ReDim Preserve udtObjInfo(intNewCount)
ReDim Preserve udtObjInfo(intNewCount).strMembers(strObjMembers.count)
'resize array of members' type
ReDim Preserve udtObjInfo(intNewCount).intMemberType(strObjMembers.count)
'loop through all the members passed in...
For intTemp = 1 To strObjMembers.count
    '...find the member type...
    Select Case VBA.Left(strObjMembers.ITem(intTemp), 1)
        Case "P":
            udtObjInfo(intNewCount).intMemberType(intTemp) = memProperty
        Case "F":
            udtObjInfo(intNewCount).intMemberType(intTemp) = memFunction
        Case "C":
            udtObjInfo(intNewCount).intMemberType(intTemp) = memConst
        Case "E":
            udtObjInfo(intNewCount).intMemberType(intTemp) = memEvent
        Case "X":
            udtObjInfo(intNewCount).intMemberType(intTemp) = memCollection
        Case "O":
            udtObjInfo(intNewCount).intMemberType(intTemp) = memObject
        Case Else:
            udtObjInfo(intNewCount).intMemberType(intTemp) = memFunction
    End Select
    '...and add member to array
    udtObjInfo(intNewCount).strMembers(intTemp) = Mid$(strObjMembers.ITem(intTemp), 2)
    'If Len(Trim$(udtObjInfo(intNewCount).strMembers(intTemp))) = 0 Then
        'Debug.Print "stop!"
    'End If
'continue loop
Next intTemp

'return new index
AddObject = intNewCount
End Function

'ObjDefined (similar to FuncDefined)---
' PURPOSE: Find out if object defined
' INPUTS:
'  strObject - Object to check
' RETURNS: True if object is defined
' EXAMPLE: bolTemp = ObjDefined("testobj")
'As with FuncDefined, i won't comment this.
Function ObjDefined(strObject As String) As Boolean
Dim strTemp As String
On Error GoTo noobj
strTemp = colObjList(strObject)
ObjDefined = True
Exit Function

noobj:
ObjDefined = False
End Function
'FuncString ---
' returns a string with all the functions in it,
' seperated by a VbLf (linefeed)
Function FuncString(ByVal seccion As String) As String

    Dim strTemp As String
    Dim intTemp As Integer
    'Dim num As Integer
    Dim ini As String
    Dim elem As String
    Dim glosa As String
    
    Dim sSections() As String
    ini = util.StripPath(App.Path) & "config\jshelp.ini"
    get_info_section seccion, sSections, ini
        
    For intTemp = 2 To UBound(sSections)
        glosa = util.Explode(sSections(intTemp), 1, "#")
        If InStr(glosa, "(") Then
            glosa = VBA.Left$(glosa, InStr(1, glosa, "(") - 1)
        End If
        'elem = Replace(glosa, "(", "")
        'elem = Replace(elem, ")", "")
        elem = Replace(glosa, " ", "")
        elem = Replace(elem, ",", "")
        elem = Replace(elem, ".", "")
        elem = Replace(elem, "[", "")
        elem = Replace(elem, "]", "")
        elem = Replace(elem, "<", "")
        elem = Replace(elem, ">", "")
        elem = Replace(elem, "/", "")
        strTemp = strTemp & elem & vbLf
    Next intTemp
        
    FuncString = Trim$(strTemp)
    
End Function
