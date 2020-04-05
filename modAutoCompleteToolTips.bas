Attribute VB_Name = "Module6"
Option Base 1
Option Explicit

Global selRange As CodeSenseCtl.IRange
'globals of CodeSense Control
Global CSGlobals As New CodeSenseCtl.Globals
'current word (for tooltip)
Global strCurrentWord As String
'current word function key in udtFuncDesc (see modFunctionDefinitions)
Global intCurrentWordItem As Integer
Global intCurrentWordItemJs As Integer

Private m_Object As String
Private m_Tipo As Integer
Public glbHideTip As Boolean
Public glbTipoTip As CodeSenseCtl.cmToolTipType
'CodeList ---
' TRIGGERED: When AutoComplete key shortcut is pressed
' PURPOSE: Shows list
' INPUTS:
'  Control - CodeSense control that caused this
'  ListCtrl - List control assigned to us by
'             the CodeSense control.
' RETURNS: True (to show list)

Function CodeList(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ImageList As MSComctlLib.ImageList) As Boolean

    'On Error Resume Next
    
    Dim intObjItem As Integer
    Dim intTemp As Integer
    Dim k As Integer
    Dim funcion As String
    Dim found As Boolean
    
    'set up list properties
'    ListCtrl.BackColor = Control.GetColor(cmClrWindow)
    ListCtrl.Font.Name = "Tahoma"
    ListCtrl.Font.Size = 10
    
    'image list for the little pictures to the side
    'on the list
    ListCtrl.ImageList = ImageList
    'ListCtrl.ImageList = ImageList

    'if current word is an object (ie. person typed
    ' a dot (.) after object name...
    If ObjDefined(Control.CurrentWord) Then
        '...look in autocomplete list for that object
        intObjItem = colObjList(Control.CurrentWord)
        'otherwise
    Else
        'use normal autocomplete list
        intObjItem = 0
    End If

    'go through all autocomplete items...
    m_Object = vbNullString
    m_Tipo = 0
    If intObjItem > 0 Then
        m_Object = Control.CurrentWord
        m_Tipo = 1
        For intTemp = 1 To UBound(udtObjInfo(intObjItem).strMembers)
            '...and add them to list
            
            If Len(Trim$(udtObjInfo(intObjItem).strMembers(intTemp))) > 0 Then
                If udtObjInfo(intObjItem).intMemberType(intTemp) = memProperty Then
                    ListCtrl.AddItem udtObjInfo(intObjItem).strMembers(intTemp), 2, 1
                ElseIf udtObjInfo(intObjItem).intMemberType(intTemp) = memFunction Then
                    ListCtrl.AddItem udtObjInfo(intObjItem).strMembers(intTemp), 1, 2
                ElseIf udtObjInfo(intObjItem).intMemberType(intTemp) = memEvent Then
                    ListCtrl.AddItem udtObjInfo(intObjItem).strMembers(intTemp), 15, 3
                ElseIf udtObjInfo(intObjItem).intMemberType(intTemp) = memConst Then
                    ListCtrl.AddItem udtObjInfo(intObjItem).strMembers(intTemp), 5, 4
                ElseIf udtObjInfo(intObjItem).intMemberType(intTemp) = memCollection Then
                    ListCtrl.AddItem udtObjInfo(intObjItem).strMembers(intTemp), 12, 5
                ElseIf udtObjInfo(intObjItem).intMemberType(intTemp) = memObject Then
                    ListCtrl.AddItem udtObjInfo(intObjItem).strMembers(intTemp), 14, 6
                Else
                    ListCtrl.AddItem udtObjInfo(intObjItem).strMembers(intTemp), 0
                End If
            End If
        Next intTemp
    Else
        'verificar los atributos para html
        Call frmMain.check_properties
        If Len(frmMain.ActiveForm.tagaux) > 0 Then
            For k = 1 To UBound(arr_html)
                If UCase$(arr_html(k).Tag) = UCase$(frmMain.ActiveForm.tagaux) Then
                    found = True
                    m_Object = frmMain.ActiveForm.tagaux
                    m_Tipo = 2
                    For intTemp = 1 To UBound(arr_html(k).elems)
                        If VBA.Left$(arr_html(k).elems(intTemp).attribute, 2) <> "on" Then
                            ListCtrl.AddItem arr_html(k).elems(intTemp).attribute, 2, 1
                        Else
                            If arr_html(k).elems(intTemp).icono = 1 Then
                                ListCtrl.AddItem arr_html(k).elems(intTemp).attribute, 15, 2
                            Else
                                ListCtrl.AddItem arr_html(k).elems(intTemp).attribute, 9, 2
                            End If
                        End If
                    Next intTemp
                    Exit For
                End If
            Next k
            
            'si el tag es invalido entonces
            If Not found Then
                'anexar los elementos de html
                For k = 1 To UBound(arr_html)
                    ListCtrl.AddItem arr_html(k).Tag, 8
                Next k
                
                'anexar los elementos de css
                For k = 1 To UBound(arr_data_css)
                    ListCtrl.AddItem arr_data_css(k).Tag, 10
                Next k
                
                'agregar los objetos del navegador
                For intTemp = 1 To UBound(udtObjetos)
                    If udtObjetos(intTemp) <> "Events" And _
                        udtObjetos(intTemp) <> "Global" And _
                        udtObjetos(intTemp) <> "Object" Then
                        ListCtrl.AddItem udtObjetos(intTemp), 0
                    End If
                Next intTemp
            
                'anexar las funciones del documento
                For intTemp = 1 To UBound(udtFuncDesc)
                    For k = 1 To UBound(udtFuncDesc(intTemp).strDef)
                        funcion = nombre_funcion(udtFuncDesc(intTemp).strDef(k))
                        ListCtrl.AddItem funcion, 1
                    Next k
                Next intTemp
            End If
        Else
            'anexar los elementos de html
            For k = 1 To UBound(arr_html)
                ListCtrl.AddItem arr_html(k).Tag, 8
            Next k
                            
            'anexar los elementos de css
            For k = 1 To UBound(arr_data_css)
                ListCtrl.AddItem arr_data_css(k).Tag, 10
            Next k
                
            'agregar los objetos del navegador
            For intTemp = 1 To UBound(udtObjetos)
                If udtObjetos(intTemp) <> "Eventos" And _
                   udtObjetos(intTemp) <> "Global" And _
                   udtObjetos(intTemp) <> "Object" Then
                   ListCtrl.AddItem udtObjetos(intTemp), 0
                End If
            Next intTemp
            
            'anexar las funciones del documento
            For intTemp = 1 To UBound(udtFuncDesc)
                For k = 1 To UBound(udtFuncDesc(intTemp).strDef)
                    funcion = nombre_funcion(udtFuncDesc(intTemp).strDef(k))
                    ListCtrl.AddItem funcion, 1
                Next k
            Next intTemp
        End If
    End If
            
'show list
CodeList = True
Err = 0
End Function

'CodeListSelMade ---
' TRIGGERED: When selection is made in list
' PURPOSE: To add item chosen to text
' INPUTS:
'  Control - CodeSense control that triggered this
'  ListCtrl - the list containing AutoComplete items
' RETURNS: False (to kill list)
Function CodeListSelMade(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
    
    Dim strItem As String
    Dim strItemAux As String
    Dim range As New CodeSenseCtl.range
    
    Dim tipomiembro As String
    Dim icono As Integer
    Dim tipo As Integer
    Dim help As String
    'Dim ctip As CodeSenseCtl.cmToolTipType
    Dim binsert As Boolean
    
    'get current word in text
    strItem = ListCtrl.GetItemText(ListCtrl.SelectedItem)
    strItemAux = strItem
    tipo = ListCtrl.GetItemData(ListCtrl.SelectedItem)
    binsert = True
    'if current word is start of word chosen in box
    '(ie. user entered 'Msg', went into AutoComplete,
    ' and chose 'MsgBox, it would replace 'Msg' with
    ' 'MsgBox' instead of inserting it. If it inserted
    ' it, the word would become 'MsgMsgbox')...
    If LCase$(Control.CurrentWord) = LCase$(VBA.Left$(strItem, Control.CurrentWordLength)) Then
        '...shorten text (if user typed 'Msg', went into
        ' autocomplete and chose MsgBox, this shortens
        ' the item to 'Box' ['Msg' + 'Box' = 'MsgBox'])
        strItem = Mid$(strItem, Control.CurrentWordLength + 1)
        
        'replace selection with this item
        Control.ReplaceSel (strItem)
    
        'get cursor position
        Set range = Control.GetSel(True)
        range.StartColNo = range.StartColNo + Len(strItem)
        range.EndColNo = range.StartColNo
        range.EndLineNo = range.StartLineNo
        'set cursor position to just after word
        Control.SetSel range, True
        binsert = False
    End If
    
    'determinar como se autocompleta
    If m_Tipo = 1 Then      'javascript
        help = frmMain.jsHlp.get_item_help(m_Object, strItemAux, tipo, tipomiembro, icono)
        
        If tipomiembro = "Property" Then
            'strItem = strItem & "="
        ElseIf tipomiembro = "Method" Then
            'strItem = strItem & "()"
        ElseIf tipomiembro = "Collection" Then
            strItem = strItem & "[]"
        ElseIf tipomiembro = "Object" Then
            strItem = strItem & "."
        ElseIf tipomiembro = "Constant" Then

        End If
                            
        If binsert Then
            frmMain.ActiveForm.Insertar strItem
        End If
        
        'setear el cursor a la posicion segun el tipo de elemento insertado
        'get cursor position
        Set range = Control.GetSel(True)
        
        range.StartColNo = range.StartColNo
        
        If tipomiembro = "Property" Then
            range.EndColNo = range.StartColNo
        ElseIf tipomiembro = "Method" Then
            range.EndColNo = range.StartColNo
        ElseIf tipomiembro = "Collection" Then
            range.EndColNo = range.StartColNo
        ElseIf tipomiembro = "Object" Then
            range.EndColNo = range.StartColNo
        End If
        
        range.EndLineNo = range.StartLineNo
        
        'llamar evento segun opcion
        If tipomiembro = "Method" Then
            glbTipoTip = CodeTip(Control, strItemAux)
            If glbTipoTip <> cmToolTipTypeNone Then
                frmMain.ActiveForm.Insertar "("
                glbHideTip = True
                Control.ExecuteCmd (cmCmdCodeTip)
            End If
        'ElseIf tipomiembro = "Object" Then
         '   CodeList Control, ListCtrl, frmMain.ActiveForm.imgAyuda
            'Control.ExecuteCmd cmCmdCodeList
        End If
    ElseIf m_Tipo = 2 Then  'html
        help = frmMain.MarkHlp.get_item_help(m_Object, strItem, tipo, tipomiembro, icono)
        
        If binsert Then
            If tipomiembro = "Property" Then
                strItem = strItem & "="
                frmMain.ActiveForm.Insertar strItem
            ElseIf tipomiembro = "Event" Then
                frmMain.ActiveForm.Insertar strItem
            End If
        Else
            If tipomiembro = "Property" Then
                strItem = "="
                frmMain.ActiveForm.Insertar strItem
            ElseIf tipomiembro = "Event" Then
                frmMain.ActiveForm.Insertar strItem
            End If
        End If
        
        Set range = Control.GetSel(True)
        
        range.StartColNo = range.StartColNo
        range.EndColNo = range.StartColNo
        range.EndLineNo = range.StartLineNo
    Else
        'verificar si es funcion de usuario la que se autocompleta
        If FuncDefined(strItemAux) Then
            glbTipoTip = CodeTip(Control, strItemAux)
            If glbTipoTip <> cmToolTipTypeNone Then
                frmMain.ActiveForm.Insertar "("
                glbHideTip = True
                Control.ExecuteCmd (cmCmdCodeTip)
            End If
        End If
    End If
    
    glbHideTip = False
    glbTipoTip = cmToolTipTypeNone
    
    'kill list
    CodeListSelMade = False
End Function
'CodeTip ---
' TRIGGERED: When ToolTip should be shown
' PURPOSE: To check if it really should be shown
' INPUTS:
'  Control - You should know by now ;)
' RETURNS: Type of tip to show (refer to documentation
'          on CodeSense control)
Function CodeTip(ByVal Control As CodeSenseCtl.ICodeSense, Optional ByVal TokenAux As String = vbNullString) As CodeSenseCtl.cmToolTipType

    Dim Token As CodeSenseCtl.cmTokenType
    
    'get current token type
    
    intCurrentWordItem = 0
    intCurrentWordItemJs = 0

    If Len(TokenAux) > 0 Then
        '...save current word
        strCurrentWord = TokenAux
        'if current word is defined...
        If FuncDefined(strCurrentWord) Then
            '...get the index of it...
            intCurrentWordItem = colFuncList(strCurrentWord)
            '...and tell codesource control to show tip.
            CodeTip = cmToolTipTypeMultiFunc
        ElseIf FuncDefinedJs(strCurrentWord) Then
            CodeTip = cmToolTipTypeMultiFunc
            'if word not defined...
        Else
            '...show no tip
            CodeTip = cmToolTipTypeNone
        End If
    Else
        Token = Control.CurrentToken
        'if current token is text or keyword...
        If ((Token = cmTokenTypeText) Or (Token = cmTokenTypeKeyword)) Then
            '...save current word
            strCurrentWord = Control.CurrentWord
            'if current word is defined...
            If FuncDefined(strCurrentWord) Then
                '...get the index of it...
                intCurrentWordItem = colFuncList(strCurrentWord)
                '...and tell codesource control to show tip.
                CodeTip = cmToolTipTypeMultiFunc
            ElseIf FuncDefinedJs(strCurrentWord) Then
                CodeTip = cmToolTipTypeMultiFunc
                'if word not defined...
            Else
                '...show no tip
                CodeTip = cmToolTipTypeNone
            End If
        'if in a comment...
        Else
            '...show no tip
            CodeTip = cmToolTipTypeNone
        End If
    End If
    
End Function

'OK, from now on I am not putting the 'Control'
'variable in the input description, as it is
'getting annoying to type it in everytime!
'And yes, instead of cutting and pasting the
'header on every function, i type it out again
'and again and again and again and again ;)


'CodeTipInitialize ---
' TRIGGERED: When ToolTip is initializing
' PURPOSE: To initialize the tooltio
' INPUTS:
'  ToolTipCtrl - the tooltip created by the control
' RETURNS: Nothing
Sub CodeTipInitialize(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ToolTipCtrl As CodeSenseCtl.ICodeTip)
Dim tip As CodeSenseCtl.CodeTipMultiFunc
'get the control
Set tip = ToolTipCtrl

'set current argument to the first one (this is zero-based)
tip.Argument = 0

'save position
Set selRange = Control.GetSel(True)
selRange.EndColNo = selRange.EndColNo '+ 1

If intCurrentWordItem > 0 Then
    'get definition count for the function
    tip.FunctionCount = UBound(udtFuncDesc(intCurrentWordItem).strDef) - 1
    'set current definition to the first one (again, 0-based)
    tip.CurrentFunction = 0
    'set tip to the first one
    tip.TipText = udtFuncDesc(intCurrentWordItem).strDef(1)
Else
    'get definition count for the function
    tip.FunctionCount = UBound(udtJsFuncs(intCurrentWordItemJs).strDef) - 1
    'set current definition to the first one (again, 0-based)
    tip.CurrentFunction = 0
    'set tip to the first one
    tip.TipText = udtJsFuncs(intCurrentWordItemJs).strDef(1)
End If
'set font
tip.Font.Name = "Tahoma"
tip.Font.Size = 10
tip.Font.Italic = True
tip.Font.Bold = False
End Sub

'CodeTipUpdate ---
' TRIGGERED: When tip should be updated
' PURPOSE: To update tip
' INPUTS:
'  ToolTipCtrl - current ToolTip
' RETURNS: Nothing
Sub CodeTipUpdate(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ToolTipCtrl As CodeSenseCtl.ICodeTip)
Dim iTrim As Integer, j As Integer
Dim bolInQuote As Boolean
Dim tip As CodeSenseCtl.CodeTipMultiFunc
'get tip
Set tip = ToolTipCtrl

Dim range As CodeSenseCtl.IRange
'get current cursor position
Set range = Control.GetSel(True)
'if user has moved up/down or before the statement...
If (range.EndLineNo <> selRange.EndLineNo) Or _
   (range.EndColNo < selRange.EndColNo) Then
   '...what do you think?
    tip.Destroy
Else
    Dim iArg, i As Integer
    Dim strLine As String

    iArg = 0
    'set i to the line number
    i = selRange.EndLineNo
    'get the line
    strLine = Control.GetLine(i)
    'set iTrim to length of line + 1
    iTrim = Len(strLine) + 1
    'if cursor isn't at end of line...
    If (range.EndColNo < iTrim) Then
        '...then iTrim = current cursor pos
        iTrim = range.EndColNo
    End If
    'get current line, up to iTrim
    strLine = VBA.Left(strLine, iTrim)
    bolInQuote = False
    j = 0
    'go through every character in line
    While ((Len(strLine) <> 0) And (j <= Len(strLine)) And (iArg <> -1))
        'check if quote encountered
        If (Mid(strLine, j + 1, 1) = """") Then
            bolInQuote = Not bolInQuote
        'if character is comma...
        ElseIf (Mid(strLine, j + 1, 1) = ",") And bolInQuote = False Then
            '...add 1 to argument count
            iArg = iArg + 1
        'if character is end bracket...
        ElseIf (Mid(strLine, j + 1, 1) = ")") And bolInQuote = False Then
            '...signal to destroy tip
            iArg = -1
        'if character is quote...
        ElseIf (Mid(strLine, j + 1, 1) = "'") And bolInQuote = False Then
            '...set iArg to -1 to destroy tip (since
            'user is starting a comment
            iArg = -1
        End If
        'add one to character count
        j = j + 1
    Wend
    'if tip should be destroyed...
    If (iArg = -1) Then
        '...destroy it, ...
        tip.Destroy
    '...otherwise
    Else
        'set number of current argument
        tip.Argument = iArg
        'set tiptext to current function description
        If intCurrentWordItem > 0 Then
            tip.TipText = udtFuncDesc(intCurrentWordItem).strDef(tip.CurrentFunction + 1)
        Else
            tip.TipText = udtJsFuncs(intCurrentWordItemJs).strDef(tip.CurrentFunction + 1)
        End If
    End If
End If
End Sub

'KeyPress ---
' TRIGGERED: When a key is pressed
' PURPOSE: To see if AutoComplete should be activated
' INPUTS:
'  KeyAscii - The ASCII code of the key
'  Shift - The KeyMask (eg. shift, alt or ctrl)
' RETURNS: Nothing
Function KeyPress(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyAscii As Long, ByVal Shift As Long) As Boolean
    Select Case KeyAscii
        'if key is starting bracket or space...
        Case (Asc("(")), (Asc(" ")), (Asc("<"))
            '...show AutoComplete
            Control.ExecuteCmd (cmCmdCodeTip)
        'if key is dot...
        Case (Asc("."))
            '...and current word is a defined object...
            If ObjDefined(Control.CurrentWord) Then
                '...show autocomplete
                Control.ExecuteCmd cmCmdCodeList
            End If
    End Select
    
End Function
Sub muestra_ayuda(ByVal ITem As String, ByVal tipo As Integer)
    
    Dim help As String
    Dim tipomiembro As String
    Dim icono As Integer
    
    If m_Tipo = 1 Then
        help = frmMain.jsHlp.get_item_help(m_Object, ITem, tipo, tipomiembro, icono)
    ElseIf m_Tipo = 2 Then
        help = frmMain.MarkHlp.get_item_help(m_Object, ITem, tipo, tipomiembro, icono)
    Else
        Exit Sub
    End If
            
    If glbquickon Then
        frmQuickTip.framex.Caption = ITem
        Set frmQuickTip.framex.Picture = LoadResPicture(icono, vbResIcon)
        frmQuickTip.lbltype.Caption = tipomiembro
        frmQuickTip.lblhelp.Caption = help
        frmQuickTip.Refresh
    End If
    
    DoEvents
    
End Sub


