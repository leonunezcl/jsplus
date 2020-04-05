VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmJsLint 
   Caption         =   "JSLINT"
   ClientHeight    =   8235
   ClientLeft      =   960
   ClientTop       =   4050
   ClientWidth     =   7200
   ControlBox      =   0   'False
   Icon            =   "frmJsLint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   7200
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   2250
      Left            =   750
      TabIndex        =   0
      Top             =   1815
      Width           =   4500
      ExtentX         =   7937
      ExtentY         =   3969
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmJsLint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Archivo As String
Private inifile As String
Private Function lee_valor_configuracion(ByVal llave As String) As String

    Dim ret As String
    Dim iniconfig As String
    
    iniconfig = util.StripPath(App.Path) & "jslint.ini"
    
    ret = IIf(util.LeeIni(iniconfig, "options", llave) <> "", util.LeeIni(iniconfig, "options", llave), 0)
    
    If ret <> "0" Then
        ret = "checked"
    Else
        ret = ""
    End If
    
    lee_valor_configuracion = ret
    
End Function

Private Sub Form_Load()
    
    Dim tmp As String
    Dim output_file As String
    Dim nFreeFile As Long
    Dim nfreefile2 As Long
    Dim tmpfile As String
    
    tmp = util.StripPath(App.Path) & "jslint\jslint.html"
    output_file = util.StripPath(App.Path) & "jslint\output.txt"
    tmpfile = util.StripPath(App.Path) & "jslint\tmpfile.js"
    
    nFreeFile = FreeFile
    
    inifile = util.StripPath(App.Path) & "jslint.ini"
    
    Open tmp For Output As #nFreeFile
        Print #nFreeFile, "<html>" & vbNewLine
        Print #nFreeFile, "<head>" & vbNewLine
        Print #nFreeFile, "<title></title>" & vbNewLine
        Print #nFreeFile, "<script src='jslint.js'></script>" & vbNewLine
        Print #nFreeFile, "</head>" & vbNewLine
        Print #nFreeFile, "<body onload='go();return false'>" & vbNewLine
        Print #nFreeFile, "<form name='jslint'>" & vbNewLine
        Print #nFreeFile, "<textarea name='input'>" & vbNewLine
    
        'solo una porcion de codigo o todo el archivo
        If ArchivoExiste2(tmpfile) Then
            nfreefile2 = FreeFile
            Open tmpfile For Input As #nfreefile2
                Print #nFreeFile, Input(LOF(nfreefile2), nfreefile2)
            Close #nfreefile2
        Else
            Print #nFreeFile, frmMain.ActiveForm.txtCode.Text
        End If
        Print #nFreeFile, "</textarea>" & vbNewLine
        Print #nFreeFile, "<div id=output></div>" & vbNewLine
        Print #nFreeFile, "<fieldset id=options>" & vbNewLine
        
        Dim k As Integer
        Dim arr_opciones() As String
        Dim valor As String
        
        inifile = util.StripPath(App.Path) & "analizer.ini"
    
        get_info_section "options", arr_opciones(), inifile
    
        For k = 1 To UBound(arr_opciones)
            valor = Explode(arr_opciones(k), 1, "|")
            Print #nFreeFile, "<input type=checkbox id=" & valor & " " & lee_valor_configuracion(valor) & ">" & vbNewLine
            Print #nFreeFile, "<br>" & vbNewLine
        Next k
        
        Print #nFreeFile, "</fieldset>" & vbNewLine
        Print #nFreeFile, "</form>" & vbNewLine
        Print #nFreeFile, "<script>" & vbNewLine
        Print #nFreeFile, "/*extern JSLINT */" & vbNewLine
        Print #nFreeFile, "var OPTIONS = [" & vbNewLine
        
        For k = 1 To UBound(arr_opciones)
            If k < UBound(arr_opciones) Then
                Print #nFreeFile, "'" & Explode(arr_opciones(k), 1, "|") & "'," & vbNewLine
            Else
                Print #nFreeFile, "'" & Explode(arr_opciones(k), 1, "|") & "'"
            End If
        Next k
        Print #nFreeFile, "];" & vbNewLine
        Print #nFreeFile, "function go(){" & vbNewLine
        Print #nFreeFile, "var b, d = new Date(), k = '{', i, o = {};" & vbNewLine
        Print #nFreeFile, "for (i = 0; i < OPTIONS.length; i += 1) {" & vbNewLine
        Print #nFreeFile, "if (document.forms.jslint[OPTIONS[i]].checked) {" & vbNewLine
        Print #nFreeFile, "o[OPTIONS[i]] = true;" & vbNewLine
        Print #nFreeFile, "if (b) {" & vbNewLine
        Print #nFreeFile, "k += ',';" & vbNewLine
        Print #nFreeFile, "}" & vbNewLine
        Print #nFreeFile, "k += '" & Chr$(34) & "' + OPTIONS[i] + '" & Chr$(34) & ":true';" & vbNewLine
        Print #nFreeFile, "b = true;" & vbNewLine
        Print #nFreeFile, "}" & vbNewLine
        Print #nFreeFile, "}" & vbNewLine
        Print #nFreeFile, "k += '}';" & vbNewLine
        Print #nFreeFile, "JSLINT(document.forms.jslint.input.value, o);" & vbNewLine
        Print #nFreeFile, "document.getElementById('output').innerHTML = JSLINT.report();" & vbNewLine
        Print #nFreeFile, "}" & vbNewLine
        Print #nFreeFile, "</script>" & vbNewLine
        Print #nFreeFile, "</body>" & vbNewLine
        Print #nFreeFile, "</html>" & vbNewLine
    Close #nFreeFile
    
    web1.Navigate tmp
    
    Do
       DoEvents
    Loop Until web1.ReadyState = READYSTATE_COMPLETE
                
    Dim webdoc As Object
            
    Set webdoc = web1.Document
        
    nFreeFile = FreeFile
    
    On Error Resume Next
    
    Open output_file For Output As #nFreeFile
        Print #nFreeFile, webdoc.Body.innerhtml
    Close #nFreeFile
    
    Set webdoc = Nothing
    
    Err = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Resize()
    
    If WindowState <> vbMinimized Then
        web1.Move 0, 0, ScaleWidth, ScaleHeight
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmJsLint = Nothing
End Sub


