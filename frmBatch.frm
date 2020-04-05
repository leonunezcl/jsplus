VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBatch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Batch Mode"
   ClientHeight    =   4155
   ClientLeft      =   3900
   ClientTop       =   1800
   ClientWidth     =   8415
   ControlBox      =   0   'False
   Icon            =   "frmBatch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   13
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Clear List"
      Height          =   375
      Index           =   3
      Left            =   6960
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Add Folder"
      Height          =   375
      Index           =   5
      Left            =   6960
      TabIndex        =   11
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Add File"
      Height          =   375
      Index           =   4
      Left            =   6960
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Analyzer Options"
      Height          =   375
      Index           =   2
      Left            =   6960
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "File Extensions"
      Height          =   375
      Index           =   6
      Left            =   6960
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Analyze Files"
      Height          =   375
      Index           =   0
      Left            =   6960
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.CheckBox chk 
      Caption         =   "Select All"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   3855
      Width           =   1035
   End
   Begin VB.Frame fra 
      Caption         =   "Selected Files"
      Height          =   3705
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6690
      Begin MSComctlLib.ListView lvwFiles 
         Height          =   3195
         Left            =   90
         TabIndex        =   1
         Top             =   210
         Width           =   6540
         _ExtentX        =   11536
         _ExtentY        =   5636
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   3435
         Width           =   45
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "XXX"
         Height          =   195
         Left            =   6270
         TabIndex        =   3
         Top             =   3450
         Visible         =   0   'False
         Width           =   315
      End
   End
   Begin CodeSenseCtl.CodeSense txtCode 
      Height          =   1095
      Left            =   1830
      OleObjectBlob   =   "frmBatch.frx":000C
      TabIndex        =   2
      Top             =   5760
      Width           =   1620
   End
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   765
      Left            =   1725
      TabIndex        =   5
      Top             =   4980
      Width           =   2280
      ExtentX         =   4022
      ExtentY         =   1349
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
Attribute VB_Name = "frmBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private contador As Integer
Private ArchivoReporte As String

Private Type eInfoAna
   Line As Integer
   Col As Integer
   sError As String
   Expresion As String
End Type

Private arr_report() As eInfoAna

Private IniFileTree As String
Private arr_ext() As String
Private Sub AddFile()

    Dim Archivo As String
    Dim glosa As String
    
    glosa = strGlosa()
    
    If LastPath = "" Then LastPath = App.Path
    
    If Not Cdlg.VBGetOpenFileName(Archivo, , , , , , glosa, , LastPath, "Select Files ...", "js", Me.hwnd) Then
        Exit Sub
    End If
        
    If ExtensionAllowed(Archivo) Then
      lvwFiles.ListItems.Add , "k" & contador, Archivo
      lvwFiles.ListItems("k" & contador).Checked = True
      contador = contador + 1
    End If
    
End Sub
Private Sub AddFolder()

    Dim Path As String
    Dim arr_archivos() As String
    Dim k As Integer
    
    Path = util.BrowseFolder(hwnd)
    
    If Len(Path) > 0 Then
        get_files_from_folder Path, arr_archivos
        
        For k = 1 To UBound(arr_archivos)
            If ExtensionAllowed(arr_archivos(k)) Then
                lvwFiles.ListItems.Add , "k" & contador, arr_archivos(k)
                lvwFiles.ListItems("k" & contador).Tag = arr_archivos(k)
                lvwFiles.ListItems("k" & contador).Checked = True
                contador = contador + 1
            End If
        Next k
    End If
    
End Sub


Private Function AnalizeSelectedFile() As Boolean

    On Error GoTo AnalizeSelectedFile
    
    Dim tmp As String
    Dim output_file As String
    Dim nFreeFile As Long
    Dim nfreefile2 As Long
    Dim tmpfile As String
    Dim inifile As String
    
    util.Hourglass hwnd, False
    
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
        nfreefile2 = FreeFile
        Open tmpfile For Input As #nfreefile2
            Print #nFreeFile, Input(LOF(nfreefile2), nfreefile2)
        Close #nfreefile2
        
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
    
    If ArchivoExiste2(tmp) Then
        web1.Navigate tmp
        
        Do
           On Error Resume Next
           DoEvents
           Err = 0
        Loop Until web1.ReadyState = READYSTATE_COMPLETE
                    
        Dim webdoc As Object
                
        Set webdoc = web1.Document
            
        nFreeFile = FreeFile
        
        Open output_file For Output As #nFreeFile
            Print #nFreeFile, webdoc.Body.innerhtml
        Close #nFreeFile
        
        Set webdoc = Nothing
        
        AnalizeSelectedFile = True
    Else
        AnalizeSelectedFile = False
    End If
    
    util.Hourglass hwnd, False
    
    Exit Function
    
AnalizeSelectedFile:
    util.Hourglass hwnd, False
    DoEvents
    AnalizeSelectedFile = False
    
End Function
Private Function ExtensionAllowed(ByVal filename As String) As Boolean

   Dim k As Integer
   Dim ext As String
   Dim ret As Boolean
   
   ext = GetFileExtension(filename)
   
   For k = 1 To UBound(arr_ext)
      If LCase$(ext) = LCase$(arr_ext(k)) Then
         ret = True
         Exit For
      End If
   Next k
   
   ExtensionAllowed = ret
   
End Function

Private Function lee_valor_configuracion(ByVal llave As String) As String

    Dim ret As String
    Dim iniconfig As String
    
    iniconfig = util.StripPath(App.Path) & "jslint.ini"
    
    ret = IIf(util.LeeIni(iniconfig, "options", llave) <> "", util.LeeIni(iniconfig, "options", llave), 0)
    
    If ret <> "0" Then
        ret = "checked"
    Else
        ret = vbNullString
    End If
    
    lee_valor_configuracion = ret
    
End Function

Private Sub AnalyzeFiles()

    On Error GoTo ErrorAnalyzeFiles
    
    Dim Archivo As String
    Dim Fuente As String
    Dim C As Integer
    Dim fret As Boolean
    
    util.Hourglass hwnd, True
    For C = 1 To lvwFiles.ListItems.count
      If lvwFiles.ListItems(C).Checked Then
         fret = True
         Exit For
      End If
    Next C
    util.Hourglass hwnd, False
    
    If Not fret Then
      MsgBox "You must select files to analyze.", vbCritical
      Exit Sub
    End If
    
    Archivo = util.StripPath(App.Path) & "jslint\jslint.js"
    
    If Not ArchivoExiste2(Archivo) Then
        MsgBox "File doesn't found : " & Archivo, vbAbortRetryIgnore
        Exit Sub
    End If
        
    Archivo = util.StripPath(App.Path) & "jslint.ini"
    
    If Not ArchivoExiste2(Archivo) Then
        MsgBox "You must first configure the JavaScript Code Analizer.", vbInformation
        frmjslitopt.Show vbModal
        Exit Sub
    End If
    
    Dim k As Integer
    Dim j As Integer
    Dim tmpfile As String
    Dim nFreeFile As Long
    Dim nfreefile2 As Long
    Dim ret As String
    
    ret = "HTML Files (*.html)|*.html|"
    ret = ret & "All Files (*.*)|*.*"
            
    If LastPath = "" Then LastPath = App.Path
    
    ArchivoReporte = vbNullString
    
    Call Cdlg.VBGetSaveFileName(ArchivoReporte, , , ret, , LastPath, "Save Report As ...", "html", Me.hwnd)
    
    If Len(ArchivoReporte) = 0 Then Exit Sub
    
    LastPath = PathArchivo(Archivo)
    
    EnabledButtons False
    
    tmpfile = util.StripPath(App.Path) & "jslint\tmpfile.js"
            
    util.Hourglass hwnd, True
    
    nFreeFile = FreeFile
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    lblFile.Caption = "Preparing ...Please Wait ..."
    
    For k = 1 To lvwFiles.ListItems.count
        lvwFiles.ListItems(k).SubItems(1) = vbNullString
    Next k
            
    Open ArchivoReporte For Output As #nFreeFile
        'cabezera del archivo
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>JavaScript Plus Report</title>"
        Print #nFreeFile, Replace("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>", "'", Chr$(34))
        Print #nFreeFile, "</head>"
        
        'titulo del reporte
        Print #nFreeFile, Replace("<body bgcolor='#FFFFFF' text='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Report Date : " & Now & " </b></p>"
        
        lblFile.Caption = vbNullString
        lblFile.Visible = True
        lblInfo.Caption = vbNullString
        lblInfo.Visible = True
        
        DoEvents
        
        C = 0
        
        'analizar los archivos seleccionados
        For k = 1 To lvwFiles.ListItems.count
            
            If lvwFiles.ListItems(k).Checked = True Then
                util.BorrarArchivo tmpfile
        
                nfreefile2 = FreeFile
                txtCode.OpenFile lvwFiles.ListItems(k).Text
                Open tmpfile For Output As #nfreefile2
                    Print #nfreefile2, txtCode.Text
                Close #nfreefile2
        
                lblFile.Caption = lvwFiles.ListItems(k).Text
                
                lvwFiles.ListItems(k).SubItems(1) = "Opening"
                
                If Err = 0 Then
                                                
                    lblInfo.Caption = CStr(k) & " of " & lvwFiles.ListItems.count & " files."
                    
                    lvwFiles.ListItems(k).SubItems(1) = "Analyzing"
                    
                    Print #nFreeFile, "<p><b>File : " & util.VBArchivoSinPath(lvwFiles.ListItems(k).Text) & "</b></p>"
                    Print #nFreeFile, "</font>"
                    
                    'generar titulos
                    Print #nFreeFile, Replace("<table width='97%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
                    Print #nFreeFile, Replace("<tr bgcolor='#999999' bordercolor='#000000'>", "'", Chr$(34))
                    
                    Print #nFreeFile, Replace("<td width='03%'><b>" & Fuente & "Number</font></b></td>", "'", Chr$(34))
                    Print #nFreeFile, Replace("<td width='03%'><b>" & Fuente & "Line</font></b></td>", "'", Chr$(34))
                    Print #nFreeFile, Replace("<td width='03%'><b>" & Fuente & "Column</font></b></td>", "'", Chr$(34))
                    Print #nFreeFile, Replace("<td width='25%'><b>" & Fuente & "Error</font></b></td>", "'", Chr$(34))
                    Print #nFreeFile, Replace("<td width='25%'><b>" & Fuente & "Expression</font></b></td>", "'", Chr$(34))
                    
                    Print #nFreeFile, "</tr>"
                
                    If AnalizeSelectedFile() Then
                        If LoadJsLintFile() Then
                            For j = 1 To UBound(arr_report) 'lvwJSlintMsg.ListItems.count
                                'Set Itmx = lvwJSlintMsg.ListItems(j)
                            
                                'imprimir informacion
                                Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
                            
                                'correlativo
                                Print #nFreeFile, Replace("<td width='03%' height='18'>" & Fuente & k & "</font></td>", "'", Chr$(34))
                                            
                                'linea
                                Print #nFreeFile, Replace("<td width='03%' height='18'><b>" & Fuente & arr_report(j).Line & "</font></b></td>", "'", Chr$(34))
                            
                                'columna
                                Print #nFreeFile, Replace("<td width='03%' height='18'>" & Fuente & arr_report(j).Col & "</font></td>", "'", Chr$(34))
                                        
                                'error
                                Print #nFreeFile, Replace("<td width='25%' height='18'><b>" & Fuente & arr_report(j).sError & "</font></b></td>", "'", Chr$(34))
                                
                                'expression
                                Print #nFreeFile, Replace("<td width='25%' height='18'><b>" & Fuente & arr_report(j).Expresion & "</font></b></td>", "'", Chr$(34))
                                Print #nFreeFile, "</tr>"
                            Next j
                        
                            Print #nFreeFile, "</table>"
                            Print #nFreeFile, "<br>"
                            lvwFiles.ListItems(k).SubItems(1) = "Analyzed"
                            
                            C = C + 1
                        Else
                            lvwFiles.ListItems(k).SubItems(1) = "Error"
                        End If
                    Else
                        lvwFiles.ListItems(k).SubItems(1) = "Failed"
                    End If
                Else
                    lvwFiles.ListItems(k).SubItems(1) = "Failed"
                End If
            Else
                lvwFiles.ListItems(k).SubItems(1) = "Not Analyzed"
            End If
        Next k
        
        util.BorrarArchivo tmpfile
        
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
    Close #nFreeFile
    
    lblFile.Visible = False
    lblInfo.Visible = False
    EnabledButtons True
    
    util.Hourglass hwnd, False
    
    MsgBox CStr(C) & " files analyzed." & vbNewLine & vbNewLine & "File report : " & ArchivoReporte, vbInformation
    
    Unload Me
    
    Exit Sub
    
ErrorAnalyzeFiles:
    Close #nFreeFile
    MsgBox "AnalyzeFiles :" & Err & " " & Error$, vbCritical
    EnabledButtons True
    
End Sub

Private Sub EnabledButtons(ByVal ret As Boolean)

    Dim k As Integer
    
    For k = 0 To cmd.count - 1
      cmd(k).Enabled = ret
      cmd(k).Refresh
    Next k
    
End Sub

Private Sub SelectFiles()

    Dim Archivos As String
    Dim glosa As String
    Dim arr_archivos() As String
    Dim k As Integer
    
    glosa = strGlosa()
    
    If Not Cdlg.VBGetOpenFileName(Archivos, , , True, , , glosa, , LastPath, "Select Files ...", "js", Me.hwnd) Then
        Exit Sub
    End If
        
    arr_archivos = Split(Archivos, Chr$(32))
    
    For k = 1 To UBound(arr_archivos)
        lvwFiles.ListItems.Add , "k" & contador, arr_archivos(k)
        lvwFiles.ListItems("k" & contador).Tag = arr_archivos(k)
        lvwFiles.ListItems("k" & contador).Selected = True
        contador = contador + 1
    Next k
    
End Sub

Private Sub chk_Click()

   Dim ret As Boolean
   Dim k As Integer
   
   If chk.Value = 1 Then
      ret = True
   End If
   
   util.Hourglass hwnd, True
   
   For k = 1 To lvwFiles.ListItems.count
      lvwFiles.ListItems(k).Checked = ret
   Next k
   
   util.Hourglass hwnd, False
   
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case 0  'analyze
            If lvwFiles.ListItems.count > 0 Then
                Call AnalyzeFiles
            End If
        Case 6 'extensions
            frmSetTreeExp.origen = 1
            frmSetTreeExp.Show vbModal
            get_info_section "tree", arr_ext, IniFileTree
        Case 2 'options
            frmjslitopt.Show vbModal
        Case 4  'add file
            Call AddFile
        Case 5  'add folder
            Call AddFolder
        Case 1  'cancel
            Unload Me
        Case 2  'file
            Call SelectFiles
        Case 3  'clear list
            lvwFiles.ListItems.Clear
    End Select
    
End Sub
Private Sub Form_Load()

   util.CenterForm Me
   util.Hourglass hwnd, True
    
   IniFileTree = StripPath(App.Path) & "batch.ini"
   
   get_info_section "tree", arr_ext, IniFileTree
   
   If UBound(arr_ext) = 0 Then
      Call util.GrabaIni(IniFileTree, "tree", "ext1", "js")
      get_info_section "tree", arr_ext, IniFileTree
   End If
      
   util.Hourglass hwnd, False
    
End Sub
Private Function LoadJsLintFile() As Boolean

    On Error GoTo ErrorLoadJsLintFile
    
    Dim output_file As String
    Dim linea As String
    Dim nFreeFile As Long
    Dim nlinea As Integer
    Dim Col As Integer
    Dim nlinea1 As Integer
    Dim col1 As Integer
    Dim analizando As Boolean
    Dim C As Integer
    
    output_file = util.StripPath(App.Path) & "jslint\output.txt"

    ReDim arr_report(0)
    
    C = 1
    
    If ArchivoExiste2(output_file) Then
        'abrir el archivo de salida y analizarlo
        nFreeFile = FreeFile
        Open output_file For Input As #nFreeFile
          Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, linea
            If Left$(LCase$(linea), 15) = "<div id=output>" Then
                analizando = True
            End If
            
            If analizando Then
                If LCase$(Left$(linea, 7)) = "<table>" Then
                    Exit Do
                End If
                
                If LCase$(linea) <> "<div id=output>" Then
                    If InStr(linea, "</P>") Then
                        linea = Trim$(Mid$(linea, 4))
                        linea = Trim$(Left$(linea, InStr(linea, "</P>") - 1))
                                            
                        If InStr(linea, "|") Then
                           If C = 1 Then
                               nlinea1 = util.Explode(linea, 1, "|")
                               col1 = util.Explode(linea, 2, "|")
                           End If
                           
                           nlinea = util.Explode(linea, 1, "|")
                           
                           Col = util.Explode(linea, 2, "|")
                           
                           ReDim Preserve arr_report(C)
                           
                           arr_report(C).Line = nlinea
                           arr_report(C).Col = Col
                           arr_report(C).sError = util.Explode(linea, 3, "|")
                           arr_report(C).Expresion = util.Explode(linea, 4, "|")
                           
                           C = C + 1
                        End If
                    Else
                        Exit Do
                    End If
                End If
            End If
          Loop
        Close #nFreeFile
    End If
    
    LoadJsLintFile = True
    
    Exit Function
    
ErrorLoadJsLintFile:
    LoadJsLintFile = False
    Close #nFreeFile
    
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBatch = Nothing
End Sub


