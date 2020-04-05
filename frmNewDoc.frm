VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmNewDoc 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New File"
   ClientHeight    =   5520
   ClientLeft      =   3495
   ClientTop       =   3375
   ClientWidth     =   7110
   Icon            =   "frmNewDoc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   474
   Begin MSComctlLib.ListView lvwTab 
      Height          =   4335
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   7646
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   4815
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "JavaScript"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "HTML"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "CSS"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Plus!"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Samples for Learners"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "AJAX Libraries"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9960
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
End
Attribute VB_Name = "frmNewDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tabactivo As Integer
Private fcargando As Boolean
Private WithEvents m_cUnzip As cUnzip
Attribute m_cUnzip.VB_VarHelpID = -1
Private Sub carga_template_ajax()

    Dim Archivo As String
    Dim ArchDest As String
    
    If Not lvwTab(5).SelectedItem Is Nothing Then
        Me.Hide
        Select Case lvwTab(5).SelectedItem.Index
            Case 1  'aflax
                Archivo = util.StripPath(App.Path) & "libraries\aflax\aflax_sample.template"
                ArchDest = util.StripPath(App.Path) & "libraries\aflax\aflax_sample.htm"
            Case 2  'dojo
                Archivo = util.StripPath(App.Path) & "libraries\dojo\dojo_sample.template"
                ArchDest = util.StripPath(App.Path) & "libraries\dojo\dojo_sample.htm"
            Case 3  'jquery
                Archivo = util.StripPath(App.Path) & "libraries\jquery\jquery_sample.template"
                ArchDest = util.StripPath(App.Path) & "libraries\jquery\jquery_sample.htm"
            Case 4  'google
                Archivo = util.StripPath(App.Path) & "libraries\google\Hello.template"
                ArchDest = util.StripPath(App.Path) & "libraries\google\Hello.java"
            Case 5  'mochikit
                Archivo = util.StripPath(App.Path) & "libraries\mochikit\mochikit_sample.template"
                ArchDest = util.StripPath(App.Path) & "libraries\mochikit\mochikit_sample.htm"
            Case 6  'prototype
                Archivo = util.StripPath(App.Path) & "libraries\prototype\prototype_sample.template"
                ArchDest = util.StripPath(App.Path) & "libraries\prototype\prototype_sample.htm"
            Case 7  'rico
                Archivo = util.StripPath(App.Path) & "libraries\rico\rico_sample.template"
                ArchDest = util.StripPath(App.Path) & "libraries\rico\rico_sample.htm"
            Case 8  'scriptaculous
                Archivo = util.StripPath(App.Path) & "libraries\scriptaculous\scriptaculous_sample.template"
                ArchDest = util.StripPath(App.Path) & "libraries\scriptaculous\scriptaculous_sample.htm"
            Case 9  'yahoo ui
                Archivo = util.StripPath(App.Path) & "libraries\yahoo\yahoo_sample.template"
                ArchDest = util.StripPath(App.Path) & "libraries\yahoo\yahoo_sample.htm"
        End Select
        
        util.BorrarArchivo ArchDest
        util.CopiarArchivo Archivo, ArchDest
        frmMain.opeEdit ArchDest
                
        Unload Me
    End If
    
End Sub

Private Sub insertar_html(ByVal tipo As Integer)

    Dim src As New cStringBuilder
    
    If tipo = 1 Then
        src.Append "<!DOCTYPE HTML PUBLIC " & Chr$(34) & "-//W3C//DTD HTML 4.01 Transitional//EN" & Chr$(34) & " "
        src.Append Chr$(34) & "http://www.w3.org/TR/html4/loose.dtd" & Chr$(34) & ">" & vbNewLine
        src.Append "" & vbNewLine
        src.Append "<html>" & vbNewLine
        src.Append "<head>" & vbNewLine
        src.Append "<title></title>" & vbNewLine
        src.Append "</head>" & vbNewLine
        src.Append "<body>" & vbNewLine
        src.Append "</body>" & vbNewLine
        src.Append "</html>" & vbNewLine
    ElseIf tipo = 2 Then
        'xhtml 1.0
        src.Append "<!DOCTYPE html PUBLIC " & Chr$(34) & "-//W3C//DTD XHTML 1.0 Transitional//EN" & Chr$(34) & " "
        src.Append Chr$(34) & "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd" & Chr$(34) & ">" & vbNewLine
        src.Append "" & vbNewLine
        src.Append "<html xmlns=" & Chr$(34) & "http://www.w3.org/1999/xhtml" & Chr$(34) & " xml:lang=" & Chr$(34) & "en" & Chr$(34) & " lang=" & Chr$(34) & "en" & Chr$(34) & ">" & vbNewLine
        src.Append "<head>" & vbNewLine
        src.Append "<title></title>" & vbNewLine
        src.Append "</head>" & vbNewLine
        src.Append "<body>" & vbNewLine
        src.Append "</body>" & vbNewLine
        src.Append "</html>" & vbNewLine
    ElseIf tipo = 3 Then
        'xhtml 1.1
       src.Append "<!DOCTYPE html PUBLIC " & Chr$(34) & "-//W3C//DTD XHTML 1.1 Transitional//EN" & Chr$(34) & " "
        src.Append Chr$(34) & "http://www.w3.org/TR/xhtml1/DTD/xhtml11-transitional.dtd" & Chr$(34) & ">" & vbNewLine
        src.Append "" & vbNewLine
        src.Append "<html xmlns=" & Chr$(34) & "http://www.w3.org/1999/xhtml" & Chr$(34) & " xml:lang=" & Chr$(34) & "en" & Chr$(34) & " lang=" & Chr$(34) & "en" & Chr$(34) & ">" & vbNewLine
        src.Append "<head>" & vbNewLine
        src.Append "<title></title>" & vbNewLine
        src.Append "</head>" & vbNewLine
        src.Append "<body>" & vbNewLine
        src.Append "</body>" & vbNewLine
        src.Append "</html>" & vbNewLine
    End If
    
    frmMain.ActiveForm.Insertar src.ToString
        
End Sub



Private Sub nuevo_html()

    If Not lvwTab(1).SelectedItem Is Nothing Then
        Me.Hide
        Select Case lvwTab(1).SelectedItem.Index
            Case 1  'default html
                frmMain.newEdit
                frmMain.HtmlBody
            Case 2  'empty html
                frmMain.newEdit
            Case 3  'frameset
                frmMain.newEdit
                frmFramesWiz.Show vbModal
            Case 4  'html 4.01
                frmMain.newEdit
                Call insertar_html(1)
            Case 5  'xhtml 1.0
                frmMain.newEdit
                Call insertar_html(2)
            Case 6  'xhtml 1.1
                frmMain.newEdit
                Call insertar_html(3)
        End Select
        Unload Me
    End If
    
End Sub

Private Sub nuevo_javascript()

    If Not lvwTab(0).SelectedItem Is Nothing Then
        Me.Hide
        
        Select Case lvwTab(0).SelectedItem.Index
            Case 1  'empty file
                frmMain.newEdit
            Case 2  'array
                frmMain.newEdit
                frmArray.Show vbModal
            Case 3  'function
                frmMain.newEdit
                If Not frmMain.ActiveForm Is Nothing Then
                    Dim funcion As String
                    funcion = InputBox("Function Name:", "New Function")
                    If Len(Trim$(funcion)) > 0 Then
                        frmMain.ActiveForm.Insertar "function " & funcion & "()" & vbNewLine & "{" & vbNewLine & vbNewLine & "}"
                    End If
                End If
            Case 4  'statemantes
                frmMain.newEdit
                frmStatements.Show vbModal
            Case 5  'regular expression
                frmMain.newEdit
                frmRegExp.Show vbModal
            Case 6  'window
                frmMain.newEdit
                frmNewWindow.Show vbModal
        End Select
        Unload Me
    End If

End Sub

Private Sub nuevo_plus()

    'verificar si hay algo seleccionado
    If Not lvwTab(3).SelectedItem Is Nothing Then
        Me.Hide
        Select Case lvwTab(3).SelectedItem.Index
            Case 1  'addfav
                frmAddFavorites.Show vbModal
            Case 2  'conmenu
                frmCountryMenus.Show vbModal
            Case 3  'dropmenu
                frmDropDownMenu.Show vbModal
            Case 4  'email link
                frmCreateEmail.Show vbModal
            Case 5  'iframe
                frmIframe.Show vbModal
            Case 6  'image rollover
                frmRollover.Show vbModal
            Case 7  'last mode date
                frmLastModDate.Show vbModal
            Case 8  'left menu
                frmLeftMenu.Show vbModal
            Case 9  'meta tag
                frmMetaTag.Show vbModal
            Case 10 'page trans
                frmPageTran.Show vbModal
            Case 11 'popup
                frmPopup.Show vbModal
            Case 12 'colored scrolbar
                frmCreateColScrollbar.Show vbModal
            Case 13 'mouseover
                frmMouseOverLinks.Show vbModal
            Case 14 'popupmenu
                frmPopupMenu.Show vbModal
            Case 15 'tabmenu
                frmTabMenu.Show vbModal
            Case 16 'treemenu
                frmTreeMenu.Show vbModal
            Case 17 'plus
                frmCalendar.Show vbModal
            Case 18 'slideshow
                frmSlideShow.Show vbModal
        End Select
        Unload Me
    End If
    
End Sub

Private Sub nuevo_template()

   If Not lvwTab(2).SelectedItem Is Nothing Then
      Me.Hide
      
      Dim Archivo As String
      Dim arr_files() As String
      Dim sFolder As String
      Dim j As Integer
      
      Archivo = lvwTab(2).SelectedItem.Tag
        
      If ArchivoExiste2(Archivo) Then
         ' Get the file directory:
         m_cUnzip.ZipFile = Archivo
         m_cUnzip.OverwriteExisting = True
         m_cUnzip.UseFolderNames = True
         m_cUnzip.Directory
      
         If m_cUnzip.FileCount > 0 Then
             For j = 1 To m_cUnzip.FileCount
                 m_cUnzip.FileSelected(j) = True
             Next j
             
             sFolder = util.PathArchivo(Archivo)
             m_cUnzip.UnzipFolder = sFolder
             m_cUnzip.Unzip
             
             Archivo = GetFileWithoutExtension(VBArchivoSinPath(Archivo))
             Archivo = Archivo & "\" & Archivo
             get_files_from_folder sFolder & Archivo, arr_files
             
             For j = 1 To UBound(arr_files)
                 If ListaLangs.IsValidExt(arr_files(j)) Then
                     frmMain.opeEdit arr_files(j)
                 End If
             Next j
         End If
      Else
          MsgBox "File :" & Archivo & " was not found", vbCritical
      End If
      DoEvents
      Unload Me
   End If
    
End Sub

Private Sub samples_learners()

    Dim src As New cStringBuilder
    
    If Not lvwTab(4).SelectedItem Is Nothing Then
        Me.Hide
        Select Case lvwTab(4).SelectedItem.Index
            Case 1  'hello world
                Call frmMain.newEdit
                src.Append "<html>" & vbNewLine
                src.Append "<body onload=" & Chr$(34) & "javascript:hello_world();" & Chr$(34) & ">" & vbNewLine
                src.Append "<script language=" & Chr$(34) & "Javascript" & Chr$(34) & ">" & vbNewLine
                src.Append "<!--" & vbNewLine
                src.Append "function hello_world()" & vbNewLine
                src.Append "{" & vbNewLine
                src.Append "   alert(" & Chr$(34) & "Hello World" & Chr$(34) & ");" & vbNewLine
                src.Append "}" & vbNewLine
                src.Append "-->" & vbNewLine
                src.Append "</script>" & vbNewLine
                src.Append "</body>" & vbNewLine
                src.Append "</html>" & vbNewLine
                
                frmMain.ActiveForm.Insertar src.ToString
            Case 2  'alert
                Call frmMain.newEdit
                frmMain.CreateAlert
            Case 3  'confirm
                Call frmMain.newEdit
                frmMain.CreateConfirm
            Case 4  'prompt
                Call frmMain.newEdit
                frmMain.CreatePrompt
            Case 5  'function
                Call frmMain.newEdit
                src.Append "<script language=" & Chr$(34) & "Javascript" & Chr$(34) & ">" & vbNewLine
                src.Append "<!--" & vbNewLine
                src.Append "function my_function(x,y)" & vbNewLine
                src.Append "{" & vbNewLine
                src.Append "   //this is my first javascript function" & vbNewLine
                src.Append "   return x+y;" & vbNewLine
                src.Append "}" & vbNewLine
                src.Append "-->" & vbNewLine
                src.Append "</script>" & vbNewLine
                frmMain.ActiveForm.Insertar src.ToString
            Case 6  'write
                src.Append "<script language=" & Chr$(34) & "Javascript" & Chr$(34) & ">" & vbNewLine
                src.Append "<!--" & vbNewLine
                src.Append "function write_function(value)" & vbNewLine
                src.Append "{" & vbNewLine
                src.Append "   document.write(value);" & vbNewLine
                src.Append "}" & vbNewLine
                src.Append "-->" & vbNewLine
                src.Append "</script>" & vbNewLine
                frmMain.ActiveForm.Insertar src.ToString
            Case 7
                '.ListItems.Add , "f7", "Javascript Block HTML", 20, 20
                src.Append "<script language=" & Chr$(34) & "JavaScript" & Chr$(34) & ">" & vbNewLine
                src.Append "<!--" & vbNewLine
                src.Append "" & vbNewLine
                src.Append "" & vbNewLine
                src.Append "-->" & vbNewLine
                src.Append "</script>"
                frmMain.ActiveForm.Insertar src.ToString
            Case 8
                '.ListItems.Add , "f8", "Javascript Block XHTML", 20, 20
                src.Append "<script language=" & Chr$(34) & "JavaScript" & Chr$(34) & ">" & vbNewLine
                src.Append "/*<![CDATA[*/" & vbNewLine
                src.Append "" & vbNewLine
                src.Append "" & vbNewLine
                src.Append "/*]]>*/" & vbNewLine
                src.Append "</script>"
                frmMain.ActiveForm.Insertar src.ToString
            Case 9
                '.ListItems.Add , "f9", "Javascript Link", 20, 20
                src.Append "<script language=" & Chr$(34) & "JavaScript" & Chr$(34) & " src=" & Chr$(34) & Chr$(34) & "></script>"
                frmMain.ActiveForm.Insertar src.ToString
        End Select
        Unload Me
    End If
    
End Sub

Private Sub setear_listviews()

    ImageList1.ListImages.Add , "k1", LoadResPicture(112, vbResIcon)
    
    'plus
    ImageList1.ListImages.Add , "k2", LoadResPicture(143, vbResIcon)
    ImageList1.ListImages.Add , "k3", LoadResPicture(144, vbResIcon)
    ImageList1.ListImages.Add , "k4", LoadResPicture(136, vbResIcon)
    ImageList1.ListImages.Add , "k5", LoadResPicture(145, vbResIcon)
    ImageList1.ListImages.Add , "k6", LoadResPicture(146, vbResIcon)
    ImageList1.ListImages.Add , "k7", LoadResPicture(147, vbResIcon)
    ImageList1.ListImages.Add , "k8", LoadResPicture(148, vbResIcon)
    ImageList1.ListImages.Add , "k9", LoadResPicture(149, vbResIcon)
    ImageList1.ListImages.Add , "k10", LoadResPicture(150, vbResIcon)
    ImageList1.ListImages.Add , "k11", LoadResPicture(151, vbResIcon)
    ImageList1.ListImages.Add , "k12", LoadResPicture(152, vbResIcon)
    ImageList1.ListImages.Add , "k13", LoadResPicture(153, vbResIcon)
    ImageList1.ListImages.Add , "k14", LoadResPicture(154, vbResIcon)
    
    ImageList1.ListImages.Add , "k15", LoadResPicture(159, vbResIcon) 'array
    ImageList1.ListImages.Add , "k16", LoadResPicture(161, vbResIcon) 'funcion
    ImageList1.ListImages.Add , "k17", LoadResPicture(162, vbResIcon) 'regexp
    ImageList1.ListImages.Add , "k18", LoadResPicture(163, vbResIcon) 'statement
    ImageList1.ListImages.Add , "k19", LoadResPicture(250, vbResIcon) 'my template
    ImageList1.ListImages.Add , "k20", LoadResPicture(251, vbResIcon) 'html doc
    ImageList1.ListImages.Add , "k21", LoadResPicture(204, vbResIcon) 'script
    ImageList1.ListImages.Add , "k22", LoadResPicture(204, vbResIcon) 'script
    
    ImageList1.ListImages.Add , "k23", LoadResPicture(255, vbResIcon) 'popup menu
    ImageList1.ListImages.Add , "k24", LoadResPicture(256, vbResIcon) 'tab menu
    ImageList1.ListImages.Add , "k25", LoadResPicture(257, vbResIcon) 'treemenu
    ImageList1.ListImages.Add , "k26", LoadResPicture(265, vbResIcon) 'calendar
    ImageList1.ListImages.Add , "k27", LoadResPicture(267, vbResIcon) 'slideshow
    
    'Javascript
    With lvwTab(0)
        .ListItems.Add , "f1", "Empty File", 1, 1
        .ListItems.Add , "f2", "Array", 15, 15
        .ListItems.Add , "f3", "Function", 16, 16
        .ListItems.Add , "f4", "Statements", 18, 18
        .ListItems.Add , "f5", "Regular Expresion", 17, 17
        .ListItems.Add , "f6", "Window", 12, 12
    End With
    
    'Html
    With lvwTab(1)
        .ListItems.Add , "f1", "Default", 20, 20
        .ListItems.Add , "f2", "Empty Page", 20, 20
        .ListItems.Add , "f3", "Frameset", 20, 20
        .ListItems.Add , "f4", "HTML 4.01", 20, 20
        .ListItems.Add , "f5", "XHTML 1.0", 20, 20
        .ListItems.Add , "f6", "XHTML 1.1", 20, 20
    End With
    
    'css
    With lvwTab(2)
        Dim arr_css() As String
        Dim k As Integer
        
        get_files_from_folder StripPath(App.Path) & "css", arr_css
        
        For k = 1 To UBound(arr_css)
            .ListItems.Add , "k" & k, VBArchivoSinPath(arr_css(k)), 19, 19
            .ListItems("k" & k).Tag = arr_css(k)
        Next k
    End With
    
    'plus
    With lvwTab(3)
        .ListItems.Add , "f1", "Add to Favorites", 2, 2
        .ListItems.Add , "f2", "Countries Menu", 3, 3
        .ListItems.Add , "f3", "Drop Down Menu", 4, 4
        .ListItems.Add , "f4", "Email Link", 5, 5
        .ListItems.Add , "f5", "IFrame", 6, 6
        .ListItems.Add , "f6", "Image Rollover", 7, 7
        .ListItems.Add , "f7", "Last Modified Date", 8, 8
        .ListItems.Add , "f8", "Left Menu", 9, 9
        .ListItems.Add , "f9", "Metatag Wizard", 10, 10
        .ListItems.Add , "f10", "Page Transition", 11, 11
        .ListItems.Add , "f11", "Popup Window", 12, 12
        .ListItems.Add , "f12", "Coloured Scrollbar", 13, 13
        .ListItems.Add , "f13", "MouseOver Textlinks", 14, 14
        .ListItems.Add , "f14", "PopupMenu", 23, 23
        .ListItems.Add , "f15", "TabMenu", 24, 24
        .ListItems.Add , "f16", "TreeMenu", 25, 25
        .ListItems.Add , "f17", "Calendar", 26, 26
        .ListItems.Add , "f18", "SlideShow", 27, 27
    End With
    
    'samples for learners
    With lvwTab(4)
        .ListItems.Add , "f1", "Hello World", 21, 21
        .ListItems.Add , "f2", "Alert", 22, 22
        .ListItems.Add , "f3", "Confirm", 21, 21
        .ListItems.Add , "f4", "Prompt", 21, 21
        .ListItems.Add , "f5", "Function", 21, 21
        .ListItems.Add , "f6", "Write", 21, 21
        .ListItems.Add , "f7", "JavaScript Block HTML", 21, 21
        .ListItems.Add , "f8", "JavaScript Block XHTML", 21, 21
        .ListItems.Add , "f9", "JavaScript Link", 21, 21
    End With
        
    'my templates
    With lvwTab(5)
        .ListItems.Add , "f1", "AFLAX", 1, 1
        .ListItems.Add , "f2", "Dojo", 1, 1
        .ListItems.Add , "f3", "JQuery", 1, 1
        .ListItems.Add , "f4", "Google", 1, 1
        .ListItems.Add , "f5", "Mochikit", 1, 1
        .ListItems.Add , "f6", "Prototype", 1, 1
        .ListItems.Add , "f7", "Rico", 1, 1
        .ListItems.Add , "f8", "Scriptaculous", 1, 1
        .ListItems.Add , "f9", "Yahoo UI", 1, 1
    End With
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If tabactivo = 0 Then   'javascript
            Call nuevo_javascript
        ElseIf tabactivo = 1 Then   'html
            Call nuevo_html
        ElseIf tabactivo = 2 Then   'css
            Call nuevo_template
        ElseIf tabactivo = 3 Then   'plus
            Call nuevo_plus
        ElseIf tabactivo = 4 Then   'samples
            Call samples_learners
        ElseIf tabactivo = 5 Then   'ajax
            Call carga_template_ajax
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Dim k As Integer
    
    fcargando = True
    
    util.CenterForm Me
    util.Hourglass hwnd, True
        
    Set m_cUnzip = New cUnzip
    
    For k = 1 To 5
        Load lvwTab(k)
        lvwTab(k).Visible = True
    Next k
    
    fcargando = False
    
    Form_Resize
    
    Call setear_listviews
    
    lvwTab(0).ZOrder 0
    
    fcargando = False
    
    util.Hourglass hwnd, False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Resize()

    If WindowState <> vbMinimized Then
        If fcargando Then Exit Sub
        lvwTab(1).Move lvwTab(0).Left, lvwTab(0).Top, lvwTab(0).Width, lvwTab(0).Height
        lvwTab(2).Move lvwTab(0).Left, lvwTab(0).Top, lvwTab(0).Width, lvwTab(0).Height
        lvwTab(3).Move lvwTab(0).Left, lvwTab(0).Top, lvwTab(0).Width, lvwTab(0).Height
        lvwTab(4).Move lvwTab(0).Left, lvwTab(0).Top, lvwTab(0).Width, lvwTab(0).Height
        lvwTab(5).Move lvwTab(0).Left, lvwTab(0).Top, lvwTab(0).Width, lvwTab(0).Height
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmNewDoc = Nothing
End Sub





Private Sub lvwTab_DblClick(Index As Integer)
    Call cmd_Click(0)
End Sub





Private Sub tabMain_Click()
    tabactivo = tabMain.SelectedItem.Index - 1
    lvwTab(tabactivo).ZOrder 0
End Sub

