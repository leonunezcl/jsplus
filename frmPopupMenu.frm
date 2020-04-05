VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPopupMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PopupMenu Wizard"
   ClientHeight    =   6750
   ClientLeft      =   5340
   ClientTop       =   2265
   ClientWidth     =   4935
   Icon            =   "frmPopupMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   4935
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3600
      TabIndex        =   18
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Add"
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   7185
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2775
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.ComboBox cboPos 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtLink 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtTitulo 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3255
   End
   Begin MSComctlLib.TreeView tvwMenu 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   7011
      _Version        =   393217
      Indentation     =   706
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tv2 
      Height          =   1095
      Left            =   8040
      TabIndex        =   10
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1931
      _Version        =   393217
      Indentation     =   706
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopupMenu.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopupMenu.frx":05A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopupMenu.frx":0B40
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Browser Compatibility"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   1980
      TabIndex        =   13
      Top             =   6465
      Width           =   1485
   End
   Begin VB.Image imgIE 
      Height          =   255
      Left            =   3585
      Top             =   6450
      Width           =   300
   End
   Begin VB.Image imgFX 
      Height          =   255
      Left            =   3900
      Top             =   6450
      Width           =   300
   End
   Begin VB.Image imgNE 
      Height          =   255
      Left            =   4215
      Top             =   6450
      Width           =   300
   End
   Begin VB.Image imgOP 
      Height          =   255
      Left            =   4545
      Top             =   6450
      Width           =   300
   End
   Begin VB.Label lblNodo 
      AutoSize        =   -1  'True
      Caption         =   "Nodes"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   465
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Position"
      Height          =   195
      Index           =   5
      Left            =   1800
      TabIndex        =   9
      Top             =   720
      Width           =   555
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Type"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Link"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Title"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   300
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Nodes"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   465
   End
End
Attribute VB_Name = "frmPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private util As New cLibrary
Private nKey As Integer
Private buffer As New cStringBuilder
Private nContador As Integer
Private ultimo_path As String
Private Sub elimina_nodos(ByVal Nodo As Node)

    Dim k As Integer
   
    If Nodo.Children > 1 Then
        For k = 1 To Nodo.Children
            elimina_nodos Nodo.Child
        Next k
        On Error Resume Next
        tvwMenu.Nodes.Remove Nodo.key
        Err = 0
    Else
        tvwMenu.Nodes.Remove Nodo.key
    End If
    
End Sub

Private Sub genera_folder(Nodo As Node)

    Dim a_id As String
        
    If InStr(Nodo.key, "folder") Then
        a_id = Replace(Trim$(Nodo.Text), " ", "")
        buffer.Append "<table border=0 cellpadding=1 cellspacing=1>" & vbNewLine
        buffer.Append "   <tr>" & vbNewLine
        buffer.Append "      <td width='16'>" & vbNewLine
        buffer.Append "          <a id=" & Chr$(34) & "x" & a_id & Chr$(34) & " href=" & Chr$(34) & "javascript:Toggle('" & a_id & "');" & Chr$(34) & ">" & vbNewLine
        buffer.Append "             <img src='folder.gif' width='16' height='16' hspace='0' vspace='0' border='0'>" & vbNewLine
        buffer.Append "          </a>" & vbNewLine
        buffer.Append "      </td>" & vbNewLine
        buffer.Append "      <td>" & vbNewLine
        buffer.Append "         <b>" & Nodo.Text & "</b>" & vbNewLine
        buffer.Append "      </td>" & vbNewLine
        buffer.Append "</table>" & vbNewLine
    End If
    
End Sub

Private Sub genera_treemenu_javascript(tvFrom As TreeView, tvTo As TreeView, CurNode As Node, Optional ByVal RootNode As Boolean = False)

    Dim lo_child As Node
    Dim lo_add As Node
            
    If CurNode.Children > 0 Then
        'Recursivamente recorrer cada hijo del nodo y verificar si este nodo hijo tiene mas hijos
        Set lo_child = CurNode.Child.FirstSibling
        If RootNode = False Then
            If Not CurNode.Parent Is Nothing Then
                Set lo_add = tvTo.Nodes.Add(CurNode.Parent.key, tvwChild, CurNode.key, CurNode.Text, CurNode.Image, CurNode.SelectedImage)
            Else
                Set lo_add = tvTo.Nodes.Add(, , CurNode.key, CurNode.Text, CurNode.Image, CurNode.SelectedImage)
            End If
            lo_add.Expanded = CurNode.Expanded
            lo_add.Tag = CurNode.Tag '.Expanded
        End If
        
        Call genera_folder(lo_add)
        
        buffer.Append "<div id=" & Chr$(34) & lo_add.Text & "Menu" & Chr$(34) & " class=" & Chr$(34) & "menu" & Chr$(34) & ">" & vbNewLine
                
        'recursivamente verificar cada hijo y agregarlo al nodo correspondiente
        Do While Not lo_child Is Nothing
            Call genera_treemenu_javascript(tvFrom, tvTo, lo_child)
            Set lo_child = lo_child.Next
        Loop
    
        buffer.Append "</div>" & vbNewLine
        
    Else ' Si no hay hijos entonces solo agregar el nodo
        If Not CurNode.Parent Is Nothing Then
            Set lo_add = tvTo.Nodes.Add(CurNode.Parent.key, tvwChild, CurNode.key, CurNode.Text, CurNode.Image, CurNode.SelectedImage)
        Else
            Set lo_add = tvTo.Nodes.Add(, , CurNode.key, CurNode.Text, CurNode.Image, CurNode.SelectedImage)
        End If
        
        lo_add.Expanded = CurNode.Expanded
        lo_add.Tag = CurNode.Tag '.Expanded
        
        buffer.Append "<a class=" & Chr$(34) & "menuItem" & Chr$(34) & " href=" & Chr$(34) & lo_add.Tag & Chr$(34) & ">" & lo_add.Text & "</a>" & vbNewLine
    End If
            
End Sub

Private Sub generar_tree_menu()

    Dim Nodo As Node
    Dim k As Integer
   
    'recorrer los nodos raiz y generar el arbol
    For k = 1 To tvwMenu.Nodes.count
        Set Nodo = tvwMenu.Nodes(k)
        If Nodo.Parent Is Nothing Then
            genera_treemenu_javascript tvwMenu, tv2, Nodo
        End If
    Next k
    
End Sub

Private Sub preview_treemenu(ByVal mostrar As Boolean)

    Dim Archivo As String
    Dim nFreeFile As Long
    Dim linea As String
    Dim Nodo As Node
    Dim k As Integer
    Dim titulo As String
    Dim pathapp As String
    Dim pathwrk As String
    Dim glosa As String
    
    util.Hourglass hwnd, True
    
    tv2.Nodes.Clear
    
    pathapp = util.StripPath(App.Path) & "plus\popupmenu\"
    pathwrk = util.StripPath(App.Path)
    
    If tvwMenu.Nodes.count > 0 Then
    
        buffer.Append "<html>" & vbNewLine
        buffer.Append "   <head>" & vbNewLine
        buffer.Append "      <title>MenuBar Test</title>" & vbNewLine
        buffer.Append "      <link rel='stylesheet' href='menubar.css'>" & vbNewLine
        buffer.Append "      <script language=javascript src=menubar.js></script>" & vbNewLine
        buffer.Append "   </head>" & vbNewLine
        buffer.Append "<body>" & vbNewLine
                
        'generar las menubar
        buffer.Append "<div class='menuBar' style='width:80%;'>" & vbNewLine
        For k = 1 To tvwMenu.Nodes.count
            Set Nodo = tvwMenu.Nodes(k)
            If Nodo.Parent Is Nothing Then
                titulo = Nodo.Text & "Menu"
                linea = "<a class=" & Chr$(34) & "menuButton" & Chr$(34) & " href=" & Chr$(34) & Chr$(34) & " onclick=" & Chr$(34) & "return buttonClick(event, '" & titulo & "');" & Chr$(34)
                linea = linea & " onmouseover=" & Chr$(34) & "buttonMouseover(event, '" & titulo & "');" & Chr$(34) & ">" & Nodo.Text & "</a>" & vbNewLine
                buffer.Append linea
            End If
        Next k
        buffer.Append "</div>" & vbNewLine
        
        'sMenu = "menu"
        nContador = 1
        Call generar_tree_menu
        
        buffer.Append "</body>" & vbNewLine
        buffer.Append "</html>" & vbNewLine
        
        nFreeFile = FreeFile
        
        glosa = "Hypertext files (*.htm)|*.htm|"
        glosa = glosa & "All Files (*.*)|*.*"
                                
        If mostrar Then
            Archivo = util.StripPath(App.Path) & "popupmenu.htm"
            
            'copiar los archivos necesarios para generar esto
            util.CopiarArchivo pathapp & "menubar.css", pathwrk & "menubar.css"
            util.CopiarArchivo pathapp & "menubar.js", pathwrk & "menubar.js"
            
            Open Archivo For Output As #nFreeFile
                Print #nFreeFile, buffer.ToString
            Close #nFreeFile
            
            util.ShellFunc Archivo, vbNormalFocus
        Else
            If ultimo_path = "" Then
                ultimo_path = App.Path
            End If
            
            If Cdlg.VBGetSaveFileName(Archivo, , , glosa, , ultimo_path, "Save File As ...", "htm") Then
                
                ultimo_path = util.PathArchivo(Archivo)
                
                If ultimo_path <> pathapp Then
                    Open Archivo For Output As #nFreeFile
                        Print #nFreeFile, buffer.ToString
                    Close #nFreeFile
                    
                    'copiar los archivos necesarios para generar esto
                    util.CopiarArchivo pathapp & "menubar.css", ultimo_path & "menubar.css"
                    util.CopiarArchivo pathapp & "menubar.js", ultimo_path & "menubar.js"
            
                    util.ShellFunc Archivo, vbNormalFocus
                Else
                    MsgBox "Invalid path. You must choice another path", vbCritical
                End If
            End If
        End If
    Else
        MsgBox "Nothing to do", vbCritical
    End If
    
    Set buffer = Nothing
    
End Sub

Private Sub cboTipo_Click()

    If cboTipo.ListIndex = 1 Then
        txtLink.Text = vbNullString
        txtLink.Enabled = False
    Else
        txtLink.Enabled = True
    End If
    
End Sub


Private Sub cmdAgregar_Click()
    
    Dim titulo As String
    Dim tipo As Integer
    Dim icono As Integer
    Dim Pos As Integer
    Dim llave As String
    Dim link As String
    
    titulo = Trim$(txtTitulo.Text)
    tipo = cboTipo.ListIndex
    link = txtLink.Text
    
    If Len(titulo) > 0 Then
        If tipo > -1 Then
            If tipo = 0 Then    'menubar
                icono = 1
                llave = "topmenu" & nKey
            ElseIf tipo = 1 Then 'menu
                icono = 1
                llave = "menu" & nKey
            ElseIf tipo = 2 Then 'submenu
                icono = 3
                llave = "submenu" & nKey
            ElseIf tipo = 3 Then 'separator
                icono = 2
                llave = "separator" & nKey
            End If
            
            Pos = cboPos.ListIndex
            
            If Not tvwMenu.SelectedItem Is Nothing Then
                If InStr(tvwMenu.SelectedItem.key, "topmenu") Then
                    If Pos = 0 Then
                        tvwMenu.Nodes.Add , tvwChild, llave, titulo, icono, icono
                    Else
                        tvwMenu.Nodes.Add tvwMenu.SelectedItem.key, tvwChild, llave, titulo, icono, icono
                    End If
                    tvwMenu.Nodes(llave).Expanded = True
                    tvwMenu.Nodes(llave).Tag = link
                ElseIf InStr(tvwMenu.SelectedItem.key, "menu") Then
                    If Pos = 0 Then
                        tvwMenu.Nodes.Add , tvwChild, llave, titulo, icono, icono
                    Else
                        tvwMenu.Nodes.Add tvwMenu.SelectedItem.key, tvwChild, llave, titulo, icono, icono
                    End If
                    tvwMenu.Nodes(llave).Expanded = True
                    tvwMenu.Nodes(llave).Tag = link
                Else
                    tvwMenu.Nodes.Add , tvwChild, llave, titulo, icono, icono
                    tvwMenu.Nodes(llave).Tag = link
                End If
                nKey = nKey + 1
            Else
                If InStr(llave, "topmenu") Then
                    tvwMenu.Nodes.Add , tvwChild, llave, titulo, icono, icono
                    tvwMenu.Nodes(llave).Expanded = True
                    tvwMenu.Nodes(llave).Tag = link
                ElseIf InStr(llave, "menu") Then
                    tvwMenu.Nodes.Add , tvwChild, llave, titulo, icono, icono
                    tvwMenu.Nodes(llave).Expanded = True
                    tvwMenu.Nodes(llave).Tag = link
                Else
                    tvwMenu.Nodes.Add , tvwChild, llave, titulo, icono, icono
                    tvwMenu.Nodes(llave).Tag = link
                End If
                nKey = nKey + 1
            End If
            txtTitulo.SetFocus
        Else
            cboTipo.SetFocus
        End If
    Else
        txtTitulo.SetFocus
    End If
    
End Sub


Private Sub cmdEliminar_Click()
    If Not tvwMenu.SelectedItem Is Nothing Then
        Call elimina_nodos(tvwMenu.SelectedItem)
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGenerar_Click()

    Call preview_treemenu(False)

End Sub

Private Sub copia_treeview(tvFrom As TreeView, tvTo As TreeView, CurNode As Node, Optional ByVal RootNode As Boolean = False)

    Dim lo_child As Node
    Dim lo_add As Node
            
    If CurNode.Children > 0 Then
        'Recursivamente recorrer cada hijo del nodo y verificar si este nodo hijo tiene mas hijos
        Set lo_child = CurNode.Child.FirstSibling
        If RootNode = False Then
            If Not CurNode.Parent Is Nothing Then
                Set lo_add = tvTo.Nodes.Add(CurNode.Parent.key, tvwChild, CurNode.key, CurNode.Text, CurNode.Image, CurNode.SelectedImage)
            Else
                Set lo_add = tvTo.Nodes.Add(, , CurNode.key, CurNode.Text, CurNode.Image, CurNode.SelectedImage)
            End If
            lo_add.Expanded = CurNode.Expanded
            lo_add.Tag = CurNode.Expanded
        End If
        
        'recursivamente verificar cada hijo y agregarlo al nodo correspondiente
        Do While Not lo_child Is Nothing
            Call copia_treeview(tvFrom, tvTo, lo_child)
            Set lo_child = lo_child.Next
        Loop
    
    Else ' Si no hay hijos entonces solo agregar el nodo
        If Not CurNode.Parent Is Nothing Then
            Set lo_add = tvTo.Nodes.Add(CurNode.Parent.key, tvwChild, CurNode.key, CurNode.Text, CurNode.Image, CurNode.SelectedImage)
        Else
            Set lo_add = tvTo.Nodes.Add(, , CurNode.key, CurNode.Text, CurNode.Image, CurNode.SelectedImage)
        End If
        
        lo_add.Expanded = CurNode.Expanded
        lo_add.Tag = CurNode.Expanded
    End If
    
End Sub


Private Sub cmdPreview_Click()
    Call preview_treemenu(True)
End Sub

Private Sub Form_Load()
    
    util.CenterForm Me
    util.Hourglass hwnd, True
        
    nKey = 1
    
    cboTipo.AddItem "TopMenu"
    cboTipo.AddItem "Menu"
    cboTipo.AddItem "SubMenu"
    cboTipo.AddItem "Separator"
    
    cboTipo.ListIndex = 0
    
    cboPos.AddItem "Top"
    cboPos.AddItem "Child"
    cboPos.ListIndex = 0
    
    Set imgIE.Picture = LoadResPicture(1007, vbResBitmap)
    Set imgFX.Picture = LoadResPicture(1008, vbResBitmap)
    Set imgNE.Picture = LoadResPicture(1009, vbResBitmap)
    Set imgOP.Picture = LoadResPicture(1010, vbResBitmap)
    
    util.Hourglass hwnd, False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmPopupMenu = Nothing
End Sub


Private Sub tvwMenu_NodeClick(ByVal Node As MSComctlLib.Node)
    lblNodo.Caption = Node.FullPath
End Sub


