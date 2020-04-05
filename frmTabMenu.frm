VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTabMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TabMenu Wizard"
   ClientHeight    =   5220
   ClientLeft      =   4080
   ClientTop       =   2685
   ClientWidth     =   6000
   Icon            =   "frmTabMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "P&review"
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "A&pply"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Includes files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   60
      TabIndex        =   6
      Top             =   4065
      Width           =   4515
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "tabmenu.css, tabmenu.js"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   6
         Left            =   105
         TabIndex        =   7
         Top             =   255
         Width           =   2250
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.ListView lvwTab 
      Height          =   2580
      Left            =   60
      TabIndex        =   4
      Top             =   1410
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   4551
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Caption"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Link"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtLink 
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Top             =   870
      Width           =   4500
   End
   Begin VB.TextBox txtCaption 
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   270
      Width           =   4515
   End
   Begin VB.Image imgOP 
      Height          =   255
      Left            =   5625
      Top             =   4875
      Width           =   300
   End
   Begin VB.Image imgNE 
      Height          =   255
      Left            =   5295
      Top             =   4875
      Width           =   300
   End
   Begin VB.Image imgFX 
      Height          =   255
      Left            =   4980
      Top             =   4875
      Width           =   300
   End
   Begin VB.Image imgIE 
      Height          =   255
      Left            =   4665
      Top             =   4875
      Width           =   300
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
      Left            =   3060
      TabIndex        =   8
      Top             =   4890
      Width           =   1485
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "TabMenu"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   1215
      Width           =   690
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Link"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   645
      Width           =   300
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Caption"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   540
   End
End
Attribute VB_Name = "frmTabMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private nKey As Integer
Private ultimo_path As String
Private Sub generar_tabmenu(ByVal preview As Boolean)

    On Error GoTo Errorpreview_treemenu
    
    Dim Archivo As String
    Dim nFreeFile As Long
    Dim buffer As New cStringBuilder
    Dim glosa As String
    Dim pathapp As String
    Dim k As Integer
    Dim inicio As String
    
    util.Hourglass hwnd, True
    
    pathapp = util.StripPath(App.Path) & "plus\tabmenu\"
    
    If lvwTab.ListItems.count > 0 Then
    
        inicio = lvwTab.ListItems(1).SubItems(1)
        
        buffer.Append "<html>" & vbNewLine
        buffer.Append "   <head>" & vbNewLine
        buffer.Append "      <title>Tab Menu</title>" & vbNewLine
        buffer.Append "      <link rel='stylesheet' href='tabmenu.css'>" & vbNewLine
        buffer.Append "      <script language=javascript src=tabmenu.js></script>" & vbNewLine
        buffer.Append "   </head>" & vbNewLine
        buffer.Append "<body>" & vbNewLine
        buffer.Append "<div class='tabBox' style='clear:both;'>" & vbNewLine
        buffer.Append "<div class='tabArea'>" & vbNewLine
                
        For k = 1 To lvwTab.ListItems.count
            buffer.Append "<a class='tab' href='" & lvwTab.ListItems(k).SubItems(1) & "' target='tabIframe2'>" & lvwTab.ListItems(k).Text & "</a>" & vbNewLine
        Next k
        
        buffer.Append "</div>" & vbNewLine
        buffer.Append "<div class='tabMain'>" & vbNewLine
        buffer.Append "<div class='tabIframeWrapper'>" & vbNewLine
        buffer.Append "<iframe class='tabContent' name='tabIframe2' src='" & inicio & "' marginheight='8' marginwidth='8' frameborder='0'></iframe>" & vbNewLine
        buffer.Append "</div>" & vbNewLine
        buffer.Append "</div>" & vbNewLine
        buffer.Append "</div>" & vbNewLine
        buffer.Append "</body>" & vbNewLine
        buffer.Append "</html>" & vbNewLine
        
        nFreeFile = FreeFile
        
        glosa = "Hypertext files (*.htm)|*.htm|"
        glosa = glosa & "All Files (*.*)|*.*"
    
        If preview Then
            Archivo = util.StripPath(App.Path) & "tabmenu.htm"
            Open Archivo For Output As #nFreeFile
                Print #nFreeFile, buffer.ToString
            Close #nFreeFile
        
            'copiar los archivos necesarios para generar esto
            util.CopiarArchivo pathapp & "tabmenu.css", util.StripPath(App.Path) & "tabmenu.css"
            util.CopiarArchivo pathapp & "tabmenu.js", util.StripPath(App.Path) & "tabmenu.js"
                    
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
                    util.CopiarArchivo pathapp & "tabmenu.css", ultimo_path & "tabmenu.css"
                    util.CopiarArchivo pathapp & "tabmenu.js", ultimo_path & "tabmenu.js"
                                    
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
    
    util.Hourglass hwnd, False
    
    Exit Sub
    
Errorpreview_treemenu:
    MsgBox "generar_tabmenu : " & Err & " " & Error$, vbCritical
    util.Hourglass hwnd, False

End Sub

Private Sub cmdAdd_Click()

    If txtCaption.Text = "" Then
        txtCaption.SetFocus
        Exit Sub
    End If
    
    If txtLink.Text = "" Then
        txtLink.SetFocus
        Exit Sub
    End If
    
    lvwTab.ListItems.Add , "key" & nKey, txtCaption.Text
    lvwTab.ListItems("key" & nKey).SubItems(1) = txtLink.Text
    
    nKey = nKey + 1
    
End Sub

Private Sub cmdAplicar_Click()
    Call generar_tabmenu(False)
End Sub

Private Sub cmdDelete_Click()

    If Not lvwTab.SelectedItem Is Nothing Then
        lvwTab.ListItems.Remove lvwTab.SelectedItem.key
    End If
    
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdPreview_Click()
    Call generar_tabmenu(True)
End Sub

Private Sub Form_Load()

    util.Hourglass hwnd, True
    util.CenterForm Me
        
    nKey = 1
    
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
    Set frmTabMenu = Nothing
End Sub


