VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLibraryManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Library Manager"
   ClientHeight    =   4635
   ClientLeft      =   1350
   ClientTop       =   3225
   ClientWidth     =   11430
   ControlBox      =   0   'False
   Icon            =   "frmLibraryManager.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Update"
      Height          =   375
      Index           =   5
      Left            =   10080
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "R&efresh"
      Height          =   375
      Index           =   4
      Left            =   10080
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&New"
      Height          =   375
      Index           =   3
      Left            =   10080
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Exit"
      Height          =   375
      Index           =   2
      Left            =   10080
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Remove"
      Height          =   375
      Index           =   1
      Left            =   10080
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   4065
      Left            =   45
      TabIndex        =   1
      Top             =   255
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   7170
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "FileName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Autor"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Version"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Active"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "If you active/deactive library you must need restart to changes take effect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   45
      TabIndex        =   2
      Top             =   4365
      Width           =   6360
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Installed Libraries"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1515
   End
End
Attribute VB_Name = "frmLibraryManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cargar_librerias()
    
    Dim k As Integer
    Dim Activa As String
    Dim Path As String
    
    Dim arr_files() As String
    Dim arr_librerias() As String
    Dim j As Integer
    
    lvw.ListItems.Clear
    
    Path = util.StripPath(App.Path) & "libraries"
    
    get_files_from_folder Path, arr_files
    
    If UBound(arr_files) > 0 Then
        For j = 1 To UBound(arr_files)
            
            get_info_section "information", arr_librerias, arr_files(j)
            
            If Len(util.LeeIni(arr_files(j), "information", "name")) > 0 Then
                lvw.ListItems.Add , "k" & k, util.VBArchivoSinPath(arr_files(j))
                lvw.ListItems("k" & k).SubItems(1) = util.LeeIni(arr_files(j), "information", "name")
                lvw.ListItems("k" & k).SubItems(2) = util.LeeIni(arr_files(j), "information", "autor")
                lvw.ListItems("k" & k).SubItems(3) = util.LeeIni(arr_files(j), "information", "Description")
                lvw.ListItems("k" & k).SubItems(4) = util.LeeIni(arr_files(j), "information", "Version")
                
                Activa = util.LeeIni(arr_files(j), "information", "active")
                If Activa = "Y" Or Activa = "N" Then
                    lvw.ListItems("k" & k).SubItems(5) = Activa
                    If Activa = "Y" Then
                        lvw.ListItems("k" & k).Checked = True
                    Else
                        lvw.ListItems("k" & k).Checked = False
                    End If
                Else
                    lvw.ListItems("k" & k).SubItems(5) = "N"
                    lvw.ListItems("k" & k).Selected = False
                End If
                k = k + 1
            End If
        Next j
    End If
    
End Sub
Private Sub remover_libreria()

    Dim Path As String
    
    Path = util.StripPath(App.Path) & "libraries\" & lvw.SelectedItem.Text
    
    If ArchivoExiste2(Path) Then
        If Confirma("Are you sure to remove this library") = vbYes Then
            util.BorrarArchivo Path
            cargar_librerias
        End If
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
    
    ElseIf Index = 1 Then   'remove
        If Not lvw.SelectedItem Is Nothing Then
            Call remover_libreria
        Else
            MsgBox "Select a library first", vbCritical
            Exit Sub
        End If
    ElseIf Index = 2 Then   'exit
        Unload Me
    ElseIf Index = 3 Then   'new
        frmImportLibrary.Show vbModal
    ElseIf Index = 4 Then   'refresh
        Call cargar_librerias
    ElseIf Index = 5 Then   'update
        If Not lvw.SelectedItem Is Nothing Then
            frmImportLibrary.Archivo = lvw.SelectedItem.Text
        Else
            MsgBox "Select a library first", vbCritical
            Exit Sub
        End If
        frmImportLibrary.Show vbModal
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
    
    Call cargar_librerias
        
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmLibraryManager = Nothing
End Sub


Private Sub lvw_ItemCheck(ByVal ITem As MSComctlLib.ListItem)

    Dim ret As String
    
    If ITem.Checked Then
        ret = "Y"
    Else
        ret = "N"
    End If
    
    util.GrabaIni util.StripPath(App.Path) & "libraries\" & ITem.Text, "information", "active", ret
        
    ITem.SubItems(5) = ret
        
End Sub

