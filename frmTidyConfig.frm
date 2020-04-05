VERSION 5.00
Object = "{FCFAF346-DE8A-4FB6-8612-5000548EFDC7}#2.0#0"; "vbsListView6.ocx"
Begin VB.Form frmTidyConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTML Tidy Configuration"
   ClientHeight    =   3465
   ClientLeft      =   4080
   ClientTop       =   3975
   ClientWidth     =   6360
   Icon            =   "frmTidyConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Default"
      Height          =   375
      Index           =   5
      Left            =   5040
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "E&xit"
      Height          =   375
      Index           =   4
      Left            =   5040
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Remove Task"
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Edit Task"
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&New Task"
      Height          =   375
      Index           =   1
      Left            =   5040
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Execute"
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin vbalListViewLib6.vbalListViewCtl lvwTasks 
      Height          =   2805
      Left            =   30
      TabIndex        =   0
      Top             =   495
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   4948
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   1
      MultiSelect     =   -1  'True
      LabelEdit       =   0   'False
      AutoArrange     =   0   'False
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTidyConfig.frx":000C
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   30
      TabIndex        =   1
      Top             =   75
      Width           =   5040
   End
End
Attribute VB_Name = "frmTidyConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub cargar_tidy()

    Dim k As Integer
    Dim Archivo As String
    Dim arr_tidy_config() As String
    
    Archivo = util.StripPath(App.Path) & "tidy\tasks.ini"
    
    get_info_section "tasks", arr_tidy_config, Archivo
    
    lvwTasks.ListItems.Clear
    For k = 2 To UBound(arr_tidy_config)
        lvwTasks.ListItems.Add , "k" & k, util.Explode(arr_tidy_config(k), 1, "|")
        lvwTasks.ListItems(k - 1).Tag = Explode(arr_tidy_config(k), 2, "|")
        If Len(util.Explode(arr_tidy_config(k), 3, "|")) > 0 Then
            lvwTasks.ListItems(k - 1).SubItems(1).Caption = "Default"
        End If
    Next k
    
End Sub

Private Sub eliminar_tarea()

    Dim Archivo As String
    Dim k As Integer
    Dim j As Integer
    
    Dim nFreeFile As Long
    
    util.Hourglass hwnd, True
    
    Archivo = util.StripPath(App.Path) & "tidy\tasks.ini"
    
    util.BorrarArchivo Archivo
    
    nFreeFile = FreeFile
    j = 1
    Open Archivo For Output As #nFreeFile
        Print #nFreeFile, "[tasks]"
        Print #nFreeFile, "num=" & lvwTasks.ListItems.count - 1
        
        For k = 1 To lvwTasks.ListItems.count
            If lvwTasks.ListItems(k).key <> lvwTasks.SelectedItem.key Then
                Print #nFreeFile, "tsk" & j & "=" & lvwTasks.ListItems(k).Text & "|" & lvwTasks.ListItems(k).Tag & "|" & lvwTasks.ListItems(k).SubItems(1).Caption
                j = j + 1
            Else
                'borrar el archivo de tarea asociado
                util.BorrarArchivo util.StripPath(App.Path) & "tidy\" & lvwTasks.ListItems(k).Tag
            End If
        Next k
    Close #nFreeFile
    
    util.Hourglass hwnd, False
    
    cargar_tidy
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Dim k As Integer
    Dim ini As String
    
    If Index = 0 Then
        'execute
        If Not lvwTasks.SelectedItem Is Nothing Then
            HTidy.Run lvwTasks.SelectedItem.Tag
        End If
    ElseIf Index = 1 Then
        'new task
        frmRunTidy.Show vbModal
    ElseIf Index = 2 Then
        'edit task
        If Not lvwTasks.SelectedItem Is Nothing Then
            frmRunTidy.file = lvwTasks.SelectedItem.Tag
            frmRunTidy.task = lvwTasks.SelectedItem.Text
            frmRunTidy.Show vbModal
        End If
    ElseIf Index = 3 Then
        'remove task
        If Not lvwTasks.SelectedItem Is Nothing Then
            If Confirma("Are you sure to remove this task") = vbYes Then
                eliminar_tarea
            End If
        End If
    ElseIf Index = 4 Then
        'exit
        Unload Me
    Else
        'default
        ini = util.StripPath(App.Path) & "tidy\tasks.ini"
        If Not lvwTasks.SelectedItem Is Nothing Then
            For k = 1 To lvwTasks.ListItems.count
                If Len(lvwTasks.ListItems(k).SubItems(1).Caption) > 0 Then
                    On Error Resume Next
                    util.GrabaIni ini, "tasks", "tsk" & k, lvwTasks.ListItems(k).Text & "|" & lvwTasks.ListItems(k).Tag & "|"
                    Err = 0
                End If
                lvwTasks.ListItems(k).SubItems(1).Caption = vbNullString
            Next k
            k = lvwTasks.SelectedItem.Index
            On Error Resume Next
            util.GrabaIni ini, "tasks", "tsk" & k, lvwTasks.ListItems(k).Text & "|" & lvwTasks.ListItems(k).Tag & "|Default"
            Err = 0
            lvwTasks.SelectedItem.SubItems(1).Caption = "Default"
        End If
    End If
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
    
    util.Hourglass hwnd, True
    
    With lvwTasks
        .Columns.Add , "k1", "Task Name", , 3400
        .Columns.Add , "k2", "Default", , 1440
    End With
    
    Call cargar_tidy
        
    util.Hourglass hwnd, False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTidyConfig = Nothing
End Sub


Private Sub lvwTasks_ItemDblClick(ITem As vbalListViewLib6.cListItem)

    If Not ITem Is Nothing Then
        cmd_Click 0
    End If
    
End Sub

