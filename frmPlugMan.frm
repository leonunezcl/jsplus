VERSION 5.00
Object = "{FCFAF346-DE8A-4FB6-8612-5000548EFDC7}#2.0#0"; "vbsListView6.ocx"
Begin VB.Form frmPlugMan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add-In Manager ..."
   ClientHeight    =   3675
   ClientLeft      =   1590
   ClientTop       =   3540
   ClientWidth     =   11535
   Icon            =   "frmPlugMan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   2
      Left            =   10200
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "R&emove"
      Height          =   375
      Index           =   1
      Left            =   10200
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Refresh"
      Height          =   375
      Index           =   4
      Left            =   10200
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Start"
      Height          =   375
      Index           =   0
      Left            =   10200
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin vbalListViewLib6.vbalListViewCtl lvwplug 
      Height          =   3090
      Left            =   75
      TabIndex        =   0
      Top             =   255
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   5450
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
      LabelEdit       =   0   'False
      FullRowSelect   =   -1  'True
      AutoArrange     =   0   'False
      Appearance      =   0
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Don't use unknown add-in if you are unsure in the Author of the add-in or in the add-in itself."
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
      Left            =   75
      TabIndex        =   2
      Top             =   3390
      Width           =   7890
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Installed Add-Ins"
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
      Left            =   75
      TabIndex        =   1
      Top             =   60
      Width           =   1440
   End
End
Attribute VB_Name = "frmPlugMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub plugin_start()

    If Not lvwplug.SelectedItem Is Nothing Then
        Call Plugins.Run(lvwplug.SelectedItem.Text, frmMain)
    End If
    
End Sub

Private Sub plugin_remove()

    If Not lvwplug.SelectedItem Is Nothing Then
        Dim Msg As String
        
        Msg = "Are you sure to remove the selected add-in"
        
        If Confirma(Msg) = vbYes Then
            MsgBox "Before proceeding the add-in deletion, please unregister your dll be useing regsvr32 /u [addinname.dll]", vbInformation
            Plugins.Remove lvwplug.SelectedItem.Text
            Plugins.Load (App.Path)
            cargar_plugins
        End If
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        'start
        plugin_start
    ElseIf Index = 1 Then
        'remove
        plugin_remove
    ElseIf Index = 2 Then
        'cancel
        Unload Me
    'ElseIf Index = 3 Then
    '    frmNewPlug.Show vbModal
    ElseIf Index = 4 Then
        Plugins.Load (App.Path)
        cargar_plugins
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
    
    With lvwplug
        .Columns.Add , "Caption", "Caption", , 2000
        .Columns.Add , "Description", "Description", , 4500
        .Columns.Add , "Version", "Version", , 1440
        .Columns.Add , "ClassId", "ClassId", , 1440
        .Columns.Add , "Autor", "Autor", , 2000
        .ImageList = frmMain.m_MainImg
    End With
    
    Call cargar_plugins
    
End Sub
Private Sub cargar_plugins()

    Dim Plugin As cPlugin
    Dim k As Integer
    
    lvwplug.ListItems.Clear
    
    For k = 1 To Plugins.count
        Set Plugin = New cPlugin
        Set Plugin = Plugins.Plugins.ITem(k)
        lvwplug.ListItems.Add , "k" & k, Plugin.Caption, 147, 147
        lvwplug.ListItems(k).SubItems(1).Caption = Plugin.description
        lvwplug.ListItems(k).SubItems(2).Caption = Plugin.Version
        lvwplug.ListItems(k).SubItems(3).Caption = Plugin.ClassId
        lvwplug.ListItems(k).SubItems(4).Caption = Plugin.Autor
        Set Plugin = Nothing
    Next k
            
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPlugMan = Nothing
End Sub


Private Sub lvwplug_ItemDblClick(ITem As vbalListViewLib6.cListItem)
    If Not ITem Is Nothing Then
        cmd_Click 0
    End If
End Sub


