VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMultiAssign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multi ...."
   ClientHeight    =   4875
   ClientLeft      =   4320
   ClientTop       =   2535
   ClientWidth     =   4920
   ControlBox      =   0   'False
   Icon            =   "frmMultiAssign.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   255
      Width           =   3480
   End
   Begin MSComctlLib.ListView lvwSelFun 
      Height          =   4215
      Left            =   30
      TabIndex        =   2
      Top             =   615
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Member"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Select Member Type"
      Height          =   195
      Left            =   30
      TabIndex        =   1
      Top             =   45
      Width           =   1470
   End
End
Attribute VB_Name = "frmMultiAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub hacer_multi()

    Dim k As Integer
    
    util.Hourglass hwnd, True
    
    With lvwSelFun
        For k = 1 To .ListItems.count
            If cboType.ListIndex = 0 Then
                lvwSelFun.ListItems(k).SubItems(1) = "Object"
            ElseIf cboType.ListIndex = 1 Then
                lvwSelFun.ListItems(k).SubItems(1) = "Property"
            ElseIf cboType.ListIndex = 2 Then
                lvwSelFun.ListItems(k).SubItems(1) = "Method"
            ElseIf cboType.ListIndex = 3 Then
                lvwSelFun.ListItems(k).SubItems(1) = "Event"
            ElseIf cboType.ListIndex = 4 Then
                lvwSelFun.ListItems(k).SubItems(1) = "Constant"
            ElseIf cboType.ListIndex = 5 Then
                lvwSelFun.ListItems(k).SubItems(1) = "Collection"
            End If
            
            Call frmImportLibrary.ActualizaInfo(cboType.ListIndex, util.Explode(lvwSelFun.ListItems(k).Tag, 1, "|"), util.Explode(lvwSelFun.ListItems(k).Tag, 2, "|"))
        Next k
    End With
    util.Hourglass hwnd, False
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 1 Then
        If cboType.ListIndex <> -1 Then
            Call hacer_multi
        Else
            Exit Sub
        End If
    End If
    
    Unload Me
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    util.CenterForm Me
        
    cboType.AddItem "Object"
    cboType.AddItem "Property"
    cboType.AddItem "Method"
    cboType.AddItem "Event"
    cboType.AddItem "Constant"
    cboType.AddItem "Collection"
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload " & Me.Name
    Set frmMultiAssign = Nothing
End Sub


