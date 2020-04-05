VERSION 5.00
Object = "{FCFAF346-DE8A-4FB6-8612-5000548EFDC7}#2.0#0"; "vbsListView6.ocx"
Begin VB.Form frmHtmlCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CheckBox"
   ClientHeight    =   5415
   ClientLeft      =   4965
   ClientTop       =   3030
   ClientWidth     =   6405
   Icon            =   "FRMCHECK.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5415
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&More"
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   17
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Events"
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   16
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   15
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Events"
      Height          =   690
      Index           =   3
      Left            =   45
      TabIndex        =   11
      Top             =   4110
      Width           =   6330
      Begin VB.CommandButton cmd 
         Caption         =   "&Delete"
         Height          =   285
         Index           =   4
         Left            =   5595
         TabIndex        =   18
         Top             =   255
         Width           =   660
      End
      Begin VB.ComboBox cboEvents 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   705
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   255
         Width           =   4860
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   12
         Top             =   285
         Width           =   510
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Preview"
      Height          =   1965
      Index           =   2
      Left            =   45
      TabIndex        =   10
      Top             =   1455
      Width           =   6330
      Begin VB.CommandButton cmd 
         Caption         =   "&Remove"
         Height          =   315
         Index           =   6
         Left            =   5400
         TabIndex        =   20
         Top             =   1560
         Width           =   735
      End
      Begin vbalListViewLib6.vbalListViewCtl lvwchk 
         Height          =   1260
         Left            =   75
         TabIndex        =   3
         Top             =   240
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   2223
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
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         AutoArrange     =   0   'False
         Appearance      =   0
         FlatScrollBar   =   -1  'True
         HeaderButtons   =   0   'False
         HeaderTrackSelect=   0   'False
         HideSelection   =   0   'False
         InfoTips        =   0   'False
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   1335
      Index           =   0
      Left            =   45
      TabIndex        =   8
      Top             =   75
      Width           =   6330
      Begin VB.CommandButton cmd 
         Caption         =   "&Add"
         Height          =   330
         Index           =   5
         Left            =   2760
         TabIndex        =   19
         Top             =   930
         Width           =   975
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   705
         TabIndex        =   0
         Top             =   270
         Width           =   5520
      End
      Begin VB.TextBox txtValue 
         Height          =   285
         Left            =   705
         TabIndex        =   1
         Top             =   585
         Width           =   5520
      End
      Begin VB.CheckBox chk 
         Caption         =   "Checked"
         Height          =   195
         Left            =   690
         TabIndex        =   2
         Top             =   915
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   585
         Width           =   405
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Insert Mode"
      Height          =   615
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   3465
      Width           =   6330
      Begin VB.OptionButton opt 
         Caption         =   "Columns"
         Height          =   240
         Index           =   1
         Left            =   3765
         TabIndex        =   5
         Top             =   270
         Width           =   1080
      End
      Begin VB.OptionButton opt 
         Caption         =   "Rows"
         Height          =   240
         Index           =   0
         Left            =   1590
         TabIndex        =   4
         Top             =   270
         Value           =   -1  'True
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmHtmlCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tipo_control As String
Private Sub Agregar()

    Dim k As Integer
    
    If txtName.Text = "" Then
        MsgBox "A name for this control.", vbCritical
        txtName.SetFocus
        Exit Sub
    End If
    
    If txtValue.Text = "" Then
        MsgBox "A value for this control.", vbCritical
        txtValue.SetFocus
        Exit Sub
    End If
    
    k = lvwchk.ListItems.count + 1
    lvwchk.ListItems.Add , "k" & k, txtName.Text
    lvwchk.ListItems("k" & k).SubItems(1).Caption = txtValue.Text
    If chk.Value Then
        lvwchk.ListItems("k" & k).SubItems(2).Caption = "Yes"
    Else
        lvwchk.ListItems("k" & k).SubItems(2).Caption = "No"
    End If
    
End Sub

Private Sub Eliminar()

    If lvwchk.ListItems.count > 0 Then
        If Not lvwchk.SelectedItem Is Nothing Then
            lvwchk.ListItems.Remove lvwchk.SelectedItem.key
        End If
    End If
    
End Sub

Private Function InsertarChk() As Boolean

    On Error GoTo ErrorInsertarChk
    
    Dim src As New cStringBuilder
    Dim k As Integer
    Dim j As Integer
    Dim nombre As String
    Dim valor As String
    Dim sel As String
    
    If lvwchk.ListItems.count = 0 Then
        MsgBox "Nothing to do.", vbCritical
        Exit Function
    End If
    
    If opt(0).Value Then
        j = 0
    Else
        j = 1
    End If
        
    For k = 1 To lvwchk.ListItems.count
        nombre = lvwchk.ListItems(k).Text
        valor = lvwchk.ListItems(k).SubItems(1).Caption
        sel = lvwchk.ListItems(k).SubItems(2).Caption
        
        src.Append "<input type=" & Chr$(34) & tipo_control & Chr$(34)
        
        'agregar atributos especiales
        If Len(Trim$(CComHtmlAttrib.Output)) > 0 Then
            src.Append CComHtmlAttrib.Output
        End If
        
        src.Append " name=" & Chr$(34) & nombre & Chr$(34)
        src.Append " value=" & Chr$(34) & valor & Chr$(34)
        If sel = "Yes" Then
            src.Append " checked"
        End If
        
        'agregar los eventos
        If Len(Trim$(CEventos.Output)) > 0 Then
            src.Append CEventos.Output
        End If
        
        src.Append ">" & valor
        If j = 0 Then
            src.Append "<br>" & vbNewLine
        End If
    Next k
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(src.ToString)
    End If
    
    InsertarChk = True
    
    Set src = Nothing
    
    Exit Function
    
ErrorInsertarChk:
    MsgBox "InsertarChk : " & Err & " " & Error$, vbCritical
    
End Function

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If InsertarChk() Then
            Unload Me
        End If
    ElseIf Index = 1 Then
        Unload Me
    ElseIf Index = 2 Then
        frmEvents.html_tag = tipo_control
        frmEvents.Show vbModal
        Call CEventos.Attach(Me.cboEvents)
    ElseIf Index = 3 Then
        frmCommonHtml.html_tag = tipo_control
        frmCommonHtml.Show vbModal
    ElseIf Index = 4 Then
        If cboEvents.ListIndex <> -1 Then
            CEventos.Remove cboEvents.Text
            cboEvents.RemoveItem cboEvents.ListIndex
        End If
    ElseIf Index = 5 Then
        Call Agregar
    ElseIf Index = 6 Then
        Call Eliminar
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
   
    With lvwchk
        .Columns.Add , "k1", "Name", , 1440
        .Columns.Add , "k2", "Value", , 1440
        .Columns.Add , "k3", "Checked", , 1440
    End With
    
    Me.Caption = "Create " & tipo_control
    
    flat_lviews Me
    
    CEventos.Clear
    CComHtmlAttrib.Clear
    
    Debug.Print "load"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Debug.Print "unload"
    Set frmHtmlCheck = Nothing
    
End Sub


