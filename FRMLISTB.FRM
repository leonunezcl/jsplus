VERSION 5.00
Object = "{246E535D-09D2-4109-80DA-2FF183F4D185}#2.1#0"; "colorpick.ocx"
Object = "{FCFAF346-DE8A-4FB6-8612-5000548EFDC7}#2.0#0"; "vbsListView6.ocx"
Begin VB.Form frmHtmlListbox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listbox"
   ClientHeight    =   6270
   ClientLeft      =   5115
   ClientTop       =   2175
   ClientWidth     =   6495
   Icon            =   "FRMLISTB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6270
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&More ..."
      Height          =   375
      Index           =   6
      Left            =   4920
      TabIndex        =   25
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Events"
      Height          =   375
      Index           =   5
      Left            =   3360
      TabIndex        =   24
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   20
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Extras"
      Height          =   1005
      Index           =   2
      Left            =   45
      TabIndex        =   14
      Top             =   1785
      Width           =   6375
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   255
         Index           =   8
         Left            =   5760
         TabIndex        =   26
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1635
         TabIndex        =   5
         Top             =   600
         Width           =   4080
      End
      Begin VB.TextBox txtColorText 
         Height          =   285
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   885
      End
      Begin ColorPick.ClrPicker ClrPicker1 
         Height          =   300
         Left            =   2550
         TabIndex        =   17
         Top             =   270
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Background Image"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   16
         Top             =   585
         Width           =   1350
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Background Color"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   285
         Width           =   1275
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Events"
      Height          =   690
      Index           =   3
      Left            =   45
      TabIndex        =   12
      Top             =   4950
      Width           =   6360
      Begin VB.CommandButton cmd 
         Caption         =   "&Delete"
         Height          =   375
         Index           =   4
         Left            =   5640
         TabIndex        =   23
         Top             =   240
         Width           =   615
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
         TabIndex        =   13
         Top             =   285
         Width           =   510
      End
   End
   Begin VB.Frame fra 
      Caption         =   "ListBox Items"
      Height          =   2115
      Index           =   1
      Left            =   45
      TabIndex        =   11
      Top             =   2820
      Width           =   6360
      Begin VB.CommandButton cmd 
         Caption         =   "&Remove"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   22
         Top             =   1680
         Width           =   1215
      End
      Begin vbalListViewLib6.vbalListViewCtl lvwchk 
         Height          =   1395
         Left            =   105
         TabIndex        =   18
         Top             =   195
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   2461
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
         FlatScrollBar   =   -1  'True
         HeaderButtons   =   0   'False
         HeaderTrackSelect=   0   'False
         HideSelection   =   0   'False
         InfoTips        =   0   'False
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   1710
      Index           =   0
      Left            =   45
      TabIndex        =   7
      Top             =   60
      Width           =   6375
      Begin VB.CommandButton cmd 
         Caption         =   "&Add"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   795
         TabIndex        =   0
         Top             =   270
         Width           =   5475
      End
      Begin VB.TextBox txtValor 
         Height          =   285
         Left            =   795
         TabIndex        =   2
         Top             =   960
         Width           =   5460
      End
      Begin VB.TextBox txtTam 
         Height          =   285
         Left            =   795
         TabIndex        =   1
         Top             =   615
         Width           =   495
      End
      Begin VB.CheckBox chk 
         Caption         =   "Multiple Select"
         Height          =   240
         Left            =   795
         TabIndex        =   3
         Top             =   1305
         Width           =   1695
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   285
         Width           =   420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Top             =   1035
         Width           =   405
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   8
         Top             =   630
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmHtmlListbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Agregar()

    Dim k As Integer
    
    If txtNombre.Text = "" Then
        MsgBox "The name for this element.", vbCritical
        txtNombre.SetFocus
        Exit Sub
    End If
    
    If txtValor.Text = "" Then
        MsgBox "The value for this element.", vbCritical
        txtValor.SetFocus
        Exit Sub
    End If
    
    k = lvwchk.ListItems.count + 1
    lvwchk.ListItems.Add , "k" & k, txtNombre.Text
    lvwchk.ListItems("k" & k).SubItems(1).Caption = txtValor.Text
    
    txtValor.SetFocus
    
End Sub

Private Sub Eliminar()

    If lvwchk.ListItems.count <= 0 Then
        MsgBox "Nothing to do.", vbCritical
        Exit Sub
    End If
     
    If Not lvwchk.SelectedItem Is Nothing Then
        lvwchk.ListItems.Remove lvwchk.SelectedItem.key
    End If
   
End Sub

Private Function InsertarListbox() As Boolean

    On Error GoTo ErrorInsertarListbox

    Dim k As Integer
    Dim src As New cStringBuilder
    Dim stam As String
    
    If lvwchk.ListItems.count <= 0 Then
        MsgBox "Nothing to do.", vbCritical
        Exit Function
    End If
    
    If txtTam.Text = "" Then
        stam = "3"
    Else
        stam = txtTam.Text
    End If
    
    src.Append "<select name=" & Chr$(34) & "listbox1" & Chr$(34)
    src.Append " size=" & Chr$(34) & stam & Chr$(34)
    
    If txtColorText.Text <> "" Then
        src.Append " style=" & Chr$(34) & "background-color:" & txtColorText.Text & Chr$(34)
    End If
    
    If Text1.Text <> "" Then
        src.Append " style=" & Chr$(34) & "background-image:url('" & Text1.Text & "')" & Chr$(34)
    End If
    
    If chk.Value Then
        src.Append " multiple"
    End If
    
    If Len(Trim$(CComHtmlAttrib.Output)) > 0 Then
        src.Append " " & Trim$(CComHtmlAttrib.Output)
    End If
    
    If Len(Trim$(CEventos.Output)) > 0 Then
        src.Append " " & Trim$(CEventos.Output)
    End If
    src.Append ">" & vbNewLine
        
    For k = 1 To lvwchk.ListItems.count
        src.Append "<option value=" & Chr$(34) & lvwchk.ListItems(k).SubItems(1).Caption & Chr$(34) & ">"
        src.Append lvwchk.ListItems(k).Text & "</option>" & vbNewLine
    Next k
    src.Append "</select>"
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(src.ToString)
    End If
    
    Set src = Nothing
    
    InsertarListbox = True
    
    Exit Function
    
ErrorInsertarListbox:
    MsgBox "InsertarListbox : " & Err & " " & Error$, vbCritical
    
End Function


Private Sub ClrPicker1_ColorSelected(m_Color As stdole.OLE_COLOR, m_Code As String)
    txtColorText.Text = m_Code
    txtColorText.Tag = m_Color
End Sub


Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If InsertarListbox() Then
            Unload Me
        End If
    ElseIf Index = 1 Then
        Unload Me
    ElseIf Index = 2 Then
        Call Agregar
    ElseIf Index = 3 Then
        Call Eliminar
    ElseIf Index = 4 Then
        If cboEvents.ListIndex <> -1 Then
            CEventos.Remove cboEvents.Text
            cboEvents.RemoveItem cboEvents.ListIndex
        End If
    ElseIf Index = 5 Then
        frmEvents.html_tag = "ListBox"
        frmEvents.Show vbModal
        Call CEventos.Attach(Me.cboEvents)
    ElseIf Index = 6 Then
        frmCommonHtml.html_tag = "ListBox"
        frmCommonHtml.Show vbModal
    ElseIf Index = 7 Then
        
    ElseIf Index = 8 Then
        Dim glosa As String
        Dim Archivo As String
        Dim LastPath As String
        
        glosa = "CompuServe Graphics Interchange (*.gif)|*.gif|"
        glosa = glosa & "JPG (*.jpg)|*.jpg|"
        glosa = glosa & "All files (*.*)|*.*"
        
        LastPath = App.Path
        
        If Not Cdlg.VBGetOpenFileName(Archivo, , , , , , glosa, , LastPath, "Open Image", "gif") Then
            Exit Sub
        End If
        Text1.Text = util.VBArchivoSinPath(Archivo)
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
    
    util.SetNumber txtTam.hwnd
    
    With lvwchk
        .Columns.Add , "k1", "Name", , 1440
        .Columns.Add , "k2", "Value", , 4000
    End With
    
    flat_lviews Me
        
    CEventos.Clear
    CComHtmlAttrib.Clear
    
    ClrPicker1.PathPaleta = App.Path & "\pal\256c.pal"
    
    Debug.Print "load"
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmHtmlListbox = Nothing
End Sub


