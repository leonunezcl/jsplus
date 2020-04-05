VERSION 5.00
Begin VB.Form frmScript 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Script"
   ClientHeight    =   2370
   ClientLeft      =   3660
   ClientTop       =   2505
   ClientWidth     =   5640
   Icon            =   "frmScript.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Configure ..."
      Height          =   1695
      Left            =   45
      TabIndex        =   4
      Top             =   105
      Width           =   5550
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   5145
         TabIndex        =   10
         Top             =   990
         Width           =   300
      End
      Begin VB.CheckBox chk 
         Caption         =   "Run at server"
         Height          =   210
         Left            =   1380
         TabIndex        =   3
         Top             =   1350
         Width           =   1395
      End
      Begin VB.TextBox txtArchivo 
         Height          =   285
         Left            =   1380
         TabIndex        =   2
         Top             =   975
         Width           =   3720
      End
      Begin VB.ComboBox cboScript 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   615
         Width           =   4095
      End
      Begin VB.ComboBox cboLanguage 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1380
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   255
         Width           =   4095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Script File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   7
         Top             =   1005
         Width           =   690
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Script Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Top             =   645
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Language"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   300
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub insertar_script()

    Dim src As New cStringBuilder
    
    src.Append "<script language=" & Chr$(34) & cboLanguage.Text & Chr$(34)
    src.Append " type=" & Chr$(34) & cboScript.Text & Chr$(34)
    If Len(txtArchivo.Text) > 0 Then
        src.Append " src=" & Chr$(34) & Replace(txtArchivo.Text, "\", "/") & Chr$(34)
        src.Append "></script>" & vbNewLine
    ElseIf chk.Value Then
        src.Append " runat=" & Chr$(34) & "server" & Chr$(34) & ">" & vbNewLine
        src.Append "" & vbNewLine
        src.Append "</script>" & vbNewLine
    Else
        src.Append ">" & vbNewLine
        src.Append "" & vbNewLine
        src.Append "</script>" & vbNewLine
    End If
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(src.ToString)
    End If
    
    Set src = Nothing
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Dim Archivo As String
    Dim glosa As String
    If Index = 0 Then
        Call insertar_script
        Unload Me
    ElseIf Index = 1 Then
        Unload Me
    Else
        glosa = strGlosa()
        If Cdlg.VBGetOpenFileName(Archivo, , , , , , glosa, , , "Select file ...", "js") Then
            txtArchivo.Text = Archivo
        End If
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Dim j As Integer
    Dim Archivo As String
    Dim sSections() As String
    
    util.CenterForm Me
        
    Archivo = IniPath
    
    get_info_section "script_language", sSections, Archivo
    
    For j = 2 To UBound(sSections)
        cboLanguage.AddItem sSections(j) 'Util.LeeIni(Archivo, "script_language", "ele" & j)
    Next j
    cboLanguage.ListIndex = 0
    
    get_info_section "script_type", sSections, Archivo
    
    For j = 2 To UBound(sSections)
        cboScript.AddItem sSections(j) 'Util.LeeIni(Archivo, "script_type", "ele" & j)
    Next j
    cboScript.ListIndex = 0
        
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmScript = Nothing
End Sub


