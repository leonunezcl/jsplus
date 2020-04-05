VERSION 5.00
Object = "{FCFAF346-DE8A-4FB6-8612-5000548EFDC7}#2.0#0"; "vbsListView6.ocx"
Begin VB.Form frmResWord 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Javascript Help"
   ClientHeight    =   4575
   ClientLeft      =   2835
   ClientTop       =   3645
   ClientWidth     =   7200
   Icon            =   "frmResWord.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin vbalListViewLib6.vbalListViewCtl lvwresw 
      Height          =   3705
      Left            =   45
      TabIndex        =   0
      Top             =   225
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   6535
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
      FullRowSelect   =   -1  'True
      AutoArrange     =   0   'False
      Appearance      =   0
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reserved Words"
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
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   1425
   End
End
Attribute VB_Name = "frmResWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CargarPalabras()

    Dim ini As String
    Dim arr_key() As String
    Dim k As Integer
    Dim C As Integer
    Dim j As Integer
    
    ini = util.StripPath(App.Path) & "config\jshelp.ini"
    
    If Not ArchivoExiste2(ini) Then
        MsgBox "File : " & ini & " doesn't exists", vbCritical
        Exit Sub
    End If
    
    get_info_section "language", arr_key, ini
    
    C = 1
    j = 1
    For k = 2 To UBound(arr_key)
        If C = 1 Then
            lvwresw.ListItems.Add , "k" & j, Explode(arr_key(k), 1, "#")
            C = C + 1
        ElseIf C = 2 Then
            lvwresw.ListItems(j).SubItems(1).Caption = Explode(arr_key(k), 1, "#")
            C = C + 1
        ElseIf C = 3 Then
            lvwresw.ListItems(j).SubItems(2).Caption = Explode(arr_key(k), 1, "#")
            C = C + 1
        ElseIf C = 4 Then
            lvwresw.ListItems(j).SubItems(3).Caption = Explode(arr_key(k), 1, "#")
            C = 1
            j = j + 1
        End If
    Next k
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Unload Me
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.Hourglass hwnd, True
    
    util.CenterForm Me
        
    With lvwresw
        .Columns.Add , "k1", "Keyword", , 1700
        .Columns.Add , "k2", "Keyword", , 1700
        .Columns.Add , "k3", "Keyword", , 1800
        .Columns.Add , "k4", "Keyword", , 1800
    End With
    
    Call CargarPalabras
    
    util.Hourglass hwnd, False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmResWord = Nothing
End Sub


