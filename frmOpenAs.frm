VERSION 5.00
Begin VB.Form frmOpenAs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open File As ..."
   ClientHeight    =   1800
   ClientLeft      =   4095
   ClientTop       =   3525
   ClientWidth     =   5280
   Icon            =   "frmOpenAs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1440
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3300
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   1125
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   5085
      Begin VB.ComboBox cboFileExt 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   255
         Width           =   3645
      End
      Begin VB.Label lbldescrip 
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   615
         Width           =   3645
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Select file type :"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   285
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmOpenAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type eFileInfo
    wildcard As String
    description As String
End Type
Private arr_wildcards() As eFileInfo


Private Sub cboFileExt_Change()
    If cboFileExt.ListIndex <> -1 Then
        lbldescrip.Caption = arr_wildcards(cboFileExt.ListIndex + 1).description
    End If
End Sub

Private Sub cboFileExt_Click()
    If cboFileExt.ListIndex <> -1 Then
        lbldescrip.Caption = arr_wildcards(cboFileExt.ListIndex + 1).description
    End If
End Sub


Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If cboFileExt.ListIndex <> -1 Then
            ListaLangs.TempExtension = cboFileExt.Text
        End If
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()

    util.CenterForm Me
        
    ReDim arr_wildcards(0)
    
    Call cargar_wildcards
        
End Sub


Private Sub cargar_wildcards()

    Dim inifile As String
    Dim V As Variant
    Dim k As Integer
    Dim j As Integer
    Dim valor As String
    
    inifile = util.StripPath(App.Path) & "filelist.ini"
    
    j = 1
    If ArchivoExiste2(inifile) Then
        V = util.LeeIni(inifile, "filelist", "num")
        For k = 1 To V
            valor = util.LeeIni(inifile, "filelist", "ele" & k)
            
            If util.Explode(valor, 2, "|") <> "*.*" Then
                ReDim Preserve arr_wildcards(j)
                arr_wildcards(j).wildcard = util.Explode(valor, 2, "|")
                arr_wildcards(j).description = util.Explode(valor, 1, "|")
                cboFileExt.AddItem arr_wildcards(j).wildcard
                j = j + 1
            End If
        Next k
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload : " & Me.Name
    Set frmOpenAs = Nothing
End Sub


