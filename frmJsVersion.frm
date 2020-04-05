VERSION 5.00
Object = "{FCFAF346-DE8A-4FB6-8612-5000548EFDC7}#2.0#0"; "vbsListView6.ocx"
Begin VB.Form frmJsVersion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "JavaScript and Navigator versions "
   ClientHeight    =   3765
   ClientLeft      =   3600
   ClientTop       =   2205
   ClientWidth     =   6945
   Icon            =   "frmJsVersion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Javascript & Navigator Table"
      Height          =   2400
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   105
      Width           =   6840
      Begin vbalListViewLib6.vbalListViewCtl lvwjsv 
         Height          =   1980
         Left            =   75
         TabIndex        =   0
         Top             =   255
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   3493
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
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   5220
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4110
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmJsVersion.frx":000C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   660
      Left            =   75
      TabIndex        =   1
      Top             =   2550
      Width           =   6630
   End
End
Attribute VB_Name = "frmJsVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
    Unload Me
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Dim k As Integer
    Dim Num As Variant
    Dim Archivo As String
    Dim linea As String
    Archivo = IniPath
    
    Num = util.LeeIni(Archivo, "versions", "num")
    If Num = "" Then
        Exit Sub
    End If
    
    With lvwjsv
        .Columns.Add , "k1", "JavaScript version", , 2000
        .Columns.Add , "k2", "Navigator version", , 4000
    End With
    
    For k = 1 To Num
        linea = util.LeeIni(Archivo, "versions", "v" & k)
        If Len(linea) > 0 Then
            lvwjsv.ListItems.Add , "k" & k, linea
            linea = util.LeeIni(Archivo, "versions", "b" & k)
            If Len(linea) > 0 Then
                lvwjsv.ListItems("k" & k).SubItems(1).Caption = linea
            End If
        End If
    Next k
    
    util.CenterForm Me
    
    flat_lviews Me
    
    Debug.Print "load"
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmJsVersion = Nothing
End Sub


