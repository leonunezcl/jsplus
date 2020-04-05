VERSION 5.00
Begin VB.Form frmTip 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "JavaScript Plus! - Tip of the day"
   ClientHeight    =   3600
   ClientLeft      =   3210
   ClientTop       =   3420
   ClientWidth     =   6480
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmd 
      Caption         =   "&Next"
      Height          =   375
      Index           =   1
      Left            =   4260
      TabIndex        =   6
      Top             =   3090
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   5355
      TabIndex        =   5
      Top             =   3090
      Width           =   975
   End
   Begin VB.ComboBox cboTip 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3105
      Width           =   4035
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2865
      Index           =   1
      Left            =   105
      ScaleHeight     =   2835
      ScaleWidth      =   6225
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2865
         Index           =   0
         Left            =   0
         ScaleHeight     =   2865
         ScaleWidth      =   975
         TabIndex        =   3
         Top             =   0
         Width           =   975
         Begin VB.Image img 
            Appearance      =   0  'Flat
            Height          =   480
            Left            =   225
            Picture         =   "frmTip.frx":000C
            Top             =   270
            Width           =   480
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   960
         X2              =   6225
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know ...."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1155
         TabIndex        =   2
         Top             =   90
         Width           =   2070
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2280
         Left            =   1035
         TabIndex        =   1
         Top             =   495
         Width           =   5130
      End
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' La base de datos en memoria de sugerencias.
Dim Tips As New Collection

' Nombre del archivo de sugerencias
Const TIP_FILE = "TIPOFDAY.TXT"

' Índice en la colección de la sugerencia actualmente mostrada.
Dim CurrentTip As Long


Private Sub DoNextTip()

    ' Seleccionar una sugerencia aleatoriamente.
    CurrentTip = Int((Tips.count * Rnd) + 1)
    
    ' O recorrer secuencialmente las sugerencias

'    CurrentTip = CurrentTip + 1
'    If Tips.Count < CurrentTip Then
'        CurrentTip = 1
'    End If
    
    ' Mostrar.
    frmTip.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    
    Dim NextTip As String   ' Leer cada sugerencia desde archivo.
    Dim InFile As Integer   ' Descriptor para archivo.
    
    ' Obtener el siguiente descriptor de archivo libre.
    InFile = FreeFile
    
    ' Asegurarse de que se especifica un archivo.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Asegurarse de que el archivo existe antes de intentar abrirlo.
    If Not ArchivoExiste2(sFile) Then
        LoadTips = False
        Exit Function
    End If
    
    ' Leer la colección desde un archivo de texto.
    Open sFile For Input As InFile
        Do While Not EOF(InFile)
            Line Input #InFile, NextTip
            Tips.Add NextTip
        Loop
    Close InFile

    ' Mostrar una sugerencia aleatoriamente.
    DoNextTip
    
    LoadTips = True
    
End Function
Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Unload Me
    Else
        Call DoNextTip
    End If
    
End Sub

Private Sub Form_Load()
    
    Dim ShowAtStartup As String
    
    util.CenterForm Me
    util.Hourglass hwnd, True
    
    cboTip.AddItem "Always show tips at startup"
    cboTip.AddItem "Never show tips at startup"
    
    ' Ver si debemos mostrar al iniciar
    ShowAtStartup = util.LeeIni(IniPath, "tips", "show")
    If ShowAtStartup = "" Then ShowAtStartup = 0
            
    cboTip.ListIndex = ShowAtStartup
        
    ' Semilla aleatoria
    Randomize
    
    ' Leer el archivo de sugerencias y mostrar una sugerencia aleatoriamente.
    If LoadTips(util.StripPath(App.Path) & TIP_FILE) = False Then
        lblTipText.Caption = "de que no se ha encontrado el archivo " & TIP_FILE & vbCrLf & vbCrLf & _
           "Cree un archivo de texto llamado " & TIP_FILE & " con el Bloc de notas, con una sugerencia por línea. " & _
           "A continuación, colóquelo en el mismo directorio que la aplicación."
    End If

    util.Hourglass hwnd, False
    
    'DrawXPCtl Me
    
End Sub

Public Sub DisplayCurrentTip()
    If Tips.count > 0 Then
        lblTipText.Caption = Tips.ITem(CurrentTip)
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error Resume Next
    Call util.GrabaIni(IniPath, "tips", "show", cboTip.ListIndex)
    Call clear_memory(Me)
    Err = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTip = Nothing
End Sub


