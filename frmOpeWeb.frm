VERSION 5.00
Begin VB.Form frmOpeWeb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open URL"
   ClientHeight    =   1455
   ClientLeft      =   3300
   ClientTop       =   4875
   ClientWidth     =   7605
   Icon            =   "frmOpeWeb.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4530
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   2010
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Open URL"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7425
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmOpeWeb.frx":000C
         Left            =   2130
         List            =   "frmOpeWeb.frx":000E
         TabIndex        =   1
         Top             =   255
         Width           =   5130
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Enter address with http://"
         Height          =   195
         Left            =   165
         TabIndex        =   2
         Top             =   300
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmOpeWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SHAutoComplete Lib "shlwapi.dll" (ByVal hWndEdit As Long, ByVal dwFlags As AutoCompleteFlags) As Integer
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Enum AutoCompleteFlags
    SHACF_DEFAULT = &H0
    SHACF_FILESYSTEM = &H1
    SHACF_URLHISTORY = &H2
    SHACF_URLMRU = &H4
    SHACF_USETAB = &H8
    SHACF_URLALL = (SHACF_URLHISTORY Or SHACF_URLMRU)
    SHACF_FILESYS_ONLY = &H10
    SHACF_FILESYS_DIRS = &H20
    SHACF_AUTOSUGGEST_FORCE_ON = &H10000000
    SHACF_AUTOSUGGEST_FORCE_OFF = &H20000000
    SHACF_AUTOAPPEND_FORCE_ON = &H40000000
    SHACF_AUTOAPPEND_FORCE_OFF = &H80000000
End Enum
Private Sub SetAutoCompleteComboBox(ByVal lngHwnd As Long)

    ' Thanks go to enmity for this fix
    Dim o_hwndEdit As Long
    o_hwndEdit = FindWindowEx(lngHwnd, 0, "EDIT", vbNullString)

    If o_hwndEdit <> 0 Then
        SetAutoCompleteTextBox o_hwndEdit
    End If

End Sub


Private Sub SetAutoCompleteTextBox(ByVal lngHwnd As Long)

    SHAutoComplete lngHwnd, SHACF_DEFAULT
    
End Sub



Private Sub cmd_Click(Index As Integer)

    Dim webfile As String
    
    If Index = 0 Then
        If Combo1.Text <> "" Then
            webfile = Combo1.Text
            Me.Hide
            
            util.Hourglass hwnd, True
            
            Load frmOpenFiles
            frmOpenFiles.Caption = "Opening URL"
            frmOpenFiles.lblFile.Caption = webfile
            frmOpenFiles.pgb.Max = 1
            frmOpenFiles.Show
            
            frmMain.opeweb webfile
            
            util.Hourglass hwnd, False
            
            Unload frmOpenFiles
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
    
    'Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    'cmd(0).Refresh
    'cmd(1).Refresh
    
    SetAutoCompleteComboBox Combo1.hwnd
    
    Debug.Print "load : " & Me.Name
    
    'DrawXPCtl Me
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload : " & Me.Name
    Set frmOpeWeb = Nothing
End Sub
