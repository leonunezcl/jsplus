VERSION 5.00
Begin VB.Form frmInstallHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Install JavaScript Reference & Ajax Libraries"
   ClientHeight    =   5700
   ClientLeft      =   4755
   ClientTop       =   3270
   ClientWidth     =   5760
   ControlBox      =   0   'False
   Icon            =   "frmInstallHelp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Install Ajax Libraries"
      Height          =   495
      Index           =   2
      Left            =   3720
      TabIndex        =   15
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   495
      Index           =   1
      Left            =   3720
      TabIndex        =   14
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Install JavaScript Reference"
      Height          =   495
      Index           =   0
      Left            =   3720
      TabIndex        =   13
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox picAjax 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   270
      ScaleHeight     =   1110
      ScaleWidth      =   2925
      TabIndex        =   10
      Top             =   3690
      Visible         =   0   'False
      Width           =   2955
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Installing ..."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1050
         TabIndex        =   12
         Top             =   255
         Width           =   810
      End
      Begin VB.Label lblzipajax 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "file"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   75
         TabIndex        =   11
         Top             =   645
         Width           =   2775
      End
   End
   Begin VB.ListBox lstAjax 
      Height          =   2205
      Left            =   30
      TabIndex        =   8
      Top             =   3105
      Width           =   3510
   End
   Begin VB.PictureBox pichelp 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   300
      ScaleHeight     =   1110
      ScaleWidth      =   2925
      TabIndex        =   5
      Top             =   1035
      Visible         =   0   'False
      Width           =   2955
      Begin VB.Label lblzip 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "file"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   75
         TabIndex        =   7
         Top             =   645
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Installing ..."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1050
         TabIndex        =   6
         Top             =   255
         Width           =   810
      End
   End
   Begin VB.PictureBox pic 
      Height          =   330
      Left            =   30
      ScaleHeight     =   270
      ScaleWidth      =   5880
      TabIndex        =   3
      Top             =   5355
      Width           =   5940
      Begin VB.Label lblfile 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ready"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   4
         Top             =   45
         Width           =   465
      End
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   7950
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6045
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.ListBox lstFiles 
      Height          =   2595
      Left            =   30
      TabIndex        =   1
      Top             =   270
      Width           =   3510
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ajax Libraries"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   30
      TabIndex        =   9
      Top             =   2895
      Width           =   930
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "JavaScript reference && Core Guide"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   2445
   End
End
Attribute VB_Name = "frmInstallHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_cUnzip As cUnzip
Attribute m_cUnzip.VB_VarHelpID = -1

Private arr_paths() As String
Private arr_paths_ajax() As String
Private Function install_files(ByVal opc As Integer) As Boolean

    On Error GoTo Errorinstall_files
    
    Dim ret As Boolean
    Dim k As Integer
    Dim j As Integer
    Dim Path As String
    Dim sFolder As String
    
    util.Hourglass hwnd, True
    
    Path = util.StripPath(App.Path)
    
    ret = True
    
    pichelp.Visible = False
    picAjax.Visible = False
    
    If opc = 1 Then
        lblzip.Caption = vbNullString
        pichelp.Visible = True
        
        For k = 1 To UBound(arr_paths)
            DoEvents
            lblzip.Caption = util.VBArchivoSinPath(Path & arr_paths(k))
            If ArchivoExiste2(Path & arr_paths(k)) Then
                ' Get the file directory:
                m_cUnzip.ZipFile = Path & arr_paths(k)
                m_cUnzip.OverwriteExisting = True
                m_cUnzip.UseFolderNames = True
                m_cUnzip.Directory
        
                If m_cUnzip.FileCount > 0 Then
                    For j = 1 To m_cUnzip.FileCount
                        m_cUnzip.FileSelected(j) = True
                    Next j
                    
                    sFolder = util.PathArchivo(Path & arr_paths(k))
                    m_cUnzip.UnzipFolder = sFolder
                    m_cUnzip.Unzip
                End If
            Else
                MsgBox "File :" & Path & arr_paths(k) & " was not found", vbCritical
                ret = False
                Exit For
            End If
            DoEvents
        Next k
        
        pichelp.Visible = False
    Else
        lblzipajax.Caption = vbNullString
        picAjax.Visible = True
        
        For k = 1 To UBound(arr_paths_ajax)
            DoEvents
            lblzipajax.Caption = util.VBArchivoSinPath(Path & arr_paths_ajax(k))
            If ArchivoExiste2(Path & arr_paths_ajax(k)) Then
                ' Get the file directory:
                m_cUnzip.ZipFile = Path & arr_paths_ajax(k)
                m_cUnzip.OverwriteExisting = True
                m_cUnzip.UseFolderNames = True
                m_cUnzip.Directory
        
                If m_cUnzip.FileCount > 0 Then
                    For j = 1 To m_cUnzip.FileCount
                        m_cUnzip.FileSelected(j) = True
                    Next j
                    
                    sFolder = util.PathArchivo(Path & arr_paths_ajax(k))
                    m_cUnzip.UnzipFolder = sFolder
                    m_cUnzip.Unzip
                End If
            Else
                MsgBox "File :" & Path & arr_paths_ajax(k) & " was not found", vbCritical
                ret = False
                Exit For
            End If
            DoEvents
        Next k
        
        picAjax.Visible = False
    End If
    
    util.Hourglass hwnd, False
    
    util.GrabaIni IniPath, "reference", "install", "1"
    
    install_files = ret
    
    Exit Function
    
Errorinstall_files:
    pichelp.Visible = False
    util.Hourglass hwnd, False
    MsgBox "install_files : " & Err & " " & Error$, vbAbortRetryIgnore
    
End Function
Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If install_files(1) Then
            MsgBox "JavaScript Reference was installed.", vbInformation
        Else
            MsgBox "Failed to install JavaScript Reference.", vbCritical
        End If
    ElseIf Index = 2 Then
        If install_files(2) Then
            MsgBox "Ajax Libraries was installed.", vbInformation
        Else
            MsgBox "Failed to install Ajax Libraries.", vbCritical
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
    
    ReDim arr_paths(7)
    ReDim arr_paths_ajax(8)
    
    Set m_cUnzip = New cUnzip
    
    lstFiles.AddItem "JavaScript Reference 1.3"
    lstFiles.AddItem "JavaScript Reference 1.4"
    lstFiles.AddItem "JavaScript Reference 1.5"
    lstFiles.AddItem "JavaScript Guide 1.3"
    lstFiles.AddItem "JavaScript Guide 1.4"
    lstFiles.AddItem "JavaScript Guide 1.5"
    lstFiles.AddItem "DOM Reference"
    
    lstAjax.AddItem "Aflax"
    lstAjax.AddItem "Dojo"
    lstAjax.AddItem "JQuery"
    lstAjax.AddItem "Mochikit"
    lstAjax.AddItem "Prototype"
    lstAjax.AddItem "Rico"
    lstAjax.AddItem "Scriptaculous"
    lstAjax.AddItem "Yahoo"
    
    arr_paths(1) = "reference\GuideJS13\ClientGuideJS13.zip"
    arr_paths(2) = "reference\GuideJS14\CoreGuideJS14.zip"
    arr_paths(3) = "reference\GuideJS15\CoreGuideJS15.zip"
    arr_paths(4) = "reference\referenceJS13\ClientReferenceJS13.zip"
    arr_paths(5) = "reference\referenceJS14\CoreReferenceJS14.zip"
    arr_paths(6) = "reference\referenceJS15\CoreReferenceJS15.zip"
    arr_paths(7) = "reference\domref\domref.zip"
    
    arr_paths_ajax(1) = "libraries\aflax\aflax.zip"
    arr_paths_ajax(2) = "libraries\dojo\dojo.zip"
    arr_paths_ajax(3) = "libraries\jquery\jquery.zip"
    arr_paths_ajax(4) = "libraries\mochikit\mochikit.zip"
    arr_paths_ajax(5) = "libraries\prototype\prototype.zip"
    arr_paths_ajax(6) = "libraries\rico\rico.zip"
    arr_paths_ajax(7) = "libraries\scriptaculous\scriptaculous.zip"
    arr_paths_ajax(8) = "libraries\yahoo\yahoo.zip"
    
    Debug.Print "load"
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set m_cUnzip = Nothing
    Set frmInstallHelp = Nothing
End Sub


Private Sub m_cUnzip_Progress(ByVal lCount As Long, ByVal sMsg As String)
    lblFile.Caption = sMsg
    DoEvents
End Sub


