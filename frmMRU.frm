VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpenMRU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MRU"
   ClientHeight    =   4560
   ClientLeft      =   2385
   ClientTop       =   4035
   ClientWidth     =   10455
   Icon            =   "frmMRU.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "R&efresh"
      Height          =   375
      Index           =   4
      Left            =   8880
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Remove Selected"
      Height          =   375
      Index           =   3
      Left            =   8880
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Clear List"
      Height          =   375
      Index           =   2
      Left            =   8880
      TabIndex        =   8
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   8880
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Open File"
      Height          =   375
      Index           =   0
      Left            =   8880
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   135
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   4620
      Width           =   8595
   End
   Begin VB.CheckBox chkSel 
      Caption         =   "Select All"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4215
      Width           =   1125
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   4590
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6900
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Frame Frame1 
      Caption         =   "Most Recent Used Files"
      Height          =   4080
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   8655
      Begin MSComctlLib.ListView lvwFiles 
         Height          =   3660
         Left            =   120
         TabIndex        =   1
         Top             =   285
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   6456
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Path"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Last Date Used"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Label lbltot 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8100
      TabIndex        =   5
      Top             =   4260
      Width           =   645
   End
End
Attribute VB_Name = "frmOpenMRU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CargaMRU()

    Dim arrFiles() As String
    Dim Archivo As String
    Dim total As Long
    Dim k As Integer
    Dim nFreeFile As Integer
    Dim linea As String
    
    Hourglass hwnd, True
    
    List1.Clear
    
    lvwFiles.ListItems.Clear
    lvwFiles.Sorted = False
    
    ReDim arrFiles(0)
    
    Archivo = util.StripPath(App.Path) & "mru.ini"
    
    If Not ArchivoExiste2(Archivo) Then
        Hourglass hwnd, False
        Exit Sub
    End If
    
    nFreeFile = FreeFile
    
    Open Archivo For Input As #nFreeFile
        Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, linea
            
            If Len(linea) > 0 Then
                k = k + 1
                ReDim Preserve arrFiles(k)
                arrFiles(k) = linea
            End If
        Loop
    Close #nFreeFile
    
    total = UBound(arrFiles)

    For k = 1 To total
        
        Archivo = arrFiles(k)
        
        lvwFiles.ListItems.Add , "k" & k, Explode(util.VBArchivoSinPath(Archivo), 1, "|")
        lvwFiles.ListItems("k" & k).SubItems(1) = util.PathArchivo(Explode(Archivo, 2, "|"))
        lvwFiles.ListItems("k" & k).SubItems(2) = Explode(Archivo, 3, "|")

    Next k
    
    lbltot.Caption = lvwFiles.ListItems.count & " Files"
    
    List1.Clear
    
    Hourglass hwnd, False
   
End Sub
Private Sub ClearList()

    util.Hourglass hwnd, True
    
    On Error Resume Next
    
    util.BorrarArchivo util.StripPath(App.Path) & "mru.ini"
    
    Err = 0
    
    lvwFiles.ListItems.Clear
    
    util.Hourglass hwnd, False
    
End Sub

Private Sub OpenFiles()

    Dim Archivo As String
    
    If lvwFiles.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    util.Hourglass hwnd, True
    
    cmd(0).Enabled = False
    cmd(1).Enabled = False
    cmd(2).Enabled = False
    cmd(3).Enabled = False
    
    Me.Hide
    
    Archivo = util.StripPath(lvwFiles.SelectedItem.SubItems(1)) & lvwFiles.SelectedItem.Text
            
    If ArchivoExiste2(Archivo) Then
       frmMain.opeEdit Archivo
    Else
       MsgBox "File : " & Archivo & " doesn't exists", vbCritical
    End If
    
    cmd(0).Enabled = True
    cmd(1).Enabled = True
    cmd(2).Enabled = True
    cmd(3).Enabled = True
    
    util.Hourglass hwnd, False
    
    Unload Me
    
End Sub

Private Sub RemoveSelected()

    Dim Archivo As String
    Dim k As Integer
    Dim j As Integer
    Dim nFreeFile As Long
    
    If lvwFiles.ListItems.count = 0 Then
        Exit Sub
    End If

    util.Hourglass hwnd, True
    
    Archivo = util.StripPath(App.Path) & "mru.ini"
    
    util.BorrarArchivo Archivo
    
    nFreeFile = FreeFile
    
    Open Archivo For Output As #nFreeFile
        j = 1
        For k = 1 To lvwFiles.ListItems.count
            If Not lvwFiles.ListItems(k).Selected Then
                Archivo = util.StripPath(lvwFiles.ListItems(k).SubItems(1)) & lvwFiles.ListItems(k).Text & "|" & lvwFiles.ListItems(k).SubItems(2)
                Print #nFreeFile, "File" & j & "|" & Archivo
                j = j + 1
            End If
        Next k
    Close #nFreeFile
    
    Call CargaMRU
    
    util.Hourglass hwnd, False
    
End Sub

Private Sub chkSel_Click()

    util.Hourglass hwnd, True
    
    Dim k As Integer
    
    For k = 1 To lvwFiles.ListItems.count
        If chkSel.Value Then
            lvwFiles.ListItems(k).Checked = True
        Else
            lvwFiles.ListItems(k).Checked = False
        End If
    Next k
    
    util.Hourglass hwnd, False
    
End Sub

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case 0  'open
            Call OpenFiles
        Case 1  'cancel
            Unload Me
        Case 2  'clear list
            Call ClearList
        Case 3  'remove selected
            Call RemoveSelected
        Case 4
            Call CargaMRU
    End Select
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
    
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    cmd(0).Refresh
    cmd(1).Refresh
    cmd(2).Refresh
    cmd(3).Refresh
    cmd(4).Refresh
    
    Call CargaMRU
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmOpenMRU = Nothing
End Sub


Private Sub lvwFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

   If lvwFiles.SortOrder = lvwAscending Then
      lvwFiles.SortOrder = lvwDescending
   Else
      lvwFiles.SortOrder = lvwAscending
   End If
   lvwFiles.Sorted = True
End Sub


