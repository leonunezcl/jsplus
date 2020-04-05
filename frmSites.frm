VERSION 5.00
Object = "{FCFAF346-DE8A-4FB6-8612-5000548EFDC7}#2.0#0"; "vbsListView6.ocx"
Begin VB.Form frmSites 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accounts"
   ClientHeight    =   3030
   ClientLeft      =   4365
   ClientTop       =   3525
   ClientWidth     =   4245
   ControlBox      =   0   'False
   Icon            =   "frmSites.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   765
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   2250
   End
   Begin vbalListViewLib6.vbalListViewCtl lvwAco 
      Height          =   2475
      Left            =   30
      TabIndex        =   0
      Top             =   240
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   4366
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      View            =   1
      MultiSelect     =   -1  'True
      LabelEdit       =   0   'False
      AutoArrange     =   0   'False
      Appearance      =   0
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin jsplus.MyButton cmdConnect 
      Height          =   405
      Left            =   2745
      TabIndex        =   4
      Top             =   195
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Connect"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePos      =   3
   End
   Begin jsplus.MyButton cmdAddAco 
      Height          =   405
      Left            =   2745
      TabIndex        =   5
      Top             =   660
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Add Account"
      AccessKey       =   "A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePos      =   3
   End
   Begin jsplus.MyButton cmdDelAco 
      Height          =   405
      Left            =   2745
      TabIndex        =   6
      Top             =   1140
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Remove Account"
      AccessKey       =   "R"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePos      =   3
   End
   Begin jsplus.MyButton cmdEdtAco 
      Height          =   405
      Left            =   2745
      TabIndex        =   7
      Top             =   1605
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Edit Account"
      AccessKey       =   "E"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePos      =   3
   End
   Begin jsplus.MyButton cmdExit 
      Height          =   405
      Left            =   2745
      TabIndex        =   8
      Top             =   2100
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "E&xit"
      AccessKey       =   "x"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePos      =   3
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Available Accounts"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   1365
   End
   Begin VB.Label lblmsg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ready"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   30
      TabIndex        =   2
      Top             =   2775
      Width           =   465
   End
End
Attribute VB_Name = "frmSites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fload As Boolean
Public tipo_conexion As Integer
Public IdDoc As Integer
Public Sub GetAccounts()
  
    Dim s As String
    Dim k As Integer
    
    fload = True
    
    lvwAco.ListItems.Clear
    
    s = Dir(App.Path & "\accounts\")
    
    k = 1
    Do While s <> ""
        If VBA.Right(s, 3) = "ftp" Then
            lvwAco.ListItems.Add , "k" & k, VBA.Left$(s, Len(s) - 4)
            k = k + 1
        End If
        s = Dir
    Loop
    
    fload = False
    
End Sub


Private Sub cmdAddAco_Click()
    frmAccount.Show vbModal
End Sub

Private Sub cmdConnect_Click()

    'Dim fFile As Long
    Dim file As String
    
    If Not lvwAco.SelectedItem Is Nothing Then
        
        file = util.StripPath(App.Path) & "accounts\" & lvwAco.SelectedItem.Text & ".ftp"
        
        lblmsg.Caption = "Connecting ..."
        
        'If Not FTPManager.open_site(File) Then
        '    MsgBox "Error reading the FTP account : " & File, vbCritical
        '    Exit Sub
        'End If
        
        If Not FTPManager.OpenSiteInformation(file) Then
            Exit Sub
        End If
        
        cmdAddAco.Enabled = False
        cmdDelAco.Enabled = False
        cmdEdtAco.Enabled = False
        cmdExit.Enabled = False
        
        'download
        lblmsg.Caption = "Please wait. Loading files from server ...."

        frmFtpFiles.tipo_conexion = tipo_conexion
        
        
        If FTPManager.lastdir = "" Then
            'frmFtpFiles.lastdir = "/"
        Else
            'frmFtpFiles.lastdir = FTPManager.lastdir
        End If
        
        cmdAddAco.Enabled = True
        cmdDelAco.Enabled = True
        cmdEdtAco.Enabled = True
        cmdExit.Enabled = True
        
        'Me.Hide
        frmFtpFiles.Show vbModal
        'Me.Show vbModal
        lblmsg.Caption = "Ready"
    Else
        MsgBox "Select or creates account first.", vbCritical
    End If
    
End Sub

Private Sub cmdDelAco_Click()

    Dim file As String
    
    If Not lvwAco.SelectedItem Is Nothing Then
        
        file = util.StripPath(App.Path) & "accounts\" & lvwAco.SelectedItem.Text & ".ftp"
        
        If Not ArchivoExiste2(file) Then
            MsgBox "File not found : " & file, vbCritical
            Exit Sub
        End If
        
        If Confirma("Are you sure to remove this account") = vbYes Then
            util.BorrarArchivo file
            Call GetAccounts
        End If
    End If
    
End Sub

Private Sub cmdEdtAco_Click()

    Dim fFile As Long, FTPInfo As Ftp
    Dim file As String
    
    If fload Then Exit Sub
    
    If Not lvwAco.SelectedItem Is Nothing Then
        file = lvwAco.SelectedItem.Text
    Else
        MsgBox "Please select an account to access first.", vbCritical
        Exit Sub
    End If
  
    file = App.Path & "\accounts\" & lvwAco.SelectedItem.Text & ".ftp"
  
    If Not ArchivoExiste2(file) Then
        MsgBox "File doesn't exist : " & file, vbCritical
        Exit Sub
    End If
    
    fFile = FreeFile()
    
    Open file For Binary Access Read As #fFile
        Get #fFile, , FTPInfo
    Close #fFile
    
    frmAccount.label = FTPInfo.Name
    frmAccount.host = FTPInfo.url
    frmAccount.Port = FTPInfo.PortNum
    frmAccount.UserName = FTPInfo.UserName
    frmAccount.pwd = Base64Decode(FTPInfo.Password)
    frmAccount.Anonymous = FTPInfo.Anonymous
    frmAccount.lastdir = FTPInfo.lastdir
    frmAccount.Passive = FTPInfo.Passive
    frmAccount.Show vbModal
    
End Sub

Private Sub cmdExit_Click()
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
    
    'set_color_form Me
    'SetLayered hwnd, True
    
    With lvwAco
        .Columns.Add , "c1", "Connections", , 2450
    End With
    
    Call GetAccounts
    
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    
    cmdConnect.Refresh
    cmdDelAco.Refresh
    cmdAddAco.Refresh
    cmdEdtAco.Refresh
    cmdExit.Refresh
    
    Debug.Print "load : " & Me.Name
    
    'DrawXPCtl Me
    
    util.Hourglass hwnd, False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name = "frmEdit" Then
            frm.Refresh
        End If
    Next
    
    Call clear_memory(Me)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload : " & Me.Name
    Set frmSites = Nothing
End Sub


Private Sub lvwAco_ItemDblClick(ITem As vbalListViewLib6.cListItem)
    cmdConnect_Click
End Sub

