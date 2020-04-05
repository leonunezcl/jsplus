VERSION 5.00
Begin VB.Form frmAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Settings"
   ClientHeight    =   2760
   ClientLeft      =   4140
   ClientTop       =   2280
   ClientWidth     =   5190
   Icon            =   "frmAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5190
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   840
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox chkPassive 
      Caption         =   "Passive"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2715
      TabIndex        =   7
      Top             =   1935
      Value           =   1  'Checked
      Width           =   945
   End
   Begin VB.TextBox txtFolder 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1410
      MaxLength       =   60
      TabIndex        =   5
      Top             =   1545
      Width           =   2445
   End
   Begin VB.TextBox txtPort 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4260
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "21"
      Top             =   435
      Width           =   825
   End
   Begin VB.TextBox txtConnect 
      Height          =   285
      Left            =   1410
      MaxLength       =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3690
   End
   Begin VB.TextBox txtURL 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1410
      MaxLength       =   60
      TabIndex        =   1
      Text            =   "ftp.microsoft.com"
      Top             =   435
      Width           =   2445
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1410
      MaxLength       =   60
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1170
      Width           =   2445
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1410
      MaxLength       =   60
      TabIndex        =   3
      Text            =   "anonymous"
      Top             =   795
      Width           =   2430
   End
   Begin VB.CheckBox chkAnonym 
      Caption         =   "Anonymous"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1410
      TabIndex        =   6
      Top             =   1935
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remote Folder"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   150
      TabIndex        =   13
      Top             =   1545
      Width           =   1035
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   3900
      TabIndex        =   12
      Top             =   465
      Width           =   330
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   11
      Top             =   120
      Width           =   435
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   150
      TabIndex        =   10
      Top             =   1185
      Width           =   735
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   150
      TabIndex        =   9
      Top             =   825
      Width           =   765
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Host:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   8
      Top             =   465
      Width           =   390
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public label As String
Public host As String
Public UserName As String
Public Port As String
Public pwd As String
Public lastdir As String
Public Anonymous As Integer
Public Passive As Integer
Private Sub chkAnonym_Click()
  If chkAnonym.Value = 1 Then
    txtUserName.Text = "anonymous"
    txtPassword.Text = "anonymous"
  End If
  
End Sub

Private Sub cmdSave_Click()
  
    On Error GoTo ErrorSave
    
    Dim FTPInfo As Ftp, fFile As Long
    Dim Path As String
    
    'validar la informacion
    If txtConnect.Text = "" Then
        txtConnect.SetFocus
        Exit Sub
    End If
    
    If txtUrl.Text = "" Then
        txtUrl.SetFocus
        Exit Sub
    End If
    
    If txtPort.Text = "" Or txtPort.Text = "0" Then
        txtPort.SetFocus
        Exit Sub
    End If
    
    If txtUserName.Text = "" Then
        txtUserName.SetFocus
        Exit Sub
    End If
    
    If txtPassword.Text = "" Then
        txtPassword.SetFocus
        Exit Sub
    End If
    
    util.Hourglass hwnd, True
    
    'generar el archivo nuevo
    
    FTPInfo.Anonymous = chkAnonym.Value
    FTPInfo.Passive = chkPassive.Value
    FTPInfo.Name = txtConnect.Text
    FTPInfo.UserName = txtUserName.Text
    FTPInfo.Password = Base64Encode(txtPassword.Text)
    FTPInfo.url = txtUrl.Text
    FTPInfo.PortNum = txtPort.Text
    
    If txtFolder.Text = "" Then
        FTPInfo.lastdir = "/"
    Else
        FTPInfo.lastdir = txtFolder.Text
    End If
    
    fFile = FreeFile()
    
    Path = App.Path & "\accounts"
    util.CrearDirectorio Path
    
    Open util.StripPath(Path) & FTPInfo.Name & ".ftp" For Binary Access Write As #fFile
        Put #fFile, , FTPInfo
    Close #fFile
    
    util.Hourglass hwnd, False
    
    frmFtpFiles.ObtenerCuentasFTP
    
    DoEvents
    
    Unload Me
    
    Exit Sub
    
ErrorSave:
    MsgBox "cmdSave_Click : " & Err & " " & Error$, vbCritical
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
  
    util.CenterForm Me
    util.Hourglass hwnd, True
        
    txtConnect.Text = label
    If Len(label) > 0 Then txtConnect.Locked = True
    txtUrl.Text = host
    txtPort.Text = Port
    txtUserName.Text = UserName
    txtPassword.Text = pwd
    txtFolder.Text = lastdir
    chkAnonym.Value = Anonymous
    chkPassive.Value = Passive
    
    util.SetNumber txtPort.hwnd
    util.Hourglass hwnd, False
    
    Debug.Print "load " & Me.Name
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload : " & Me.Name
    Set frmAccount = Nothing
End Sub

Private Sub txtConnect_GotFocus()
  txtConnect.BackColor = 14073525
  SelAll txtConnect
End Sub

Private Sub txtFolder_GotFocus()
    txtFolder.BackColor = 14073525
    SelAll txtPassword
End Sub


Private Sub txtFolder_LostFocus()
    txtFolder.BackColor = vbWindowBackground
End Sub


Private Sub txtPort_GotFocus()
  txtPort.BackColor = 14073525
  SelAll txtPort
End Sub

Private Sub txtPort_LostFocus()
  txtPort.BackColor = vbWindowBackground
End Sub
Private Sub txtConnect_LostFocus()
  txtConnect.BackColor = vbWindowBackground
End Sub
Private Sub txtURL_GotFocus()
  txtUrl.BackColor = 14073525
  SelAll txtUrl
End Sub

Private Sub txtURL_LostFocus()
  txtUrl.BackColor = vbWindowBackground
End Sub

Private Sub txtUserName_GotFocus()
  txtUserName.BackColor = 14073525
  SelAll txtUserName
End Sub

Private Sub txtUserName_LostFocus()
  txtUserName.BackColor = vbWindowBackground
End Sub

Private Sub txtPassword_GotFocus()
  txtPassword.BackColor = 14073525
  SelAll txtPassword
End Sub

Private Sub txtPassword_LostFocus()
  txtPassword.BackColor = vbWindowBackground
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub SelAll(txt As TextBox)
  txt.SelStart = 0
  txt.SelLength = Len(txt.Text)
End Sub
