VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conectar"
   ClientHeight    =   2895
   ClientLeft      =   3105
   ClientTop       =   3465
   ClientWidth     =   5175
   ControlBox      =   0   'False
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   Begin VB.CheckBox chkPassive 
      Caption         =   "Pasivo"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Nueva sesión"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtLocalDir 
      Height          =   285
      Left            =   2280
      TabIndex        =   12
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox txtRemoteDir 
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   480
      Width           =   2535
   End
   Begin VB.ComboBox cmbProfiles 
      Height          =   315
      ItemData        =   "frmConnect.frx":27A2
      Left            =   2280
      List            =   "frmConnect.frx":27A4
      TabIndex        =   7
      Text            =   "cmbProfiles"
      Top             =   90
      Width           =   2175
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Conectar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Directorio local inicial:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Directorio inicial remoto:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Host Name/Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sesión:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'WinFTP, created by the KPD-Team 2000
'This file can be downloaded from http://www.allapi.net/
'For questions or comments, contact us at KPDTeam@Allapi.net

' You are free to use this code within your own applications,
' but you are expressly forbidden from selling or otherwise
' distributing this source code without prior written consent.
' This includes both posting free demo projects made from this
' code as well as reproducing the code in text or html format.

Private Type tProfile
    ProfileName As String
    Host As String
    Username As String
    Password As String
    RemoteDir As String
    LocalDir As String
End Type
Dim Profiles() As tProfile, nProfiles As Long, nActive As Long
Private Sub cmbProfiles_Change()
    Dim cnt As Long, OldStart As Long
    Profiles(nActive).ProfileName = cmbProfiles.Text
    OldStart = cmbProfiles.SelStart
    cmbProfiles.Clear
    For cnt = 1 To nProfiles
        cmbProfiles.AddItem Profiles(cnt).ProfileName
    Next cnt
    cmbProfiles.ListIndex = nActive - 1
    cmbProfiles.SelStart = OldStart
    cmbProfiles.SelLength = 0
End Sub
Private Sub cmbProfiles_Click()
    nActive = cmbProfiles.ListIndex + 1
    ZetProfile
End Sub
Private Sub cmdConnect_Click()
    SaveSetting "KPD FTP", "Last Connection", "lc", CStr(cmbProfiles.ListIndex)
    frmMainFTP.cmdConnect.Caption = "Desconectar"
    frmMainFTP.txtStatus.Text = ""
    If Dir(txtLocalDir.Text) <> "" And txtLocalDir.Text <> "" Then
        If (GetAttr(txtLocalDir.Text) And vbDirectory) = vbDirectory Then
            frmMainFTP.FillLocalListView txtLocalDir.Text
        End If
    End If
    frmMainFTP.rfConnection.CreateConnection True, txtHost.Text, txtUser.Text, txtPassword.Text
    GetStatus
    If txtRemoteDir.Text <> "" Then
        frmMainFTP.rfConnection.SetNewDirectory txtRemoteDir.Text
        GetStatus
    End If
    frmMainFTP.FillRemoteListView
    Unload Me
    frmProgress.Visible = True
End Sub
Private Sub cmdNew_Click()
    nProfiles = nProfiles + 1
    ReDim Profiles(1 To nProfiles) As tProfile
    Profiles(nProfiles).ProfileName = "Nueva sesion"
    cmbProfiles.AddItem Profiles(nProfiles).ProfileName
    cmbProfiles.ListIndex = nProfiles - 1
End Sub
Private Sub Form_Load()
    ReadProfiles
    If nProfiles = 0 Then
        cmdNew_Click
    Else
        If cmbProfiles.ListCount > Val(GetSetting("KPD FTP", "Last Connection", "lc", 0)) Then
            cmbProfiles.ListIndex = GetSetting("KPD FTP", "Last Connection", "lc", 0)
        End If
    End If
End Sub
'Very easy encryption
Public Function Encrypt(UnEncrypted As String, ByVal iEncrypt As Integer) As String
    Dim cnt As Long, NewChr As Long
    For cnt = 1 To Len(UnEncrypted)
        NewChr = Asc(Mid$(UnEncrypted, cnt, 1)) + iEncrypt * 30
        If NewChr > 255 Then
            Do: NewChr = NewChr - 256: Loop Until NewChr < 255
        ElseIf NewChr < 0 Then
            Do: NewChr = NewChr + 256: Loop Until NewChr > 0
        End If
        Mid$(UnEncrypted, cnt, 1) = Chr$(NewChr)
    Next cnt
    Encrypt = UnEncrypted
End Function
Sub ReadProfiles()
    nProfiles = 0
    cmbProfiles.Clear
    Do
        If GetSetting("KPD FTP", "Profiles", "Profile" + CStr(nProfiles + 1), False) = False Then
            Exit Do
        End If
        nProfiles = nProfiles + 1
        ReDim Preserve Profiles(1 To nProfiles) As tProfile
        Profiles(nProfiles).ProfileName = GetSetting("KPD FTP", "Profiles\Profile" + CStr(nProfiles), "ProfileName", "")
        cmbProfiles.AddItem Profiles(nProfiles).ProfileName
        Profiles(nProfiles).Host = GetSetting("KPD FTP", "Profiles\Profile" + CStr(nProfiles), "Host", "")
        Profiles(nProfiles).Username = GetSetting("KPD FTP", "Profiles\Profile" + CStr(nProfiles), "User", "")
        Profiles(nProfiles).Password = Encrypt(GetSetting("KPD FTP", "Profiles\Profile" + CStr(nProfiles), "Password", ""), -1)
        Profiles(nProfiles).RemoteDir = GetSetting("KPD FTP", "Profiles\Profile" + CStr(nProfiles), "RemoteDir", "")
        Profiles(nProfiles).LocalDir = GetSetting("KPD FTP", "Profiles\Profile" + CStr(nProfiles), "LocalDir", "")
    Loop
End Sub
Sub ZetProfile()
    txtHost.Text = Profiles(nActive).Host
    txtUser.Text = Profiles(nActive).Username
    txtPassword.Text = Profiles(nActive).Password
    txtRemoteDir.Text = Profiles(nActive).RemoteDir
    txtLocalDir.Text = Profiles(nActive).LocalDir
End Sub
Sub SaveProfiles()
    Dim cnt As Long
    For cnt = 1 To nProfiles
        SaveSetting "KPD FTP", "Profiles", "Profile" + CStr(cnt), True
        SaveSetting "KPD FTP", "Profiles\Profile" + CStr(cnt), "ProfileName", Profiles(cnt).ProfileName
        SaveSetting "KPD FTP", "Profiles\Profile" + CStr(cnt), "Host", Profiles(cnt).Host
        SaveSetting "KPD FTP", "Profiles\Profile" + CStr(cnt), "User", Profiles(cnt).Username
        SaveSetting "KPD FTP", "Profiles\Profile" + CStr(cnt), "Password", Encrypt(Profiles(cnt).Password, 1)
        SaveSetting "KPD FTP", "Profiles\Profile" + CStr(cnt), "RemoteDir", Profiles(cnt).RemoteDir
        SaveSetting "KPD FTP", "Profiles\Profile" + CStr(cnt), "LocalDir", Profiles(cnt).LocalDir
    Next cnt
    SaveSetting "KPD FTP", "Profiles", "Profile" + CStr(nProfiles + 1), False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveProfiles
End Sub

Private Sub txtHost_Change()
    Profiles(nActive).Host = txtHost.Text
End Sub
Private Sub txtHost_GotFocus()
    txtHost.SelStart = 0
    txtHost.SelLength = Len(txtHost.Text)
End Sub
Private Sub txtLocalDir_Change()
    Profiles(nActive).LocalDir = txtLocalDir.Text
End Sub
Private Sub txtLocalDir_GotFocus()
    txtLocalDir.SelStart = 0
    txtLocalDir.SelLength = Len(txtLocalDir.Text)
End Sub
Private Sub txtPassword_Change()
    Profiles(nActive).Password = txtPassword.Text
End Sub
Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub
Private Sub txtRemoteDir_Change()
    Profiles(nActive).RemoteDir = txtRemoteDir.Text
End Sub
Private Sub txtRemoteDir_GotFocus()
    txtRemoteDir.SelStart = 0
    txtRemoteDir.SelLength = Len(txtRemoteDir.Text)
End Sub
Private Sub txtUser_Change()
    Profiles(nActive).Username = txtUser.Text
End Sub
Private Sub txtUser_GotFocus()
    txtUser.SelStart = 0
    txtUser.SelLength = Len(txtUser.Text)
End Sub
