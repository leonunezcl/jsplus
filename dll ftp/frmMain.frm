VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMainFTP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sesión FTP"
   ClientHeight    =   6750
   ClientLeft      =   2040
   ClientTop       =   1680
   ClientWidth     =   10230
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   682
   Begin VB.CommandButton cmdProgress 
      Caption         =   "Mostrar Avance"
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      ToolTipText     =   "Mostrar progreso de envío"
      Top             =   6240
      Width           =   1575
   End
   Begin VB.ListBox lstTemp 
      Height          =   255
      Left            =   4080
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   3480
      Top             =   6165
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27A2
            Key             =   "Directory"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":289C
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2996
            Key             =   "File"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A90
            Key             =   "Drive"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8400
      TabIndex        =   9
      ToolTipText     =   "Salir de sesión FTP"
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1740
      TabIndex        =   7
      ToolTipText     =   "Cancelar"
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Conectar"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Conectar"
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   9975
      Begin VB.TextBox txtStatus 
         Height          =   700
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   180
         Width           =   9795
      End
   End
   Begin VB.OptionButton chkBinary 
      Caption         =   "Binario"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   4920
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton chkASCII 
      Caption         =   "ASCII"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   4920
      Width           =   855
   End
   Begin VB.Frame fraRemote 
      Caption         =   "Sistema Remoto"
      Height          =   4815
      Left            =   4920
      TabIndex        =   1
      Top             =   0
      Width           =   5295
      Begin VB.TextBox txtRemPath 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   5055
      End
      Begin VB.CommandButton cmdrRefresh 
         Caption         =   "UPD"
         Height          =   375
         Left            =   4320
         TabIndex        =   26
         ToolTipText     =   "Actualizar contenido"
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmdrDelete 
         Caption         =   "DEL"
         Height          =   375
         Left            =   4320
         TabIndex        =   25
         ToolTipText     =   "Borrar"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton cmdrRename 
         Caption         =   "REN"
         Height          =   375
         Left            =   4320
         TabIndex        =   24
         ToolTipText     =   "Cambiar nombre"
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdrExec 
         Cancel          =   -1  'True
         Caption         =   "VER"
         Height          =   375
         Left            =   4320
         TabIndex        =   23
         ToolTipText     =   "Ver contenido"
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdrMkDir 
         Caption         =   "MKD"
         Height          =   375
         Left            =   4320
         TabIndex        =   22
         ToolTipText     =   "Crear directorio"
         Top             =   600
         Width           =   855
      End
      Begin MSComctlLib.ListView lvRemote 
         Height          =   4095
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         Icons           =   "ilList"
         SmallIcons      =   "ilList"
         ColHdrIcons     =   "ilList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Name"
            Text            =   "Nombre"
            Object.Width           =   3987
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Size"
            Text            =   "Tamaño"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label lblNumFiles 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   4320
         TabIndex        =   30
         Top             =   4440
         Width           =   855
      End
   End
   Begin VB.Frame fraLocal 
      Caption         =   "Sistema Local"
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtPattern 
         Height          =   285
         Left            =   3720
         TabIndex        =   29
         Text            =   "*.*"
         ToolTipText     =   "Digita patrón de búsqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdPutNew 
         Caption         =   "-->"
         Height          =   615
         Left            =   3900
         Picture         =   "frmMain.frx":2B8A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3840
         Width           =   495
      End
      Begin VB.CommandButton cmdPut 
         Caption         =   "-->"
         Height          =   495
         Left            =   3900
         TabIndex        =   21
         ToolTipText     =   "Enviar"
         Top             =   3240
         Width           =   495
      End
      Begin VB.CommandButton cmdRetrieve 
         Caption         =   "<--"
         Height          =   495
         Left            =   3900
         TabIndex        =   20
         ToolTipText     =   "Recibir"
         Top             =   2640
         Width           =   495
      End
      Begin VB.CommandButton cmdlRefresh 
         Caption         =   "UPD"
         Height          =   375
         Left            =   3720
         TabIndex        =   19
         ToolTipText     =   "Actualizar contenido"
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmdlDelete 
         Caption         =   "DEL"
         Height          =   375
         Left            =   3720
         TabIndex        =   18
         ToolTipText     =   "Borrar archivo"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton cmdlRename 
         Caption         =   "REN"
         Height          =   375
         Left            =   3720
         TabIndex        =   17
         ToolTipText     =   "Renombrar archivo"
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdlExec 
         Caption         =   "EXE"
         Height          =   375
         Left            =   3720
         TabIndex        =   16
         ToolTipText     =   "Ejecutar aplicación"
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdlMkDir 
         Caption         =   "MKD"
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         ToolTipText     =   "Crear directorio"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtCurPath 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   3495
      End
      Begin VB.PictureBox picBack 
         BackColor       =   &H80000005&
         Height          =   4095
         Left            =   120
         ScaleHeight     =   269
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   229
         TabIndex        =   12
         Top             =   600
         Width           =   3495
         Begin MSComctlLib.ListView lvLocal 
            Height          =   3915
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   6906
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   393217
            Icons           =   "ilList"
            SmallIcons      =   "ilList"
            ColHdrIcons     =   "ilList"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "Name"
               Text            =   "Nombre"
               Object.Width           =   4022
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "Size"
               Text            =   "Tamaño"
               Object.Width           =   1587
            EndProperty
         End
      End
      Begin VB.Label lblLocFiles 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   3720
         TabIndex        =   31
         Top             =   4470
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMainFTP"
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
'
'  Changes:
'     03/14/01, TPA:  Removed file writing from cmdrExec_Click()
'                     because it has been added to cRemoteFile.GetFile()

Private sDrives As String

Const SW_SHOWNORMAL = 1
Const FO_DELETE = &H3
Const FO_RENAME = &H4
Const FOF_ALLOWUNDO = &H40
Const FOF_NOCONFIRMATION = &H10
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function Putfocus Lib "User32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public WithEvents rfFile As cRemoteFile
Attribute rfFile.VB_VarHelpID = -1
Public WithEvents rfConnection As cConnection
Attribute rfConnection.VB_VarHelpID = -1
Public cFiles As New Collection, cAttrs As New Collection, cSize As New Collection, cRemAttrs As New Collection
Public nTotal As Long, DriveCol As New Collection, sCurPath As String
Private Sub chkASCII_Click()
    UploadFlag = FTP_TRANSFER_TYPE_ASCII
End Sub
Private Sub chkBinary_Click()
    UploadFlag = FTP_TRANSFER_TYPE_BINARY
End Sub
Private Sub cmdAbout_Click()
    MsgBox "WinFTP - Created by the KPD-Team 2000" + vbCrLf + "Visit our site at http://www.allapi.net/" + vbCrLf + "Or E-Mail us at KPDTeam@allapi.net", vbInformation + vbOKOnly, App.Title
End Sub

Private Sub cmdConnect_Click()
    If cmdConnect.Caption = "Conectar" Then
        frmConnect.Show vbModal, Me
    Else
        rfConnection.Disconnect
        GetStatus
        'While lvRemote.ListItems.Count > 0
            'lvRemote.ListItems.Remove lvRemote.ListItems.Count
        'Wend
        'If bFOBusy Then Do: DoEvents: Loop Until bFOBusy = False
        lvRemote.ListItems.Clear
        txtRemPath.Text = ""
        cmdConnect.Caption = "Conectar"
    End If
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdlDelete_Click()
    Dim Msg As VbMsgBoxResult, cnt As Long
    Msg = MsgBox("Seguro de borrar?", vbQuestion + vbYesNo, App.Title)
    If Msg = vbYes Then
        For cnt = 2 To lvLocal.ListItems.Count - DriveCol.Count
            If lvLocal.ListItems.Item(cnt).Selected = True Then
                If lvLocal.ListItems.Item(cnt).Text = ".." Or VBA.Right$(lvLocal.ListItems.Item(cnt).Text, 2) <> ":\" Then
                    Dim FO As SHFILEOPSTRUCT
                    FO.pFrom = sCurPath + lvLocal.ListItems.Item(cnt).Text
                    FO.fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
                    FO.wFunc = FO_DELETE
                    SHFileOperation FO
                End If
            End If
        Next cnt
        FillLocalListView sCurPath
    End If
End Sub
Private Sub cmdlExec_Click()
    If lvLocal.SelectedItem <> ".." Then ShellExecute 0, vbNullString, sCurPath + lvLocal.SelectedItem, vbNullString, sCurPath, SW_SHOWNORMAL
End Sub
Private Sub cmdlMkDir_Click()
    Dim sRet As String
    sRet = InputBox("Ingresa el nombre del directorio:")
    If sRet <> "" Then
        MkDir sCurPath + sRet
        FillLocalListView sCurPath
    End If
End Sub
Private Sub cmdlRefresh_Click()
    FillLocalListView sCurPath
End Sub
Private Sub cmdlRename_Click()
    If lvLocal.SelectedItem = ".." Or VBA.Right$(lvLocal.SelectedItem, 2) = ":\" Then Exit Sub
    Dim sRet As String
    sRet = InputBox("Nuevo nombre para " + lvLocal.SelectedItem)
    If sRet <> "" Then
        Dim FO As SHFILEOPSTRUCT
        FO.pFrom = sCurPath + lvLocal.SelectedItem
        FO.pTo = sCurPath + sRet
        FO.fFlags = FOF_NOCONFIRMATION
        FO.wFunc = FO_RENAME
        SHFileOperation FO
        FillLocalListView sCurPath
        FillLocalListView sCurPath
    End If
End Sub
Private Sub cmdProgress_Click()
    frmProgress.Visible = True
End Sub
Private Sub cmdPut_Click()
    Dim cnt As Long, bOk As Boolean
    For cnt = 1 To lvLocal.ListItems.Count
        If lvLocal.ListItems.Item(cnt).Selected = True Then
            If lvLocal.ListItems.Item(cnt).Text <> ".." And VBA.Right$(lvLocal.ListItems.Item(cnt).Text, 2) <> ":\" Then
                If (GetAttr(sCurPath + lvLocal.ListItems.Item(cnt).Text) And vbDirectory) <> vbDirectory Then
                    AddToCollection FOP_UPLOAD, lvLocal.ListItems.Item(cnt).Text, sCurPath, FileLen(sCurPath + lvLocal.ListItems.Item(cnt).Text)
                    bOk = True
                End If
            End If
        End If
    Next cnt
    If bFOBusy = False And bOk Then meStartFO
End Sub
Function FindRemoteFileSize(ByVal sInput As String) As Long
    Dim cnt As Long
    sInput = LCase$(sInput)
    FindRemoteFileSize = -1
    For cnt = 1 To cFiles.Count
        If LCase$(cFiles.Item(cnt)) = sInput Then
            FindRemoteFileSize = cSize(cnt)
            Exit For
        End If
    Next cnt
End Function
Private Sub cmdPutNew_Click()
    Dim cnt As Long, bOk As Boolean, Ret As Long
    For cnt = 1 To lvLocal.ListItems.Count
        If lvLocal.ListItems.Item(cnt).Selected = True Then
            If lvLocal.ListItems.Item(cnt).Text <> ".." And VBA.Right$(lvLocal.ListItems.Item(cnt).Text, 2) <> ":\" Then
                If (GetAttr(sCurPath + lvLocal.ListItems.Item(cnt).Text) And vbDirectory) <> vbDirectory Then
                    Ret = FindRemoteFileSize(lvLocal.ListItems.Item(cnt).Text)
                    If Ret <> -1 Then
                        If FileLen(sCurPath + lvLocal.ListItems.Item(cnt).Text) <> Ret Then
                            AddToCollection FOP_UPLOAD, lvLocal.ListItems.Item(cnt).Text, sCurPath, FileLen(sCurPath + lvLocal.ListItems.Item(cnt).Text)
                            bOk = True
                        End If
                    Else
                        AddToCollection FOP_UPLOAD, lvLocal.ListItems.Item(cnt).Text, sCurPath, FileLen(sCurPath + lvLocal.ListItems.Item(cnt).Text)
                        bOk = True
                    End If
                End If
            End If
        End If
    Next cnt
    If bFOBusy = False And bOk Then meStartFO
End Sub
Private Sub cmdrDelete_Click()
    Dim Msg As VbMsgBoxResult, cnt As Long
    Msg = MsgBox("Seguro de borrar?", vbQuestion + vbYesNo, App.Title)
    If Msg = vbYes Then
        For cnt = 1 To cRemAttrs.Count
            If lvRemote.ListItems.Item(cnt).Selected = True And (cRemAttrs.Item(cnt) And vbDirectory) <> vbDirectory Then
                rfFile.RemoteFile = lvRemote.ListItems.Item(cnt).Text
                rfFile.DeleteFile rfConnection
                GetStatus
            ElseIf lvRemote.ListItems.Item(cnt).Selected = True And (cRemAttrs.Item(cnt) And vbDirectory) = vbDirectory And lvRemote.ListItems.Item(cnt).Text <> ".." Then
                rfConnection.RemoveDirectory lvRemote.ListItems.Item(cnt).Text
                GetStatus
            End If
        Next cnt
        FillRemoteListView
    End If
End Sub
Private Sub cmdRetrieve_Click()
    Dim cnt As Long, bOk As Boolean
    For cnt = 1 To lvRemote.ListItems.Count
        If lvRemote.ListItems.Item(cnt).Selected = True Then
            If lvRemote.ListItems.Item(cnt).Text <> ".." Then
                If (cRemAttrs.Item(cnt) And vbDirectory) <> vbDirectory Then
                    AddToCollection FOP_DOWNLOAD, lvRemote.ListItems.Item(cnt).Text, sCurPath, Val(lvRemote.ListItems.Item(cnt).SubItems(1))
                    bOk = True
                End If
            End If
        End If
    Next cnt
    If bFOBusy = False And bOk Then meStartFO
End Sub
Private Sub cmdrExec_Click()
    If (cRemAttrs.Item(lvRemote.SelectedItem.Index) And vbDirectory) <> vbDirectory Then
        Dim strTemp As String
        strTemp = String(100, 0)
        GetTempPath 100, strTemp
        strTemp = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)
        rfFile.RemoteFile = lvRemote.SelectedItem
        rfFile.GetFile rfConnection, strTemp + lvRemote.SelectedItem
        'Open strTemp + lvRemote.SelectedItem For Binary As #1
        '    Put #1, , rfFile.FileData
        'Close
        ShellExecute 0, vbNullString, strTemp + lvRemote.SelectedItem, vbNullString, vbNullString, SW_SHOWNORMAL
    End If
End Sub
Private Sub cmdrMkDir_Click()
    Dim Ret As String
    Ret = InputBox("Ingresa el nombre del directorio:")
    If Ret <> "" Then
        rfConnection.CreateDirectory Ret
        GetStatus
        FillRemoteListView
    End If
End Sub
Private Sub cmdrRefresh_Click()
    FillRemoteListView
End Sub
Private Sub cmdrRename_Click()
    If lvRemote.SelectedItem.Index = 1 Then Exit Sub
    Dim Ret As String
    Ret = InputBox("Nuevo nombre para " + lvRemote.SelectedItem.Text)
    If Ret <> "" Then
        rfFile.RemoteFile = lvRemote.SelectedItem.Text
        rfFile.RenameFile rfConnection, Ret
        GetStatus
        FillRemoteListView
    End If
End Sub
Private Sub Form_Load()
    lvLocal.Move 0, 0, picBack.ScaleWidth, picBack.ScaleHeight
    GetDrives
    SetEnabled False
    txtPattern.Text = GetSetting("KPD FTP", "Pattern", "Pattern", "*.*")
    FillLocalListView GetSetting("KPD FTP", "Path", "LastPath", "C:\Windows\")
    Set rfFile = New cRemoteFile
    Set rfConnection = New cConnection
    UploadFlag = FTP_TRANSFER_TYPE_BINARY
    'frmProgress.Show
    Load frmProgress
    frmProgress.Visible = False
    Me.Show
    frmConnect.Show vbModal, Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If rfConnection.Connected Then rfConnection.Disconnect
    SaveSetting "KPD FTP", "Pattern", "Pattern", txtPattern.Text
    SaveSetting "KPD FTP", "Path", "LastPath", sCurPath
    Unload frmProgress
    Set frmMainFTP = Nothing
End Sub

Private Sub lvLocal_DblClick()
    If VBA.Right$(lvLocal.SelectedItem, 2) = ":\" Then
        FillLocalListView lvLocal.SelectedItem
    ElseIf lvLocal.SelectedItem = ".." Then
        FillLocalListView RemoveLastDir(sCurPath)
    ElseIf (GetAttr(sCurPath + lvLocal.SelectedItem) And vbDirectory) = vbDirectory Then
        FillLocalListView sCurPath + lvLocal.SelectedItem + "\"
    Else
        Dim cnt As Long, bOk As Boolean
        For cnt = 1 To lvLocal.ListItems.Count
            If lvLocal.ListItems.Item(cnt).Selected = True Then
                If lvLocal.ListItems.Item(cnt).Text <> ".." And VBA.Right$(lvLocal.ListItems.Item(cnt).Text, 2) <> ":\" Then
                    If (GetAttr(sCurPath + lvLocal.ListItems.Item(cnt).Text) And vbDirectory) <> vbDirectory Then
                        AddToCollection FOP_UPLOAD, lvLocal.ListItems.Item(cnt).Text, sCurPath, FileLen(sCurPath + lvLocal.ListItems.Item(cnt).Text)
                        bOk = True
                    End If
                End If
            End If
        Next cnt
        If bFOBusy = False And bOk Then meStartFO
    End If
End Sub
Private Sub lvRemote_DblClick()
    If bFOBusy Then
        MsgBox "Unable to execute command...", vbExclamation + vbOKOnly, App.Title
        Exit Sub
    End If
    Dim Ret As Long
    Ret = GetRemoteIndex
    If Ret <> -1 Then
        If (cRemAttrs.Item(Ret) And vbDirectory) = vbDirectory Then
            rfConnection.SetNewDirectory lvRemote.SelectedItem
            GetStatus
            FillRemoteListView
        End If
    End If
End Sub
Public Function GetRemoteIndex() As Long
    Dim bOk As Boolean
    For GetRemoteIndex = 1 To lvRemote.ListItems.Count
        If lvRemote.ListItems.Item(GetRemoteIndex).Selected Then
            bOk = True
            Exit For
        End If
    Next GetRemoteIndex
    If bOk = False Then GetRemoteIndex = -1
End Function
Sub FillLocalListView(sPath As String)
    Dim Ret As String, cnt As Long, Tel As Long
    If IsDriveAvailable(sPath) = False Then
        MsgBox "Unidad no está lista!", vbCritical + vbOKOnly, App.Title
        Exit Sub
    End If
    lvLocal.Visible = False
    lvLocal.ListItems.Clear
    lvLocal.ListItems.Add , , "..", , "Up"
    lstTemp.Clear
    Ret = Dir(sPath, vbDirectory)
    While Ret <> ""
        If (GetAttr(sPath + Ret) And vbDirectory) = vbDirectory And Ret <> ".." And Ret <> "." Then lstTemp.AddItem Ret
        Ret = Dir()
    Wend
    Tel = lstTemp.ListCount
    For cnt = 0 To lstTemp.ListCount - 1
        lvLocal.ListItems.Add , , lstTemp.List(cnt), , "Directory"
    Next cnt
    lstTemp.Clear
    Ret = Dir(sPath + txtPattern.Text, vbNormal)
    While Ret <> ""
        If (GetAttr(sPath + Ret) And vbDirectory) <> vbDirectory Then lstTemp.AddItem Ret
        Ret = Dir()
    Wend
    Tel = Tel + lstTemp.ListCount
    For cnt = 0 To lstTemp.ListCount - 1
        lvLocal.ListItems.Add , , lstTemp.List(cnt), , "File"
        lvLocal.ListItems.Item(lvLocal.ListItems.Count).SubItems(1) = FileLen(sPath + lstTemp.List(cnt))
    Next cnt
    For cnt = 1 To DriveCol.Count
        lvLocal.ListItems.Add , , DriveCol.Item(cnt), , "Drive"
    Next cnt
    lvLocal.Visible = True
    txtCurPath.Text = sPath
    lblLocFiles.Caption = CStr(Tel)
    sCurPath = sPath
    Putfocus lvLocal.hWnd
End Sub
Sub GetDrives()
    Dim LDs As Long, cnt As Long
    LDs = GetLogicalDrives
    sDrives = "Available drives:"
    For cnt = 0 To 25
        If (LDs And 2 ^ cnt) <> 0 Then
            DriveCol.Add Chr$(65 + cnt) + ":\"
        End If
    Next cnt
End Sub
Public Function IsDriveAvailable(sDrive As String) As Boolean
    If GetVolumeInformation(Left$(sDrive, 3), vbNullString, 0, ByVal 0&, 0, 0, vbNullString, 0) <> 0 Then IsDriveAvailable = True
End Function
Function RemoveLastDir(ByVal sInput As String) As String
   Dim cnt As Long
   
    RemoveLastDir = sInput
    If VBA.Right$(sInput, 1) = "\" Then sInput = Left$(sInput, Len(sInput) - 1)
    For cnt = 0 To Len(sInput) - 1
        If Mid$(sInput, Len(sInput) - cnt, 1) = "\" Then
            RemoveLastDir = Left$(sInput, Len(sInput) - cnt)
            Exit For
        End If
    Next
End Function
Sub FillRemoteListView()
    Dim cnt As Long
    'While lvRemote.ListItems.Count > 0
        'lvRemote.ListItems.Remove lvRemote.ListItems.Count
    'Wend
    lvRemote.ListItems.Clear
    rfConnection.EnumFiles cFiles, cAttrs, cSize
    rfConnection.ClearCollection cRemAttrs
    lvRemote.ListItems.Add , , "..", , "Up"
    cRemAttrs.Add vbDirectory
    lstTemp.Clear
    For cnt = 1 To cFiles.Count
        If (cAttrs(cnt) And vbDirectory) = vbDirectory Then
            lstTemp.AddItem cFiles(cnt)
            cRemAttrs.Add vbDirectory
        End If
    Next cnt
    For cnt = 0 To lstTemp.ListCount - 1
        lvRemote.ListItems.Add , , lstTemp.List(cnt), , "Directory"
    Next cnt
    lstTemp.Clear
    For cnt = 1 To cFiles.Count
        If (cAttrs(cnt) And vbDirectory) <> vbDirectory Then
            lstTemp.AddItem cFiles(cnt) + "/" + CStr(cSize(cnt))
            cRemAttrs.Add vbNormal
        End If
    Next cnt
    For cnt = 0 To lstTemp.ListCount - 1
        lvRemote.ListItems.Add , , Left$(lstTemp.List(cnt), InStr(1, lstTemp.List(cnt), "/") - 1), , "File"
        lvRemote.ListItems.Item(lvRemote.ListItems.Count).SubItems(1) = VBA.Right$(lstTemp.List(cnt), Len(lstTemp.List(cnt)) - InStr(1, lstTemp.List(cnt), "/"))
    Next cnt
    txtRemPath.Text = rfConnection.GetCurrentDirectory
End Sub
Sub meStartFO()
    Dim hThread As Long, hThreadID As Long
    hThread = CreateThread(ByVal 0&, ByVal 0&, AddressOf StartFO, ByVal 0&, ByVal 0&, hThreadID)
    CloseHandle hThread
    'StartFO
End Sub
Public Sub SetEnabled(bEnabled As Boolean)
    lvRemote.Enabled = bEnabled
    cmdrMkDir.Enabled = bEnabled
    cmdrExec.Enabled = bEnabled
    cmdrRename.Enabled = bEnabled
    cmdrDelete.Enabled = bEnabled
    cmdrRefresh.Enabled = bEnabled
    cmdRetrieve.Enabled = bEnabled
    cmdPut.Enabled = bEnabled
    cmdCancel.Enabled = bEnabled
    cmdPutNew.Enabled = bEnabled
End Sub
Private Sub rfConnection_StatusChanged(NewStatus As tNewStatus, sOptionalInfo As String)
    If NewStatus = nsConnected Then
        SetEnabled True
    ElseIf NewStatus = nsDisconnected Then
        SetEnabled False
    End If
End Sub

