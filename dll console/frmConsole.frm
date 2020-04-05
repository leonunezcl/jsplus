VERSION 5.00
Begin VB.Form frmConsole 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Basic Console"
   ClientHeight    =   8790
   ClientLeft      =   2685
   ClientTop       =   1485
   ClientWidth     =   11775
   Icon            =   "frmConsole.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   11775
   Begin VB.ComboBox cboCommand 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Text            =   "Enter your command here"
      Top             =   8430
      Width           =   10575
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   8220
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   -15
      Width           =   11775
   End
   Begin VB.CheckBox cmdChDir 
      BackColor       =   &H00000000&
      Caption         =   "&ChDir"
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8205
      Width           =   1200
   End
   Begin VB.Label lblCurDir 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "C:\>"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8190
      Width           =   420
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Name:         FMain
'' Filename:     FMain.bas
'' Project:      CommandOutput sample 2
'' Author:       Mattias Sjögren (mattias@mvps.org)
''               http://www.msjogren.net/dotnet/
''
'' Description:  Startup form
''
'' Dependencies: MGetCmdOutput module (MGetCmdOutput.bas)
''
''
'' Copyright ©2000-2001, Mattias Sjögren
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'''''''''''''''''''''
'''   Constants   '''
'''''''''''''''''''''

Private Const MAX_PATH = 260

' SHBrowseForFolder flags
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_EDITBOX = &H10
Private Const BIF_VALIDATE = &H20
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)


'''''''''''''''''
'''   Types   '''
'''''''''''''''''

Private Type BROWSEINFO
  hwndOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type


''''''''''''''''''''
'''   Declares   '''
''''''''''''''''''''

Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" ( _
  lpbi As BROWSEINFO) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" ( _
  ByVal pidl As Long, _
  ByVal pszPath As String) As Long

Private Declare Sub CoTaskMemFree Lib "ole32" ( _
  pv As Any)

Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" ( _
  ByVal lpPathName As String) As Long

Private Declare Function GetVersion Lib "kernel32" () As Long


'''''''''''''''''''''''''''
'''   Private methods   '''
'''''''''''''''''''''''''''

Private Sub cmdChDir_Click()

  Dim bi As BROWSEINFO
  Dim pidl As Long
  
  
  If cmdChDir.Value = vbChecked Then
    
    cmdChDir.Value = vbUnchecked        ' don't stay pressed

    With bi
      .hwndOwner = Me.hWnd
      .lpszTitle = "Select a directory."
      .pszDisplayName = String$(MAX_PATH, 0)
      .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
                 BIF_VALIDATE Or BIF_USENEWUI
    End With
    
    ' show "browse for folder" dialog to select a directory
    pidl = SHBrowseForFolder(bi)
    
    If pidl Then
      If SHGetPathFromIDList(pidl, bi.pszDisplayName) Then
        CD Left$(bi.pszDisplayName, InStr(bi.pszDisplayName, vbNullChar) - 1)
      End If
      CoTaskMemFree ByVal pidl
    End If
  
    cboCommand.SetFocus
    
  End If
  
End Sub

Private Sub Form_Load()
  
  CD App.Path

End Sub

Private Sub cboCommand_KeyPress(KeyAscii As Integer)

  Dim sCmd As String
  
  
  If KeyAscii = 13 Then ' Enter was pressed
    ' The entered command should be executed by the default command processor, given by %COMSPEC%
    ' This allows us to type in commands such as DIR and TYPE.
    sCmd = Environ$("COMSPEC") & " /c " & cboCommand.Text
    
    ' If we are running Windows 9x, we have to launch the command using an
    ' intermediate Win32 console application (RedirStub.exe in this case),
    ' since Command.com is a 16-bit program. See KB article Q150956.
    If IsWin9x Then sCmd = "RedirStub " & sCmd
        
    txtOutput.Text = GetCommandOutput(sCmd, True, True)
    txtOutput.SelStart = Len(txtOutput.Text)    ' scroll textbox to last line
    cboCommand.AddItem cboCommand.Text, 0       ' add command to the Combo list
    cboCommand.Text = ""
  End If
  
End Sub

'
' Sub CD
'
' Description:  Sets the current directory of the process, and updates the "prompt" label
'
' sPath:  [in] Path of directory to switch to
'
Private Sub CD(sPath As String)

  If SetCurrentDirectory(sPath) Then lblCurDir.Caption = sPath & ">"
  
End Sub

'
' Function IsWin9x
'
' Description:  Returns True on Windows 9x, False on Windows NT/2000
'
Private Function IsWin9x() As Boolean
  IsWin9x = CBool(GetVersion() And &H80000000)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmConsole = Nothing
End Sub


