VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Apply CRC32 "
   ClientHeight    =   3675
   ClientLeft      =   2895
   ClientTop       =   3240
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   7305
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   4080
      Pattern         =   "*.EXE"
      TabIndex        =   3
      Top             =   360
      Width           =   3135
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3855
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   7095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Locate and double-click on your project exe file to calculate its checksum and append it to the file"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Dir1_Change()
On Error Resume Next
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
Dim TheFile As String
TheFile = File1.Path
If Right(TheFile, 1) <> "\" Then TheFile = TheFile & "\"
TheFile = TheFile & File1.FileName
    
    Dim lCrc32Value As Long
    Dim CRCStr As String * 8
    Dim FL As Long  'file length
    On Error Resume Next
    Dim FileStr$
    FL = FileLen(TheFile)
    FileStr$ = String(FL, 0)
    Open TheFile For Binary As #1
     Get #1, 1, FileStr$
    Close #1
    lCrc32Value = InitCrc32()
    lCrc32Value = AddCrc32(FileStr$, lCrc32Value)
    Dim RealCRC As String * 8
    RealCRC = CStr(Hex$(GetCrc32(lCrc32Value)))
    'MsgBox "Real CRC=" & RealCRC & vbCrLf & "File CRC=" & CRCStr, vbInformation + vbOKOnly, "CRC32 Results"
    Open TheFile For Binary As #1
     Put #1, FL + 1, RealCRC
    Close #1
    MsgBox "CRC32 checksum " & RealCRC & " has been appended to the end of this file!", vbInformation + vbOKOnly, "Success!"
    
End Sub

Private Sub Form_Load()
    Drive1.Drive = "c:"
    Dir1.Path = "c:\lnunez"
End Sub


