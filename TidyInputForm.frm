VERSION 5.00
Begin VB.Form TidyInputForm 
   Caption         =   "TidyInputForm"
   ClientHeight    =   7200
   ClientLeft      =   2595
   ClientTop       =   2160
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   4980
   Begin VB.CommandButton DoTidyStrings 
      Caption         =   "Tidy Strings!"
      Height          =   432
      Left            =   3660
      TabIndex        =   18
      Top             =   3345
      Width           =   1272
   End
   Begin VB.TextBox MarkupOutput 
      Height          =   1212
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   5400
      Width           =   4752
   End
   Begin VB.TextBox MarkupInput 
      Height          =   1212
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   3840
      Width           =   4752
   End
   Begin VB.CommandButton Browse4Error 
      Caption         =   "Choose..."
      Height          =   312
      Left            =   2280
      TabIndex        =   12
      Top             =   2700
      Width           =   972
   End
   Begin VB.TextBox ErrorFile 
      Height          =   288
      Left            =   120
      TabIndex        =   11
      Top             =   3060
      Width           =   3132
   End
   Begin VB.CommandButton DoTidy 
      Caption         =   "&Tidy!"
      Height          =   432
      Left            =   3780
      TabIndex        =   9
      Top             =   225
      Width           =   1092
   End
   Begin VB.CommandButton Browse4Output 
      Caption         =   "Choose..."
      Height          =   312
      Left            =   2280
      TabIndex        =   8
      Top             =   1920
      Width           =   972
   End
   Begin VB.TextBox OutputFile 
      Height          =   288
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   3132
   End
   Begin VB.CommandButton Browse4Input 
      Caption         =   "Browse..."
      Height          =   312
      Left            =   2280
      TabIndex        =   5
      Top             =   1080
      Width           =   972
   End
   Begin VB.TextBox InputFile 
      Height          =   288
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   3132
   End
   Begin VB.CommandButton Browse4Config 
      Caption         =   "Browse..."
      Height          =   312
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   972
   End
   Begin VB.TextBox ConfigFile 
      Height          =   288
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3132
   End
   Begin VB.Label Label3 
      Caption         =   "Output"
      Height          =   192
      Left            =   180
      TabIndex        =   17
      Top             =   5160
      Width           =   2232
   End
   Begin VB.Label Label2 
      Caption         =   "Input"
      Height          =   192
      Left            =   180
      TabIndex        =   15
      Top             =   3540
      Width           =   2172
   End
   Begin VB.Label StatusInfo 
      Caption         =   "Choose input and output files and press ""Tidy!"""
      Height          =   372
      Left            =   60
      TabIndex        =   13
      Top             =   6720
      Width           =   4812
   End
   Begin VB.Label Label1 
      Caption         =   "Error File"
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   2700
      Width           =   1632
   End
   Begin VB.Label OutputLabel 
      Caption         =   "Output File"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1632
   End
   Begin VB.Label InputLabel 
      Caption         =   "Input File"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1572
   End
   Begin VB.Label ConfigLabel 
      Caption         =   "Configuration File"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1632
   End
End
Attribute VB_Name = "TidyInputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents tdoc As TidyDocument
Attribute tdoc.VB_VarHelpID = -1
Private Sub Browse4Config_Click()

'OpenConfigDlg.CancelError = True
On Error GoTo SkipIt

'OpenConfigDlg.ShowOpen
'ConfigFile = OpenConfigDlg.FileName
SkipIt:

End Sub

Private Sub Browse4Error_Click()

'ChooseErrorDlg.CancelError = True
On Error GoTo SkipIt

'ChooseErrorDlg.ShowOpen
'ErrorFile = ChooseErrorDlg.FileName
SkipIt:

End Sub

Private Sub Browse4Input_Click()

'OpenInputDlg.CancelError = True
On Error GoTo SkipIt

'OpenInputDlg.ShowOpen
'InputFile = OpenInputDlg.FileName
SkipIt:

End Sub


Private Sub Browse4Output_Click()

'ChooseOutputDlg.CancelError = True
On Error GoTo SkipIt

'ChooseOutputDlg.ShowOpen
'OutputFile = ChooseOutputDlg.FileName
SkipIt:

End Sub

Private Sub DoTidy_Click()
  Dim stat As Long, ok As Long
  
  On Error GoTo DamnIt
  stat = 0
  If Len(ErrorFile) > 0 Then
    stat = tdoc.SetErrorFile(ErrorFile)
  End If
  If stat >= 0 Then
      stat = tdoc.LoadConfig(ConfigFile)
  End If
  If stat >= 0 Then
    stat = tdoc.ParseFile(InputFile)
  End If
  If stat >= 0 Then
    stat = tdoc.CleanAndRepair()
  End If
  If stat >= 0 Then
    stat = tdoc.RunDiagnostics()
  End If
  If stat >= 0 Then
    stat = tdoc.SaveFile(OutputFile)
  End If
  GoTo Done
  
DamnIt:
  MsgBox "Major FuBar running Tidy"
  
Done:
End Sub

Private Sub DoTidyStrings_Click()
  Dim stat As Long, ok As Long
  Dim Msg As String
  
  Msg = StatusInfo
  StatusInfo = ""
  
  On Error GoTo DamnIt
  stat = 0
  If Len(ErrorFile) > 0 Then
    stat = tdoc.SetErrorFile(ErrorFile)
  End If
  If stat >= 0 Then
    stat = tdoc.LoadConfig(ConfigFile)
  End If
  If stat >= 0 Then
    stat = tdoc.ParseString(MarkupInput)
  End If
  If stat >= 0 Then
    stat = tdoc.CleanAndRepair()
  End If
  If stat >= 0 Then
    stat = tdoc.RunDiagnostics()
  End If
  If stat > 1 Then
    ok = tdoc.SetOptBool(TidyForceOutput, True)
  End If
  If stat >= 0 Then
    ok = tdoc.SetOptInt(TidyOutputBOM, False)
    MarkupOutput = tdoc.SaveString()
  End If
  If StatusInfo = "" Then
    StatusInfo = Msg
  End If
  
  GoTo Done
  
DamnIt:
  MsgBox "Major FuBar running Tidy"
  
Done:
End Sub

Private Sub Form_Load()
    Set tdoc = New TidyDocument
End Sub

Private Sub tdoc_OnMessage(ByVal level As TidyReportLevel, ByVal line As Long, ByVal col As Long, ByVal Msg As String)
  Dim lvl As String, lin As String
  
  If level = TidyInfo Then
    lvl = "Info: "
  ElseIf level = TidyAccess Then
    lvl = "Access: "
  ElseIf level = TidyWarning Then
    lvl = "Warning: "
  ElseIf level = TidyConfig Then
    lvl = "Config: "
  ElseIf level = TidyError Then
    lvl = "Error: "
  ElseIf level = TidyBadDocument Then
    lvl = "Doc: "
  ElseIf level = TidyFatal Then
    lvl = "Fatal: "
  Else
    lvl = "???: "
  End If
  
  If line > 0 Then
    lin = lvl & "Line " & line & "Col " & col & ", " & Msg
  Else
    lin = lvl & Msg
  End If
  
  If Len(StatusInfo) > 0 Then
    StatusInfo = StatusInfo & Chr(13) & Chr(10) & lin
  Else
    StatusInfo = lin
  End If
End Sub
