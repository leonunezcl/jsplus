VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Object = "{FCFAF346-DE8A-4FB6-8612-5000548EFDC7}#2.0#0"; "vbsListView6.ocx"
Begin VB.Form frmTidyOut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tidy Results"
   ClientHeight    =   8340
   ClientLeft      =   4095
   ClientTop       =   1980
   ClientWidth     =   12060
   Icon            =   "frmTidyOut.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "Save As .."
      Height          =   375
      Index           =   4
      Left            =   7080
      TabIndex        =   8
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Cut to Clipboard"
      Height          =   375
      Index           =   3
      Left            =   5280
      TabIndex        =   7
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Copy to Clipboard"
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   6
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Select All"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Close"
      Height          =   375
      Index           =   0
      Left            =   8880
      TabIndex        =   4
      Top             =   7800
      Width           =   1455
   End
   Begin vbalListViewLib6.vbalListViewCtl lvwMsg 
      Height          =   2865
      Left            =   45
      TabIndex        =   1
      Top             =   4755
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   5054
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   1
      MultiSelect     =   -1  'True
      LabelEdit       =   0   'False
      FullRowSelect   =   -1  'True
      AutoArrange     =   0   'False
      Appearance      =   0
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin CodeSenseCtl.CodeSense txtCode 
      Height          =   4245
      Left            =   45
      OleObjectBlob   =   "frmTidyOut.frx":000C
      TabIndex        =   0
      Top             =   285
      Width           =   11955
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Errors, Warnings , Suggestions  - Double click goto line"
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
      Left            =   45
      TabIndex        =   3
      Top             =   4545
      Width           =   4740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output Text"
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
      Left            =   45
      TabIndex        =   2
      Top             =   30
      Width           =   1020
   End
End
Attribute VB_Name = "frmTidyOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Unload Me
    ElseIf Index = 1 Then   'select all
        If Not frmMain.ActiveForm Is Nothing Then
            If frmMain.ActiveForm.Name = "frmEdit" Then
                txtCode.ExecuteCmd cmCmdSelectAll
            End If
        End If
    ElseIf Index = 2 Then   'copy
        If Not frmMain.ActiveForm Is Nothing Then
            If frmMain.ActiveForm.Name = "frmEdit" Then
                If txtCode.CanCopy Then
                    txtCode.Copy
                End If
            End If
        End If
    ElseIf Index = 3 Then   'cut
        If Not frmMain.ActiveForm Is Nothing Then
            If frmMain.ActiveForm.Name = "frmEdit" Then
                If txtCode.CanCut Then
                    txtCode.Cut
                End If
            End If
        End If
    ElseIf Index = 4 Then   'nuevo documento
        Dim archivotmp As String
        
        If Cdlg.VBGetSaveFileName(archivotmp, , , strGlosa(), , LastPath, , "htm", Me.hwnd) Then
            txtCode.SaveFile archivotmp, False
            frmMain.opeEdit archivotmp
        End If
    End If
    
End Sub

Private Sub Form_Load()

    Dim Archivo As String
    Dim linea As String
    Dim nFreeFile As Long
    Dim C As Integer
    
    util.Hourglass hwnd, True
    
    frmMain.ActiveForm.vbsMsg1.LoadTidyFile
    
    util.CenterForm Me
        
    txtCode.Language = "html"
        
    With lvwMsg
        .Columns.Add , "Line", "Line", , 700
        .Columns.Add , "Col", "Col", , 700
        .Columns.Add , "Warning", "Warning/Error", , 8000
    End With
    
    If ArchivoExiste2(util.StripPath(App.Path) & "tidy\output.tidy") Then
        ListaLangs.SetLang "output.tidy", txtCode
        txtCode.OpenFile util.StripPath(App.Path) & "tidy\output.tidy"
    ElseIf ArchivoExiste2(util.StripPath(App.Path) & "tidy\inputfile.tidy") Then
        ListaLangs.SetLang "inputfile.tidy", txtCode
        txtCode.OpenFile util.StripPath(App.Path) & "tidy\inputfile.tidy"
    Else
        MsgBox "Tidy Error. Output file is empty", vbCritical
    End If
        
    Archivo = util.StripPath(App.Path) & "tidy\errors.tidy"
            
    If ArchivoExiste2(Archivo) Then
        nFreeFile = FreeFile
        C = 1
        Open Archivo For Input As #nFreeFile
            Do While Not EOF(nFreeFile)
                Line Input #1, linea
                If IsNumeric(util.Explode(linea, 2, " ")) Then
                    lvwMsg.ListItems.Add , "k" & C, util.Explode(linea, 2, " ")
                    lvwMsg.ListItems(C).SubItems(1).Caption = util.Explode(linea, 4, " ")
                    lvwMsg.ListItems(C).SubItems(2).Caption = util.Explode(linea, 2, "-")
                    C = C + 1
                End If
            Loop
        Close #nFreeFile
    End If
        
    util.Hourglass hwnd, False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmTidyOut = Nothing
End Sub


Private Sub lvwMsg_ItemDblClick(ITem As vbalListViewLib6.cListItem)

    On Error Resume Next
    
    Dim linea As String
    'Dim num As Variant
    Dim Col As String
    
    If Not ITem Is Nothing Then
        linea = ITem.Text
        Col = ITem.SubItems(1).Caption
            
        Call txtCode.SetCaretPos(linea - 1, Col)
        Call txtCode_SelChange(txtCode)
        txtCode.SetFocus
    End If
    
    Err = 0
    
End Sub

Private Function txtCode_RClick(ByVal Control As CodeSenseCtl.ICodeSense) As Boolean
    txtCode_RClick = True
End Function

Private Sub txtCode_SelChange(ByVal Control As CodeSenseCtl.ICodeSense)

    On Error Resume Next
    
    Dim r As CodeSenseCtl.IRange
    'Dim Indice As Integer
    Dim colorh As Long
    
    Set r = Control.GetSel(True)
        
    colorh = Control.GetColor(cmClrHighlightedLine)
    Call Control.SetColor(cmClrHighlightedLine, Control.GetColor(cmClrWindow))
    Control.HighlightedLine = r.StartLineNo
    DoEvents
    Call Control.SetColor(cmClrHighlightedLine, colorh)
    
    Set r = Nothing
    
    Err = 0
    
End Sub

Private Function txtCode_ShowProps(ByVal Control As CodeSenseCtl.ICodeSense) As Boolean
    txtCode_ShowProps = True
End Function


