VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.UserControl vbsMsg 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LockControls    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdOpc 
      Height          =   300
      Left            =   30
      Picture         =   "vbsMsg.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Configure Analyzer Settings"
      Top             =   1245
      Width           =   315
   End
   Begin VB.CommandButton cmdHtml 
      Height          =   300
      Left            =   30
      Picture         =   "vbsMsg.ctx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Preview in HTML"
      Top             =   585
      Width           =   315
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   300
      Left            =   30
      Picture         =   "vbsMsg.ctx":06D4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Send output to printer"
      Top             =   915
      Width           =   315
   End
   Begin VB.CommandButton cmdSave 
      Height          =   300
      Left            =   30
      Picture         =   "vbsMsg.ctx":081E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Save result to file ..."
      Top             =   249
      Width           =   315
   End
   Begin MSComctlLib.ListView lvwJSlintMsg 
      Height          =   1050
      Left            =   315
      TabIndex        =   4
      Top             =   930
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1852
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Line"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Column"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Error"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Expression"
         Object.Width           =   8819
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2115
      Top             =   1515
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "vbsMsg.ctx":0DA8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwTidyMsg 
      Height          =   1170
      Left            =   585
      TabIndex        =   3
      Top             =   2190
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   2064
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Line"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Col"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Warning/Error"
         Object.Width           =   14111
      EndProperty
   End
   Begin VB.PictureBox cmd 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4455
      Picture         =   "vbsMsg.ctx":10C2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   1545
      Width           =   240
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   1935
      Left            =   375
      TabIndex        =   0
      Top             =   315
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   3413
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tidy Output"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "JavaScript Errors"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblerrors 
      AutoSize        =   -1  'True
      Caption         =   "0 Errors"
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
      Width           =   675
   End
   Begin VB.Image imgDown 
      Height          =   240
      Left            =   4125
      Picture         =   "vbsMsg.ctx":120C
      Top             =   450
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgUp 
      Height          =   240
      Left            =   4110
      Picture         =   "vbsMsg.ctx":1356
      Top             =   90
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "vbsMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_EdtCtl As Control
Private m_NumError As Integer
Private m_FileName As String

Event HidePanel()
Event ShowPanel()

Public Sub Load()
        
    Set cmd.Picture = imgDown.Picture
    cmd.Tag = "hide"
    cmd.ToolTipText = "Hide Panel"
    
End Sub







Public Sub LoadJsLintFile()

    Dim output_file As String
    Dim linea As String
    Dim nFreeFile As Long
    Dim nlinea As Integer
    Dim Col As Integer
    Dim nlinea1 As Integer
    Dim col1 As Integer
    Dim analizando As Boolean
    Dim c As Integer
    
    output_file = util.StripPath(App.Path) & "jslint\output.txt"

    lvwJSlintMsg.ListItems.Clear
    
    c = 1
    m_NumError = 0
    
    If ArchivoExiste2(output_file) Then
        'abrir el archivo de salida y analizarlo
        nFreeFile = FreeFile
        Open output_file For Input As #nFreeFile
          Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, linea
            If Left$(LCase$(linea), 15) = "<div id=output>" Then
                analizando = True
            End If
            
            If analizando Then
                'If InStr(linea, "TABLE") Then
                '    Debug.Print "STOP!"
                'End If
                
                If LCase$(Left$(linea, 7)) = "<table>" Then
                    Exit Do
                End If
                
                If LCase$(linea) <> "<div id=output>" Then
                    If InStr(linea, "</P>") Then
                        linea = Trim$(Mid$(linea, 4))
                        linea = Trim$(Left$(linea, InStr(linea, "</P>") - 1))
                                            
                        If InStr(linea, "|") Then
                           If c = 1 Then
                               nlinea1 = util.Explode(linea, 1, "|")
                               col1 = util.Explode(linea, 2, "|")
                           End If
                           
                           nlinea = util.Explode(linea, 1, "|")
                           Col = util.Explode(linea, 2, "|")
                           
                           lvwJSlintMsg.ListItems.Add , "k" & c, CStr(nlinea), 1, 1
                           lvwJSlintMsg.ListItems(c).SubItems(1) = CStr(Col)
                           lvwJSlintMsg.ListItems(c).SubItems(2) = util.Explode(linea, 3, "|")
                           lvwJSlintMsg.ListItems(c).SubItems(3) = util.Explode(linea, 4, "|")
                               
                           c = c + 1
                        End If
                    Else
                        Exit Do
                    End If
                End If
            End If
          Loop
        Close #nFreeFile
    End If
    
    If c > 1 Then
        On Error Resume Next
        tabMain.Tabs(2).Selected = True
        Call m_EdtCtl.SetCaretPos(nlinea1 - 1, col1)
        UserControl.Parent.SelChange
        m_EdtCtl.SetFocus
        Err = 0
    End If
    
    lblerrors.Caption = c - 1 & " Errors"
    
End Sub
Public Sub LoadTidyFile()

    Dim Archivo As String
    Dim linea As String
    Dim nFreeFile As Long
    Dim c As Integer
    
    util.Hourglass hwnd, True
        
    lvwTidyMsg.ListItems.Clear
        
    Archivo = util.StripPath(App.Path) & "tidy\errors.tidy"
            
    If ArchivoExiste2(Archivo) Then
        nFreeFile = FreeFile
        c = 1
        Open Archivo For Input As #nFreeFile
            Do While Not EOF(nFreeFile)
                Line Input #1, linea
                
                If Left$(linea, 4) = "line" Then
                    lvwTidyMsg.ListItems.Add , "k" & c, util.Explode(linea, 2, " "), 1, 1
                    lvwTidyMsg.ListItems(c).SubItems(1) = util.Explode(linea, 4, " ")
                    lvwTidyMsg.ListItems(c).SubItems(2) = util.Explode(linea, 2, "-")
                    c = c + 1
                End If
            Loop
        Close #nFreeFile
    End If
    
    util.Hourglass hwnd, False

End Sub

Private Sub cmd_Click()

    If cmd.Tag = "hide" Then
        tabMain.Visible = False
        cmd.Picture = imgup.Picture
        cmd.Tag = "show"
        cmd.ToolTipText = "Show Panel"
        RaiseEvent HidePanel
    Else
        cmd.Picture = imgDown.Picture
        cmd.Tag = "hide"
        cmd.ToolTipText = "Hide Panel"
        tabMain.Visible = True
        RaiseEvent ShowPanel
    End If
    
End Sub





Private Sub cmdHtml_Click()

    On Local Error GoTo ErrorImprimir
    
    Dim Archivo As String
    Dim k As Integer
    Dim Itmx As ListItem
    Dim nFreeFile As Integer
    Dim Fuente As String
    Dim ret As String
    
    If tabMain.SelectedItem.Index = 1 Then
        If lvwTidyMsg.ListItems.count = 0 Then Exit Sub
    Else
        If lvwJSlintMsg.ListItems.count = 0 Then Exit Sub
    End If
    
    ret = "HTML Files (*.html)|*.html|"
    ret = ret & "All Files (*.*)|*.*"
            
    Call Cdlg.VBGetSaveFileName(Archivo, , , ret, , LastPath, "Save As ...", "html", frmMain.hwnd)
    
    If Len(Archivo) = 0 Then Exit Sub
    
    nFreeFile = FreeFile
    
    Call Hourglass(hwnd, True)
    
    Fuente = Replace("<font face='Verdana, Arial, Helvetica, sans-serif' size='1'>", "'", Chr$(34))
    
    Open Archivo For Output As #nFreeFile
        'cabezera del archivo
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head>"
        Print #nFreeFile, "<title>JavaScript Plus Report</title>"
        Print #nFreeFile, Replace("<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>", "'", Chr$(34))
        Print #nFreeFile, "</head>"
        
        'titulo del reporte
        Print #nFreeFile, Replace("<body bgcolor='#FFFFFF' text='#000000'>", "'", Chr$(34))
        Print #nFreeFile, Fuente
        Print #nFreeFile, "<p><b>Report Date : " & Now & " </b></p>"
        Print #nFreeFile, "<p><b>File : " & util.VBArchivoSinPath(m_FileName) & "</b></p>"
        Print #nFreeFile, "</font>"
        
        'generar titulos
        Print #nFreeFile, Replace("<table width='97%' border='1' bordercolor='#FFFFFF'>", "'", Chr$(34))
        Print #nFreeFile, Replace("<tr bgcolor='#999999' bordercolor='#000000'>", "'", Chr$(34))
        
        If tabMain.SelectedItem.Index = 1 Then
            Print #nFreeFile, Replace("<td width='03%'><b>" & Fuente & "Number</font></b></td>", "'", Chr$(34))
            Print #nFreeFile, Replace("<td width='03%'><b>" & Fuente & "Line</font></b></td>", "'", Chr$(34))
            Print #nFreeFile, Replace("<td width='03%'><b>" & Fuente & "Column</font></b></td>", "'", Chr$(34))
            Print #nFreeFile, Replace("<td width='25%'><b>" & Fuente & "Warning/Error</font></b></td>", "'", Chr$(34))
        Else
            Print #nFreeFile, Replace("<td width='03%'><b>" & Fuente & "Number</font></b></td>", "'", Chr$(34))
            Print #nFreeFile, Replace("<td width='03%'><b>" & Fuente & "Line</font></b></td>", "'", Chr$(34))
            Print #nFreeFile, Replace("<td width='03%'><b>" & Fuente & "Column</font></b></td>", "'", Chr$(34))
            Print #nFreeFile, Replace("<td width='25%'><b>" & Fuente & "Error</font></b></td>", "'", Chr$(34))
            Print #nFreeFile, Replace("<td width='25%'><b>" & Fuente & "Expression</font></b></td>", "'", Chr$(34))
        End If
        Print #nFreeFile, "</tr>"
        
        If tabMain.SelectedItem.Index = 1 Then
            For k = 1 To lvwTidyMsg.ListItems.count
                Set Itmx = lvwTidyMsg.ListItems(k)
            
                'imprimir informacion
                Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
            
                'correlativo
                Print #nFreeFile, Replace("<td width='03%' height='18'>" & Fuente & k & "</font></td>", "'", Chr$(34))
                            
                'linea
                Print #nFreeFile, Replace("<td width='03%' height='18'><b>" & Fuente & Itmx.Text & "</font></b></td>", "'", Chr$(34))
            
                'columna
                Print #nFreeFile, Replace("<td width='03%' height='18'>" & Fuente & Itmx.SubItems(1) & "</font></td>", "'", Chr$(34))
                        
                'warning/error
                Print #nFreeFile, Replace("<td width='25%' height='18'><b>" & Fuente & Itmx.SubItems(2) & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, "</tr>"
            Next k
        Else
            For k = 1 To lvwJSlintMsg.ListItems.count
                Set Itmx = lvwJSlintMsg.ListItems(k)
            
                'imprimir informacion
                Print #nFreeFile, Replace("<tr bordercolor='#000000'>", "'", Chr$(34))
            
                'correlativo
                Print #nFreeFile, Replace("<td width='03%' height='18'>" & Fuente & k & "</font></td>", "'", Chr$(34))
                            
                'linea
                Print #nFreeFile, Replace("<td width='03%' height='18'><b>" & Fuente & Itmx.Text & "</font></b></td>", "'", Chr$(34))
            
                'columna
                Print #nFreeFile, Replace("<td width='03%' height='18'>" & Fuente & Itmx.SubItems(1) & "</font></td>", "'", Chr$(34))
                        
                'error
                Print #nFreeFile, Replace("<td width='25%' height='18'><b>" & Fuente & Itmx.SubItems(2) & "</font></b></td>", "'", Chr$(34))
                
                'expression
                Print #nFreeFile, Replace("<td width='25%' height='18'><b>" & Fuente & Itmx.SubItems(3) & "</font></b></td>", "'", Chr$(34))
                Print #nFreeFile, "</tr>"
            Next k
        End If
        Print #nFreeFile, "</table>"
                
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
    Close #nFreeFile
        
    GoTo SalirImprimir
    
ErrorImprimir:
    Resume SalirImprimir
    
SalirImprimir:
    Call Hourglass(hwnd, False)
    Err = 0
    
End Sub

Private Sub cmdOpc_Click()

    If tabMain.SelectedItem.Index = 1 Then
        frmTidyConfig.Show vbModal
    Else
        frmjslitopt.Show vbModal
    End If

End Sub


Private Sub cmdPrint_Click()
    cmdHtml_Click
End Sub


Private Sub cmdSave_Click()

    Dim Archivo As String
    Dim nFreeFile As Long
    Dim k As Integer
    Dim ret As String
    
    If tabMain.SelectedItem.Index = 1 Then
        If lvwTidyMsg.ListItems.count = 0 Then Exit Sub
    Else
        If lvwJSlintMsg.ListItems.count = 0 Then Exit Sub
    End If
    
    ret = "Text Files (*.txt)|*.txt|"
    ret = ret & "All Files (*.*)|*.*"
            
    Call Cdlg.VBGetSaveFileName(Archivo, , , ret, , LastPath, "Save result ...", "txt", frmMain.hwnd)
        
    If Len(Archivo) > 0 Then
        nFreeFile = FreeFile
        Open Archivo For Output As #nFreeFile
        If tabMain.SelectedItem.Index = 1 Then
            For k = 1 To lvwTidyMsg.ListItems.count
                Print #nFreeFile, "Line : " & lvwJSlintMsg.ListItems(k).Text & " Column : " & lvwJSlintMsg.ListItems(k).SubItems(1) & _
                                  " Warning/Error : " & lvwJSlintMsg.ListItems(k).SubItems(2)
            Next k
        Else
            For k = 1 To lvwJSlintMsg.ListItems.count
                Print #nFreeFile, "Line : " & lvwJSlintMsg.ListItems(k).Text & " Column : " & lvwJSlintMsg.ListItems(k).SubItems(1) & _
                                  " Error : " & lvwJSlintMsg.ListItems(k).SubItems(2) & " Expression :" & lvwJSlintMsg.ListItems(k).SubItems(3)
            Next k
        End If
        Close #nFreeFile
        MsgBox "Finished", vbInformation
    End If
    
End Sub
Private Sub lvwJSlintMsg_ItemClick(ByVal ITem As MSComctlLib.ListItem)

    Dim linea As Long
    Dim Col As Long
    
    On Error Resume Next
    If Not ITem Is Nothing Then
        linea = ITem.Text
        Col = ITem.SubItems(1)
        
        Call m_EdtCtl.SetCaretPos(linea - 1, Col)
        UserControl.Parent.SelChange
        m_EdtCtl.SetFocus
    End If
    Err = 0
    
End Sub
Private Sub lvwTidyMsg_ItemClick(ByVal ITem As MSComctlLib.ListItem)

    On Error Resume Next

    Dim Num As Long
    Dim Col As Long

    If Not ITem Is Nothing Then
        Num = ITem.Text
        Col = ITem.SubItems(1)
        Call m_EdtCtl.SetCaretPos(Num - 1, Col)
        UserControl.Parent.SelChange
        UserControl.Parent.SetFocus
        m_EdtCtl.SetFocus
    End If
    
    Err = 0
    
End Sub


Private Sub tabMain_Click()
    If tabMain.SelectedItem.Index = 1 Then
        lvwTidyMsg.Visible = True
        lvwTidyMsg.ZOrder 0
    Else
        lvwJSlintMsg.Visible = True
        lvwJSlintMsg.ZOrder 0
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    tabMain.Move 375, 240, UserControl.Width, UserControl.Height - 260
    lvwTidyMsg.Move 375 + 40, 580, tabMain.Width - 500, tabMain.Height - 400
    lvwJSlintMsg.Move 375 + 40, 580, tabMain.Width - 500, tabMain.Height - 400
    cmd.Move UserControl.Width - cmd.Height - 50, lblerrors.Top '+ 30
    Err = 0
End Sub


Public Property Set CtlEdit(ByVal pEdtCtl As Control)
    Set m_EdtCtl = pEdtCtl
End Property

Public Property Get NumErrors() As Long
    NumErrors = m_NumError
End Property

Public Property Let NumErrors(ByVal pNumErrors As Long)
    m_NumError = pNumErrors
End Property

Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal pFileName As String)
    m_FileName = pFileName
End Property
