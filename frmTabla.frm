VERSION 5.00
Object = "{246E535D-09D2-4109-80DA-2FF183F4D185}#2.1#0"; "colorpick.ocx"
Begin VB.Form frmTabla 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Table"
   ClientHeight    =   5295
   ClientLeft      =   4395
   ClientTop       =   2250
   ClientWidth     =   6450
   Icon            =   "frmTabla.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&More"
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   39
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Events"
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   38
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   37
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   36
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Color Settings"
      Height          =   1020
      Index           =   2
      Left            =   45
      TabIndex        =   21
      Top             =   2985
      Width           =   6330
      Begin ColorPick.ClrPicker ClrPicker1 
         Height          =   300
         Index           =   0
         Left            =   2010
         TabIndex        =   32
         Top             =   225
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   4350
         MaxLength       =   9
         TabIndex        =   28
         Top             =   585
         Width           =   900
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   4350
         MaxLength       =   9
         TabIndex        =   26
         Top             =   240
         Width           =   900
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1095
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   24
         Top             =   585
         Width           =   900
      End
      Begin VB.TextBox txtColor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1095
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   22
         Top             =   240
         Width           =   900
      End
      Begin ColorPick.ClrPicker ClrPicker1 
         Height          =   300
         Index           =   1
         Left            =   2010
         TabIndex        =   33
         Top             =   570
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
      End
      Begin ColorPick.ClrPicker ClrPicker1 
         Height          =   300
         Index           =   2
         Left            =   5265
         TabIndex        =   34
         Top             =   225
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
      End
      Begin ColorPick.ClrPicker ClrPicker1 
         Height          =   300
         Index           =   3
         Left            =   5265
         TabIndex        =   35
         Top             =   570
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dark Border"
         Height          =   195
         Index           =   10
         Left            =   3420
         TabIndex        =   29
         Top             =   615
         Width           =   855
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Light Border"
         Height          =   195
         Index           =   9
         Left            =   3405
         TabIndex        =   27
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Border"
         Height          =   195
         Index           =   8
         Left            =   150
         TabIndex        =   25
         Top             =   615
         Width           =   465
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Background"
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   23
         Top             =   255
         Width           =   870
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Design Settings"
      Height          =   2130
      Index           =   1
      Left            =   45
      TabIndex        =   16
      Top             =   840
      Width           =   6330
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   255
         Index           =   5
         Left            =   5925
         TabIndex        =   41
         Top             =   1665
         Width           =   315
      End
      Begin VB.TextBox txtArchivo 
         Height          =   285
         Left            =   1215
         TabIndex        =   31
         Top             =   1650
         Width           =   4665
      End
      Begin VB.OptionButton opt 
         Caption         =   "Percentage"
         Height          =   195
         Index           =   1
         Left            =   4710
         TabIndex        =   9
         Top             =   720
         Value           =   -1  'True
         Width           =   1260
      End
      Begin VB.OptionButton opt 
         Caption         =   "Pixels"
         Height          =   195
         Index           =   0
         Left            =   4710
         TabIndex        =   8
         Top             =   510
         Width           =   735
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   3645
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "100"
         Top             =   510
         Width           =   975
      End
      Begin VB.CheckBox chk 
         Caption         =   "Set Width"
         Height          =   225
         Left            =   3645
         TabIndex        =   6
         Top             =   270
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.ComboBox cboAlign 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   255
         Width           =   1350
      End
      Begin VB.TextBox txtCellSpa 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1215
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "0"
         Top             =   1305
         Width           =   975
      End
      Begin VB.TextBox txtBorder 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1215
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "1"
         Top             =   615
         Width           =   975
      End
      Begin VB.TextBox txtCellPad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1215
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "0"
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Picture Background"
         Height          =   375
         Index           =   11
         Left            =   195
         TabIndex        =   30
         Top             =   1590
         Width           =   960
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Align"
         Height          =   195
         Index           =   5
         Left            =   195
         TabIndex        =   20
         Top             =   270
         Width           =   345
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cell Spacing"
         Height          =   195
         Index           =   4
         Left            =   195
         TabIndex        =   19
         Top             =   1335
         Width           =   885
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cell Padding"
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   18
         Top             =   990
         Width           =   885
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Border Size"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   17
         Top             =   660
         Width           =   810
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Events"
      Height          =   690
      Index           =   3
      Left            =   45
      TabIndex        =   14
      Top             =   4020
      Width           =   6330
      Begin VB.CommandButton cmd 
         Caption         =   "&Delete"
         Height          =   315
         Index           =   4
         Left            =   5595
         TabIndex        =   40
         Top             =   270
         Width           =   645
      End
      Begin VB.ComboBox cboEvents 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   705
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   255
         Width           =   4860
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   15
         Top             =   285
         Width           =   510
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Size Settings"
      Height          =   705
      Index           =   0
      Left            =   45
      TabIndex        =   11
      Top             =   90
      Width           =   6330
      Begin VB.TextBox txtRows 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "2"
         Top             =   255
         Width           =   975
      End
      Begin VB.TextBox txtColumns 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5235
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "2"
         Top             =   255
         Width           =   975
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rows"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   255
         Width           =   405
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Columns"
         Height          =   195
         Index           =   1
         Left            =   4185
         TabIndex        =   12
         Top             =   255
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function CrearTabla() As Boolean

    On Error GoTo ErrorCrearTabla
    
    Dim src As New cStringBuilder
    
    Dim k As Integer
    Dim j As Integer
    Dim rows As String
    Dim Columns As String
    Dim Align As String
    Dim Border As String
    Dim cellpad As String
    Dim cellspa As String
    Dim Width As String
    
    Dim background As String
    Dim BorderColor As String
    Dim bordercolorlight As String
    Dim bordercolordark As String
    Dim bgcolor As String
    
    util.Hourglass hwnd, True
                        
    Align = cboAlign.Text
    rows = txtRows.Text
    Columns = txtColumns.Text
    Border = txtBorder.Text
    cellpad = txtCellPad.Text
    cellspa = txtCellSpa.Text
    Width = txtWidth.Text
    background = txtArchivo.Text
    BorderColor = txtColor(1).Text
    bordercolorlight = txtColor(2).Text
    bordercolordark = txtColor(3).Text
    bgcolor = txtColor(0).Text
    
    If Align = "" Then
        MsgBox "Must select table align.", vbCritical
        cboAlign.SetFocus
        Exit Function
    End If
    
    If rows = "" Then
        MsgBox "Input table rows please.", vbCritical
        txtRows.SetFocus
        Exit Function
    End If
    
    If Columns = "" Then
        MsgBox "Input table columns please.", vbCritical
        txtColumns.SetFocus
        Exit Function
    End If
    
    If Border = "" Then
        MsgBox "Input table border please.", vbCritical
        txtBorder.SetFocus
        Exit Function
    End If
    
    If cellpad = "" Then
        MsgBox "Input table cellpadding please.", vbCritical
        txtCellPad.SetFocus
        Exit Function
    End If
    
    If cellspa = "" Then
        MsgBox "Input table cellspacing please.", vbCritical
        txtCellSpa.SetFocus
        Exit Function
    End If
    
    If cboAlign.ListIndex = -1 Then
        MsgBox "Select table align please.", vbCritical
        cboAlign.SetFocus
        Exit Function
    End If
    
    src.Append "<table border=" & Chr$(34) & Border & Chr$(34)
    
    If cboAlign.ListIndex <> 0 Then
        src.Append " align=" & Chr$(34) & Align & Chr$(34)
    End If
    
    src.Append " cellpadding=" & Chr$(34) & cellpad & Chr$(34)
    src.Append " cellspacing=" & Chr$(34) & cellspa & Chr$(34)
    
    If chk.Value Then
        If opt(1).Value Then
            src.Append " width=" & Chr$(34) & Width & "%" & Chr$(34)
        Else
            src.Append " width=" & Chr$(34) & Width & Chr$(34)
        End If
    End If
    
    If Len(BorderColor) > 0 Then
        src.Append " bordercolor=" & Chr$(34) & BorderColor & Chr$(34)
    End If
    
    If Len(bordercolorlight) > 0 Then
        src.Append " bordercolorlight=" & Chr$(34) & bordercolorlight & Chr$(34)
    End If
    
    If Len(bordercolordark) > 0 Then
        src.Append " bordercolordark=" & Chr$(34) & bordercolordark & Chr$(34)
    End If
    
    If Len(bgcolor) > 0 Then
        src.Append " bgcolor=" & Chr$(34) & bgcolor & Chr$(34)
    End If
    
    If Len(background) > 0 Then
        src.Append " background=" & Chr$(34) & txtArchivo.Text & Chr$(34)
    End If
    
    If Len(Trim$(CComHtmlAttrib.Output)) > 0 Then
        src.Append " " & Trim$(CComHtmlAttrib.Output)
    End If
    
    If Len(Trim$(CEventos.Output)) > 0 Then
        src.Append " " & Trim$(CEventos.Output)
    End If
    src.Append ">" & vbNewLine
        
    For k = 1 To txtRows.Text
        src.Append vbTab & "<tr>" & vbNewLine
        For j = 1 To txtColumns.Text
            If chk.Value Then
                If opt(1).Value Then
                    If Width = "100" Then
                        src.Append vbTab & vbTab & "<td>&nbsp;</td>" & vbNewLine
                    Else
                        src.Append vbTab & vbTab & "<td width=" & Chr$(34) & Width & "%" & Chr$(34) & ">&nbsp;</td>" & vbNewLine
                    End If
                Else
                    src.Append vbTab & vbTab & "<td>&nbsp;</td>" & vbNewLine
                End If
            Else
                src.Append vbTab & vbTab & "<td>&nbsp;</td>" & vbNewLine
            End If
        Next j
        src.Append vbTab & "</tr>" & vbNewLine
    Next k
    src.Append "</table>" & vbNewLine
        
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(src.ToString)
    End If
    
    Set src = Nothing
    
    util.Hourglass hwnd, False
    
    CrearTabla = True
    
    Exit Function
ErrorCrearTabla:
    MsgBox "CrearTabla : " & Err & " " & Error$, vbCritical
    Exit Function
    
End Function

Private Sub SelBackground()

    Dim glosa As String
    Dim Archivo As String
    Dim LastPath As String
    
    glosa = "CompuServe Graphics Interchange (*.gif)|*.gif|"
    glosa = glosa & "JPG (*.jpg)|*.jpg|"
    glosa = glosa & "All files (*.*)|*.*"
    
    LastPath = App.Path
    
    If Not Cdlg.VBGetOpenFileName(Archivo, , , , , , glosa, , LastPath, "Open Image", "gif") Then
        Exit Sub
    End If
        
    txtArchivo.Text = util.VBArchivoSinPath(Archivo)
    
End Sub

Private Sub chk_Click()

    Dim ret As Boolean
    
    If chk.Value Then
        ret = True
    Else
        ret = False
    End If
    
    txtWidth.Enabled = ret
    opt(0).Enabled = ret
    opt(1).Enabled = ret
    
End Sub


Private Sub ClrPicker1_ColorSelected(Index As Integer, m_Color As stdole.OLE_COLOR, m_Code As String)
    txtColor(Index).Text = m_Code
    txtColor(Index).Tag = m_Color
End Sub


Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If CrearTabla() Then
            Unload Me
        End If
    ElseIf Index = 1 Then
        Unload Me
    ElseIf Index = 2 Then
        '#If LITE = 1 Then
        '    MsgBox C_MSG, vbInformation
        '#Else
            frmEvents.html_tag = "Image"
            frmEvents.Show vbModal
            Call CEventos.Attach(Me.cboEvents)
        '#End If
    ElseIf Index = 3 Then
        frmCommonHtml.html_tag = "table"
        frmCommonHtml.Show vbModal
    ElseIf Index = 4 Then
        If cboEvents.ListIndex <> -1 Then
            CEventos.Remove cboEvents.Text
            cboEvents.RemoveItem cboEvents.ListIndex
        End If
    ElseIf Index = 5 Then
        Call SelBackground
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    util.CenterForm Me
    
    util.SetNumber txtRows.hwnd
    util.SetNumber txtColumns.hwnd
    util.SetNumber txtBorder.hwnd
    util.SetNumber txtCellPad.hwnd
    util.SetNumber txtCellSpa.hwnd
    util.SetNumber txtWidth.hwnd
    
    cboAlign.AddItem "default"
    cboAlign.AddItem "left"
    cboAlign.AddItem "center"
    cboAlign.AddItem "right"
    cboAlign.AddItem "justify"
    
    CEventos.Clear
    CComHtmlAttrib.Clear
    
    ClrPicker1(0).PathPaleta = App.Path & "\pal\256c.pal"
    ClrPicker1(1).PathPaleta = App.Path & "\pal\256c.pal"
    ClrPicker1(2).PathPaleta = App.Path & "\pal\256c.pal"
    ClrPicker1(3).PathPaleta = App.Path & "\pal\256c.pal"
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmTabla = Nothing
End Sub


