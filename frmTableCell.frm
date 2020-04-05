VERSION 5.00
Object = "{246E535D-09D2-4109-80DA-2FF183F4D185}#2.1#0"; "colorpick.ocx"
Begin VB.Form frmTableCell 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cell"
   ClientHeight    =   5490
   ClientLeft      =   3135
   ClientTop       =   2925
   ClientWidth     =   6570
   Icon            =   "frmTableCell.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   37
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   36
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Background"
      Height          =   1080
      Index           =   2
      Left            =   45
      TabIndex        =   30
      Top             =   3885
      Width           =   6480
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   6075
         TabIndex        =   38
         Top             =   645
         Width           =   315
      End
      Begin VB.TextBox txtpic 
         Height          =   285
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   630
         Width           =   4515
      End
      Begin VB.TextBox txtBorColor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   285
         Width           =   1140
      End
      Begin ColorPick.ClrPicker ClrBorPicker1 
         Height          =   285
         Index           =   3
         Left            =   2655
         TabIndex        =   18
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Picture:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   150
         TabIndex        =   32
         Top             =   630
         Width           =   540
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   150
         TabIndex        =   31
         Top             =   330
         Width           =   405
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   2730
      Index           =   1
      Left            =   60
      TabIndex        =   23
      Top             =   15
      Width           =   6480
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   5490
         ScaleHeight     =   465
         ScaleWidth      =   915
         TabIndex        =   33
         Top             =   720
         Width           =   915
         Begin VB.OptionButton optHPix 
            Caption         =   "pixels"
            Height          =   210
            Index           =   0
            Left            =   15
            TabIndex        =   35
            Top             =   15
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optHPix 
            Caption         =   "percent"
            Height          =   210
            Index           =   1
            Left            =   15
            TabIndex        =   34
            Top             =   210
            Width           =   915
         End
      End
      Begin VB.CheckBox chknowra 
         Caption         =   "No wrap"
         Height          =   210
         Left            =   150
         TabIndex        =   11
         Top             =   2250
         Width           =   1005
      End
      Begin VB.CheckBox chkHeader 
         Caption         =   "Header cell"
         Height          =   210
         Left            =   150
         TabIndex        =   10
         Top             =   1935
         Width           =   1260
      End
      Begin VB.TextBox txtrowspa 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1500
         Width           =   1140
      End
      Begin VB.TextBox txtcolspa 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1125
         Width           =   1140
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   765
         Width           =   855
      End
      Begin VB.OptionButton optWPix 
         Caption         =   "percent"
         Height          =   210
         Index           =   1
         Left            =   5490
         TabIndex        =   8
         Top             =   480
         Width           =   915
      End
      Begin VB.OptionButton optWPix 
         Caption         =   "pixels"
         Height          =   210
         Index           =   0
         Left            =   5490
         TabIndex        =   7
         Top             =   285
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   300
         Width           =   855
      End
      Begin VB.ComboBox cbovalign 
         Height          =   315
         Left            =   1515
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cbohalign 
         Height          =   315
         Left            =   1515
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   285
         Width           =   1335
      End
      Begin jsplus.ctxUpDown updrowspa 
         Height          =   300
         Left            =   2670
         Top             =   1485
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   529
      End
      Begin jsplus.ctxUpDown updcolspa 
         Height          =   300
         Left            =   2670
         Top             =   1110
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   529
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rows spanned:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   150
         TabIndex        =   29
         Top             =   1515
         Width           =   1110
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Columns spanned:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   28
         Top             =   1155
         Width           =   1305
      End
      Begin jsplus.ctxUpDown updheight 
         Height          =   300
         Left            =   5235
         Top             =   765
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   529
      End
      Begin jsplus.ctxUpDown updwidth 
         Height          =   300
         Left            =   5250
         Top             =   300
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   529
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   3405
         TabIndex        =   27
         Top             =   765
         Width           =   510
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   3405
         TabIndex        =   26
         Top             =   315
         Width           =   465
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vertical align:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   25
         Top             =   720
         Width           =   945
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Horizontal align:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   24
         Top             =   315
         Width           =   1125
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Borders"
      Height          =   1080
      Index           =   0
      Left            =   45
      TabIndex        =   9
      Top             =   2775
      Width           =   6480
      Begin VB.TextBox txtBorColor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   645
         Width           =   1140
      End
      Begin VB.TextBox txtBorColor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   1140
      End
      Begin VB.TextBox txtBorColor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   285
         Width           =   1140
      End
      Begin ColorPick.ClrPicker ClrBorPicker1 
         Height          =   285
         Index           =   0
         Left            =   2655
         TabIndex        =   12
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
      End
      Begin ColorPick.ClrPicker ClrBorPicker1 
         Height          =   285
         Index           =   1
         Left            =   5520
         TabIndex        =   14
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
      End
      Begin ColorPick.ClrPicker ClrBorPicker1 
         Height          =   285
         Index           =   2
         Left            =   5520
         TabIndex        =   16
         Top             =   645
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dark Border:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3405
         TabIndex        =   22
         Top             =   690
         Width           =   900
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Light Border:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3405
         TabIndex        =   21
         Top             =   315
         Width           =   900
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   20
         Top             =   300
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmTableCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub insertar_celda()

    Dim src As New cStringBuilder
    
    If chkHeader.Value Then
        src.Append "<th"
    Else
        src.Append "<td"
    End If
    
    If cbohalign.Text <> "" Then
        src.Append " align=" & Chr$(34) & cbohalign.Text & Chr$(34)
    End If
    
    If cbovalign.Text <> "" Then
        src.Append " valign=" & Chr$(34) & cbovalign.Text & Chr$(34)
    End If
    
    If txtcolspa.Text <> "1" Then
        src.Append " colspan=" & Chr$(34) & txtcolspa.Text & Chr$(34)
    End If
    
    If txtrowspa.Text <> "1" Then
        src.Append " rowspan=" & Chr$(34) & txtrowspa.Text & Chr$(34)
    End If
    
    If txtHeight.Text <> "" Then
        If optHPix(1).Value Then
            src.Append " height=" & Chr$(34) & txtHeight.Text & "%" & Chr$(34)
        Else
            src.Append " height=" & Chr$(34) & txtHeight.Text & Chr$(34)
        End If
    End If
    
    If txtWidth.Text <> "" Then
        If optWPix(1).Value Then
            src.Append " width=" & Chr$(34) & txtWidth.Text & "%" & Chr$(34)
        Else
            src.Append " width=" & Chr$(34) & txtWidth.Text & Chr$(34)
        End If
    End If
    
    If txtBorColor(3).Text <> "" Then
        src.Append " bgcolor=" & Chr$(34) & txtBorColor(3).Text & Chr$(34)
    End If
        
    If txtBorColor(0).Text <> "" Then
        src.Append " bordercolor=" & Chr$(34) & txtBorColor(0).Text & Chr$(34)
    End If
    
    If txtBorColor(1).Text <> "" Then
        src.Append " bordercolorlight=" & Chr$(34) & txtBorColor(1).Text & Chr$(34)
    End If
        
    If txtBorColor(2).Text <> "" Then
        src.Append " bordercolordark=" & Chr$(34) & txtBorColor(2).Text & Chr$(34)
    End If
    
    If chknowra.Value Then
        src.Append " nowrap=" & Chr$(34) & "nowrap" & Chr$(34)
    End If
        
    If txtpic.Text <> "" Then
        src.Append " background=" & Chr$(34) & txtpic.Text & Chr$(34)
    End If
    
    If chkHeader.Value Then
        src.Append "></th>"
    Else
        src.Append "></td>"
    End If
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        frmMain.ActiveForm.Insertar src.ToString
    End If
    
    Set src = Nothing
    
End Sub

Private Sub ClrBorPicker1_ColorSelected(Index As Integer, m_Color As stdole.Ole_Color, m_Code As String)
    txtBorColor(Index).Text = m_Code
End Sub


Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call insertar_celda
        Unload Me
        Exit Sub
    ElseIf Index = 1 Then
        Unload Me
    Else
        Dim Archivo As String
        Dim glosa As String
        
        glosa = "CompuServe Graphics Interchange (*.gif)|*.gif|"
        glosa = glosa & "jpg (*.jpg)|*.jpg|"
        glosa = glosa & "ico (*.ico)|*.ico|"
        glosa = glosa & "All files (*.*)|*.*"
    
        If Cdlg.VBGetOpenFileName(Archivo, , , , , , glosa, , LastPath) Then
            txtpic.Text = Replace(Archivo, "\", "/")
        End If
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
    
    cbohalign.AddItem "left"
    cbohalign.AddItem "right"
    cbohalign.AddItem "center"
    cbohalign.AddItem "justify"
    
    cbovalign.AddItem "top"
    cbovalign.AddItem "middle"
    cbovalign.AddItem "bottom"
    cbovalign.AddItem "baseline"
    
    txtcolspa.Text = "1"
    txtrowspa.Text = "1"
    
    util.SetNumber txtcolspa.hwnd
    util.SetNumber txtrowspa.hwnd
    util.SetNumber txtWidth.hwnd
    util.SetNumber txtHeight.hwnd

    ClrBorPicker1(0).PathPaleta = App.Path & "\pal\256c.pal"
    ClrBorPicker1(1).PathPaleta = App.Path & "\pal\256c.pal"
    ClrBorPicker1(2).PathPaleta = App.Path & "\pal\256c.pal"
    ClrBorPicker1(3).PathPaleta = App.Path & "\pal\256c.pal"
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTableCell = Nothing
End Sub


Private Sub updcolspa_Change(ByVal lValue As Long)
    Dim valor
    If txtcolspa.Text = "" Then txtcolspa.Text = 0
    valor = txtcolspa.Text + lValue
    If CInt(valor) > 0 Then
        txtcolspa.Text = valor
    Else
        txtcolspa.Text = 1
    End If
End Sub

Private Sub updheight_Change(ByVal lValue As Long)
    Dim valor
    If txtHeight.Text = "" Then txtHeight.Text = 0
    valor = txtHeight.Text + lValue
    If CInt(valor) > 0 Then
        txtHeight.Text = valor
    Else
        txtHeight.Text = 1
    End If
End Sub

Private Sub updrowspa_Change(ByVal lValue As Long)
    Dim valor
    If txtrowspa.Text = "" Then txtrowspa.Text = 0
    valor = txtrowspa.Text + lValue
    If CInt(valor) > 0 Then
        txtrowspa.Text = valor
    Else
        txtrowspa.Text = 1
    End If
End Sub

Private Sub updwidth_Change(ByVal lValue As Long)
    Dim valor
    If txtWidth.Text = "" Then txtWidth.Text = 0
    valor = txtWidth.Text + lValue
    If CInt(valor) > 0 Then
        txtWidth.Text = valor
    Else
        txtWidth.Text = 1
    End If
End Sub


