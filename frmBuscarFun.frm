VERSION 5.00
Object = "{FCFAF346-DE8A-4FB6-8612-5000548EFDC7}#2.0#0"; "vbsListView6.ocx"
Begin VB.Form frmBuscarFun 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Wizard"
   ClientHeight    =   4935
   ClientLeft      =   2745
   ClientTop       =   1455
   ClientWidth     =   8280
   ControlBox      =   0   'False
   Icon            =   "frmBuscarFun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Begin vbalListViewLib6.vbalListViewCtl lvwfun 
      Height          =   3060
      Left            =   60
      TabIndex        =   2
      Top             =   1275
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   5398
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
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
      AutoArrange     =   0   'False
      Appearance      =   0
      FlatScrollBar   =   -1  'True
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   6450
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Frame fra 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Function ..."
      ForeColor       =   &H00FF8080&
      Height          =   1020
      Left            =   60
      TabIndex        =   6
      Top             =   30
      Width           =   8130
      Begin VB.ComboBox cboArchivo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   6165
      End
      Begin VB.TextBox txtFun 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1890
         TabIndex        =   0
         Top             =   255
         Width           =   6165
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Select File:"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   600
         Width           =   780
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   255
         Width           =   510
      End
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   0
      Left            =   1590
      TabIndex        =   3
      Top             =   4440
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Ok"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePos      =   3
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   1
      Left            =   3495
      TabIndex        =   4
      Top             =   4440
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Cancel"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePos      =   3
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   2
      Left            =   5415
      TabIndex        =   5
      Top             =   4440
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Select"
      AccessKey       =   "S"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePos      =   3
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   9
      Top             =   1065
      Width           =   555
   End
End
Attribute VB_Name = "frmBuscarFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Buscar()

    Dim k As Integer
    Dim j As Integer
    Dim i As Integer
    Dim Todos As Boolean
    Dim funcion As String
    Dim Archivo As String
    'Dim Buffer As String
    
    Archivo = ""
    funcion = txtFun.Text
    
    Todos = False
    If cboArchivo.ListIndex = 0 Then
        Todos = True
    Else
        Archivo = cboArchivo.Text
    End If
    
    lvwfun.ListItems.Clear
    
    i = 1
    
    Dim frm As Form
    For Each frm In Forms
        If TypeName(frm) = "frmEdit" Then
            If Not Todos Then
                If frm.Caption = Archivo Then
                    With frm
                        For j = 0 To .txtCode.LineCount
                            If InStr(.txtCode.GetLine(j), funcion) > 0 Then
                                lvwfun.ListItems.Add , "k" & i, CStr(i)
                                lvwfun.ListItems(i).SubItems(1).Caption = frm.Caption
                                lvwfun.ListItems(i).SubItems(2).Caption = j
                                lvwfun.ListItems(i).SubItems(3).Caption = Util.SacarBasura(.txtCode(k).GetLine(j))
                                i = i + 1
                            End If
                        Next j
                    End With
                End If
            Else
                With frm
                    For j = 0 To .txtCode.LineCount
                        If InStr(.txtCode.GetLine(j), funcion) > 0 Then
                            lvwfun.ListItems.Add , "k" & i, CStr(i)
                            lvwfun.ListItems(i).SubItems(1).Caption = frm.Caption
                            lvwfun.ListItems(i).SubItems(2).Caption = j
                            lvwfun.ListItems(i).SubItems(3).Caption = Util.SacarBasura(.txtCode.GetLine(j))
                            i = i + 1
                        End If
                    Next j
                End With
            End If
        End If
    Next
    
End Sub
Private Sub Seleccionar()

    Dim itmx As cListItem
    Dim k As Integer
    
    If lvwfun.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    If lvwfun.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    Set itmx = lvwfun.SelectedItem
    
    Dim frm As Form
    
    For Each frm In Forms
        If TypeName(frm) = "frmEdit" Then
            If frm.Caption = itmx.SubItems(1).Caption Then
                With frmMain
                    '.SelectTabMdi frm.hWnd
                    Call frm.txtCode.SelectLine(itmx.SubItems(2).Caption, True)
                    Call frm.txtCode.SetCaretPos(itmx.SubItems(2).Caption, 0)
                    frm.txtCode.HighlightedLine = itmx.SubItems(2).Caption
                    '.m_cMDITabs.ForceRefresh
                End With
                Exit For
            End If
        End If
    Next
    
    Set itmx = Nothing
    
End Sub

Private Function Validar() As Boolean

    If txtFun.Text = "" Then
        MsgBox "Funtion name.", vbCritical
        txtFun.SetFocus
        Exit Function
    End If

    If cboArchivo.ListIndex = -1 Then
        MsgBox "Select a file to search.", vbCritical
        cboArchivo.SetFocus
        Exit Function
    End If
    
    Validar = True
    
End Function
Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If Validar() Then
            Call Buscar
        End If
    ElseIf Index = 2 Then
        Call Seleccionar
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Dim k As Integer
    
    Util.CenterForm Me
    flat_lviews Me
    set_color_form Me
    cboArchivo.AddItem "All"
    
    Dim frm As Form
    
    For Each frm In Forms
        If TypeName(frm) = "frmEdit" Then
            cboArchivo.AddItem frm.Caption
        End If
    Next
                
    With lvwfun
        .Columns.Add , "k1", "Nº", , 800
        .Columns.Add , "k2", "File", , 1500
        .Columns.Add , "k3", "Line", , 700
        .Columns.Add , "k4", "Code", , 5000
    End With
    
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    cmd(0).Refresh
    cmd(1).Refresh
    cmd(2).Refresh
    
    DrawXPCtl Me
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmBuscarFun = Nothing
End Sub


Private Sub lvwfun_DblClick()

    Call Seleccionar
    
End Sub


Private Sub lvwfun_ItemClick(Item As vbalListViewLib6.cListItem)
    
    If Not Item Is Nothing Then
        Call Seleccionar
    End If
    
End Sub


