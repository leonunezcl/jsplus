VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FCFAF346-DE8A-4FB6-8612-5000548EFDC7}#2.0#0"; "vbsListView6.ocx"
Begin VB.Form frmAddFunc 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Function"
   ClientHeight    =   6870
   ClientLeft      =   2670
   ClientTop       =   1740
   ClientWidth     =   6990
   ControlBox      =   0   'False
   Icon            =   "frmAgregarFuncion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   285
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7560
      Visible         =   0   'False
      Width           =   2250
   End
   Begin MSComctlLib.ImageList img 
      Left            =   6165
      Top             =   6405
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
            Picture         =   "frmAgregarFuncion.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Parameters"
      ForeColor       =   &H00FF8080&
      Height          =   2835
      Index           =   1
      Left            =   75
      TabIndex        =   17
      Top             =   1215
      Width           =   6825
      Begin VB.TextBox txtParDes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   975
         MaxLength       =   60
         TabIndex        =   3
         Top             =   615
         Width           =   5730
      End
      Begin VB.TextBox txtNomPar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   975
         MaxLength       =   255
         TabIndex        =   2
         Top             =   255
         Width           =   5730
      End
      Begin jsplus.MyButton cmd 
         Height          =   390
         Index           =   2
         Left            =   1275
         TabIndex        =   5
         Top             =   2355
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   688
         SPN             =   "MyButtonDefSkin"
         Text            =   "&Add"
         AccessKey       =   "A"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin jsplus.MyButton cmd 
         Height          =   390
         Index           =   3
         Left            =   3780
         TabIndex        =   6
         Top             =   2355
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   688
         SPN             =   "MyButtonDefSkin"
         Text            =   "&Delete"
         AccessKey       =   "D"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vbalListViewLib6.vbalListViewCtl lvw 
         Height          =   1260
         Left            =   120
         TabIndex        =   4
         Top             =   990
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   2223
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
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comment"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   5
         Left            =   105
         TabIndex        =   19
         Top             =   645
         Width           =   660
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   18
         Top             =   300
         Width           =   420
      End
   End
   Begin VB.Frame fr 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Other Info"
      ForeColor       =   &H00FF8080&
      Height          =   2220
      Left            =   60
      TabIndex        =   13
      Top             =   4080
      Width           =   6840
      Begin VB.TextBox txtDescrip 
         Appearance      =   0  'Flat
         Height          =   1290
         Left            =   90
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   825
         Width           =   6645
      End
      Begin VB.TextBox txtAutor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   735
         TabIndex        =   7
         Top             =   240
         Width           =   5970
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   615
         Width           =   795
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Autor"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   14
         Top             =   285
         Width           =   375
      End
   End
   Begin VB.Frame fra 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name"
      ForeColor       =   &H00FF8080&
      Height          =   1125
      Index           =   0
      Left            =   60
      TabIndex        =   11
      Top             =   75
      Width           =   6840
      Begin VB.TextBox txtRetorno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Text            =   "true"
         Top             =   630
         Width           =   5775
      End
      Begin VB.TextBox txtNomFun 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   270
         Width           =   5760
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Returns"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   16
         Top             =   675
         Width           =   555
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   12
         Top             =   330
         Width           =   420
      End
   End
   Begin jsplus.MyButton cmd 
      Height          =   390
      Index           =   0
      Left            =   1320
      TabIndex        =   9
      Top             =   6360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
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
   End
   Begin jsplus.MyButton cmd 
      Height          =   390
      Index           =   1
      Left            =   3840
      TabIndex        =   10
      Top             =   6360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
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
   End
   Begin vbalListViewLib6.vbalListViewCtl lvwtmp 
      Height          =   1620
      Left            =   4605
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7680
      Visible         =   0   'False
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   2858
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
End
Attribute VB_Name = "frmAddFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Indice As Integer
Public Nodo As String
Private Nombre As String
Private Autor As String
Private Parametros As String
Private ParDes As String
Private Comentario As String
Private Retorno As String
Private Sub AgregarParametro()

    Dim k As Integer
    
    If Len(txtNomPar.Text) = 0 Then
        MsgBox "Parameter name please.", vbCritical
        txtNomPar.SetFocus
        Exit Sub
    End If
    
    For k = 1 To lvw.ListItems.count
        If LCase$(lvw.ListItems(k).SubItems(1).Caption) = LCase$(txtNomPar.Text) Then
            MsgBox "Parameter already exists.", vbCritical
            txtNomPar.SetFocus
            Exit Sub
        End If
    Next k
    
    k = lvw.ListItems.count + 1
    
    Call lvw.ListItems.Add(, "k" & k, CStr(k), 1, 1)
    lvw.ListItems(k).SubItems(1).Caption = txtNomPar.Text
    lvw.ListItems(k).SubItems(2).Caption = txtParDes.Text
    
    txtNomPar.SetFocus
    
End Sub

Private Sub EliminarParametro()

    Dim k As Integer
    Dim J As Integer
    Dim param As String
    
    If lvw.ListItems.count > 0 Then
        If Not lvw.SelectedItem Is Nothing Then
            lvwtmp.ListItems.Clear
            J = 1
            
            param = lvw.SelectedItem.SubItems(1).Caption
            For k = 1 To lvw.ListItems.count
                If lvw.ListItems(k).SubItems(1).Caption <> param Then
                    lvwtmp.ListItems.Add , "k" & J, CStr(J)
                    lvwtmp.ListItems(J).SubItems(1).Caption = lvw.ListItems(k).SubItems(1).Caption
                    lvwtmp.ListItems(J).SubItems(2).Caption = lvw.ListItems(k).SubItems(2).Caption
                    J = J + 1
                End If
            Next k
            
            lvw.ListItems.Clear
            
            For k = 1 To lvwtmp.ListItems.count
                lvw.ListItems.Add , "k" & k, CStr(k), 1, 1
                lvw.ListItems(k).SubItems(1).Caption = lvwtmp.ListItems(k).SubItems(1).Caption
                lvw.ListItems(k).SubItems(2).Caption = lvwtmp.ListItems(k).SubItems(2).Caption
            Next k
        End If
    End If

End Sub

Private Function Validar() As Boolean
    
    Dim k As Integer
    
    If Len(txtNomFun.Text) = 0 Then
        MsgBox "Function name please.", vbCritical
        txtNomFun.SetFocus
        Exit Function
    End If

    If Not util.ValidPattern(txtNomFun.Text) Then
        MsgBox "Function name is not valid.", vbCritical
        txtNomFun.SetFocus
        Exit Function
    End If
    
    Nombre = Trim$(txtNomFun.Text)
    Autor = Trim$(txtAutor.Text)
    Comentario = Trim$(txtDescrip.Text)
    Parametros = ""
    Retorno = Trim$(txtRetorno.Text)
    
    If lvw.ListItems.count > 0 Then
        For k = 1 To lvw.ListItems.count
            Parametros = Parametros & lvw.ListItems(k).SubItems(1).Caption & ","
            If Len(lvw.ListItems(k).SubItems(1).Caption) > 0 Then
                ParDes = ParDes & lvw.ListItems(k).SubItems(2).Caption & vbNewLine
            End If
        Next k
        Parametros = Left$(Parametros, Len(Parametros) - 1)
        
        If Len(ParDes) > 0 Then
            ParDes = Left$(ParDes, Len(ParDes) - 1)
        End If
    End If
    
    Validar = True
    
End Function


Private Sub cmd_Click(Index As Integer)
    
    If Index = 0 Then       'validar
        If Validar() Then
            If frmMain.ActiveForm.Name = "frmEdit" Then
                Call frmMain.ActiveForm.AgregarFuncion(Nombre, Parametros, ParDes, Autor, Comentario, Retorno)
            End If
            Unload Me
        End If
    ElseIf Index = 2 Then   'agregar parametro
        Call AgregarParametro
    ElseIf Index = 3 Then   'eliminar parametro
        Call EliminarParametro
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
    
    Call util.CenterForm(Me)
    set_color_form Me
    'SetLayered hwnd, True
    
    With lvw
        .Columns.Add , "k1", "Nº", , 500
        .Columns.Add , "k2", "Name", , 2000
        .Columns.Add , "k3", "Comments", , 4000
    End With
    
    With lvwtmp
        .Columns.Add , "k1", "Nº", , 500
        .Columns.Add , "k2", "Name", , 2000
        .Columns.Add , "k3", "Comments", , 4000
    End With
    
    flat_lviews Me
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    cmd(0).Refresh
    cmd(1).Refresh
    cmd(2).Refresh
    cmd(3).Refresh
    Debug.Print "load"
    DrawXPCtl Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmAddFunc = Nothing
End Sub


Private Sub txtNomPar_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Len(Trim$(txtNomFun.Text)) > 0 Then
            Call AgregarParametro
        End If
    End If
    
End Sub

