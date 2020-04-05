VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegExp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create a new regular expressions"
   ClientHeight    =   6150
   ClientLeft      =   3090
   ClientTop       =   2400
   ClientWidth     =   7545
   Icon            =   "frmRegExp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   6
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   3030
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   825
      Width           =   7380
      Begin MSComctlLib.ListView lvwregexp 
         Height          =   2655
         Left            =   105
         TabIndex        =   1
         Top             =   255
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Help"
            Object.Width           =   8819
         EndProperty
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Regular Expression"
      Height          =   675
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   105
      Width           =   7380
      Begin VB.TextBox txtVar 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Text            =   "pattern = /\s*;\s*/;"
         Top             =   255
         Width           =   7110
      End
   End
   Begin VB.TextBox txtHelp 
      ForeColor       =   &H000000C0&
      Height          =   1350
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4140
      Width           =   7380
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "More Help ...."
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   3
      Top             =   3915
      Width           =   960
   End
End
Attribute VB_Name = "frmRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub insertar_regexp()

    If Len(txtVar.Text) > 0 Then
        If frmMain.ActiveForm.Name = "frmEdit" Then
            Call frmMain.ActiveForm.Insertar(txtVar.Text)
        End If
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call insertar_regexp
        Unload Me
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
    Dim Num As Variant
    Dim num2 As Variant
    Dim Archivo As String
    Dim linea As String
    
    util.CenterForm Me
    
    Archivo = util.StripPath(App.Path) & "config\regexp.ini"
    
    Num = util.LeeIni(Archivo, "regular_expresion", "num")
    
    If Num = "" Then
        Exit Sub
    End If
    
    For k = 1 To Num
        linea = util.LeeIni(Archivo, "regexp" & k, "name")
        If Len(linea) > 0 Then
            lvwregexp.ListItems.Add , "k" & k, linea
            lvwregexp.ListItems("k" & k).Tag = k
            num2 = util.LeeIni(Archivo, "regexp" & k, "help")
            If Len(num2) > 0 Then
                linea = util.LeeIni(Archivo, "regexp" & k, "h1")
                If Len(linea) > 0 Then
                    lvwregexp.ListItems("k" & k).SubItems(1) = linea
                End If
            End If
        End If
    Next k
    
    flat_lviews Me
    'DrawXPCtl Me
    
    'Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    'cmd(0).Refresh
    'cmd(1).Refresh
    
    Debug.Print "load"
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmRegExp = Nothing
End Sub


Private Sub lvwregexp_ItemClick(ByVal ITem As MSComctlLib.ListItem)

    Dim linea As String
    Dim Num As Variant
    Dim k As Integer
    Dim src As New cStringBuilder
    Dim Archivo As String
    
    Archivo = util.StripPath(App.Path) & "config\regexp.ini"
    
    Num = util.LeeIni(Archivo, "regexp" & ITem.Tag, "help")
    txtHelp.Text = ""
    If Len(Num) > 0 Then
        For k = 1 To Num
            linea = util.LeeIni(Archivo, "regexp" & ITem.Tag, "h" & k)
            If Len(linea) > 0 Then
                src.Append linea & vbNewLine & vbNewLine
            End If
        Next k
        txtHelp.Text = src.ToString
    End If
    
    Set src = Nothing
    
End Sub


