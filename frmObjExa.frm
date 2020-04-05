VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmObjExa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Javascript Object Browser"
   ClientHeight    =   7605
   ClientLeft      =   375
   ClientTop       =   720
   ClientWidth     =   11010
   Icon            =   "frmObjExa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra 
      Caption         =   "Properties, methods and constants"
      Height          =   5790
      Index           =   1
      Left            =   2190
      TabIndex        =   5
      Top             =   75
      Width           =   8790
      Begin MSComctlLib.ListView lvwMembers 
         Height          =   5460
         Left            =   75
         TabIndex        =   6
         Top             =   240
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   9631
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgForm"
         SmallIcons      =   "imgForm"
         ColHdrIcons     =   "imgForm"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Members"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Help"
            Object.Width           =   8819
         EndProperty
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Object List"
      Height          =   5790
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   75
      Width           =   2055
      Begin MSComctlLib.ListView lvwObject 
         Height          =   5475
         Left            =   75
         TabIndex        =   4
         Top             =   240
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   9657
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "imgForm"
         SmallIcons      =   "imgForm"
         ColHdrIcons     =   "imgForm"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Class"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.TextBox txtHelp 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000080&
      Height          =   1365
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   6105
      Width           =   10815
   End
   Begin MSComctlLib.ImageList imgForm 
      Left            =   6030
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjExa.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjExa.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjExa.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjExa.frx":095A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjExa.frx":0C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjExa.frx":10C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjExa.frx":13E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjExa.frx":153A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjExa.frx":1694
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjExa.frx":17EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjExa.frx":1948
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbltype 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TYPE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   2190
      TabIndex        =   7
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label lblmember 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LBLMEMBER"
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
      Left            =   4200
      TabIndex        =   2
      Top             =   5880
      Width           =   1140
   End
   Begin VB.Label lblobj 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LBLOBJ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   5880
      Width           =   690
   End
End
Attribute VB_Name = "frmObjExa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cargamiembros(ByVal objeto As String)

    Dim k As Integer
    
    Dim ini As String
    'Dim num As String
    Dim tipo As String
    Dim miembro As String
    Dim glosa As String
    
    Dim sSections() As String
    
    ini = util.StripPath(App.Path) & "config\jshelp.ini"
        
    If Not ArchivoExiste2(ini) Then
        MsgBox "File not found : " & ini, vbCritical
        Exit Sub
    End If
    
    lvwMembers.ListItems.Clear
    
    get_info_section objeto, sSections, ini
    
    For k = 2 To UBound(sSections)
        glosa = sSections(k) 'Util.LeeIni(ini, Objeto, "ele" & k)
        If Len(glosa) > 0 Then
            miembro = util.Explode(glosa, 1, "#")
            tipo = util.Explode(glosa, 2, "#")
            glosa = util.Explode(glosa, 3, "#")
            
            If Len(miembro) > 0 Then
                If tipo = "1" Then
                    lvwMembers.ListItems.Add , "k" & k, miembro, 1, 1
                ElseIf tipo = "2" Then
                    lvwMembers.ListItems.Add , "k" & k, miembro, 2, 2
                ElseIf tipo = "3" Then
                    lvwMembers.ListItems.Add , "k" & k, miembro, 10, 10
                ElseIf tipo = "4" Then
                    lvwMembers.ListItems.Add , "k" & k, miembro, 11, 11
                ElseIf tipo = "5" Then  'coleccion
                    lvwMembers.ListItems.Add , "k" & k, miembro, 7, 7
                ElseIf tipo = "6" Then  'objeto
                    lvwMembers.ListItems.Add , "k" & k, miembro, 9, 9
                End If
                
                If Len(tipo) > 0 Then
                    If tipo = "1" Then
                        lvwMembers.ListItems("k" & k).SubItems(1) = "Property"
                    ElseIf tipo = "2" Then
                        lvwMembers.ListItems("k" & k).SubItems(1) = "Method"
                    ElseIf tipo = "3" Then
                        lvwMembers.ListItems("k" & k).SubItems(1) = "Event"
                    ElseIf tipo = "4" Then
                        lvwMembers.ListItems("k" & k).SubItems(1) = "Constant"
                    ElseIf tipo = "5" Then
                        lvwMembers.ListItems("k" & k).SubItems(1) = "Collection"
                    ElseIf tipo = "6" Then
                        lvwMembers.ListItems("k" & k).SubItems(1) = "Object"
                    End If
                    lvwMembers.ListItems("k" & k).SubItems(2) = glosa
                End If
            End If
        End If
    Next k
    
End Sub

Private Sub CargarObjetos()

    Dim k As Integer
    Dim ini As String
    'Dim num As String
    
    ini = util.StripPath(App.Path) & "config\jshelp.ini"
    
    If Not ArchivoExiste2(ini) Then
        MsgBox "File not found.", vbCritical
        Exit Sub
    End If
    
    For k = 1 To UBound(udtObjetos)
        lvwObject.ListItems.Add , "k" & k, udtObjetos(k), 5, 5
    Next k
    
    'num = Util.LeeIni(ini, "Objetos", "num")
    'For k = 1 To num
    '    glosa = Util.LeeIni(ini, "objetos", "ele" & k)
    'Next k
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Call util.Hourglass(hwnd, True)
    Call util.CenterForm(Me)
    Call CargarObjetos

    lblobj.Caption = vbNullString
    lbltype.Caption = vbNullString
    lblMember.Caption = vbNullString
    
    Debug.Print "load"
    
    Call util.Hourglass(hwnd, False)
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmObjExa = Nothing
End Sub


Private Sub lvwMembers_ItemClick(ByVal ITem As MSComctlLib.ListItem)

    lblMember.Caption = ITem.Text '& "-" & Item.SubItems(1)
    lbltype.Caption = ITem.SubItems(1)
    txtHelp.Text = ITem.SubItems(2)
    
End Sub


Private Sub lvwObject_ItemClick(ByVal ITem As MSComctlLib.ListItem)

    Call cargamiembros(ITem.Text)
    lblobj.Caption = ITem.Text
    If lvwMembers.ListItems.count > 0 Then
        lvwMembers_ItemClick lvwMembers.ListItems(1)
    End If
    
End Sub

