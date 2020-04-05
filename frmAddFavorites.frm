VERSION 5.00
Begin VB.Form frmAddFavorites 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add to Favorites Wizard"
   ClientHeight    =   3780
   ClientLeft      =   3840
   ClientTop       =   2880
   ClientWidth     =   5955
   Icon            =   "frmAddFavorites.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   12
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Output"
      Height          =   660
      Index           =   1
      Left            =   60
      TabIndex        =   9
      Top             =   2160
      Width           =   5820
      Begin VB.CheckBox chk 
         Caption         =   "Mozilla Firefox"
         Height          =   225
         Index           =   1
         Left            =   3135
         TabIndex        =   11
         Top             =   300
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.CheckBox chk 
         Caption         =   "Internet Explorer"
         Height          =   225
         Index           =   0
         Left            =   1125
         TabIndex        =   10
         Top             =   300
         Value           =   1  'Checked
         Width           =   1515
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      ForeColor       =   &H00000000&
      Height          =   2055
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   5820
      Begin VB.TextBox txtTexto 
         Height          =   285
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1605
         Width           =   5655
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1050
         Width           =   5655
      End
      Begin VB.TextBox txtUrl 
         Height          =   285
         Left            =   60
         TabIndex        =   0
         Top             =   465
         Width           =   5655
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   7
         Top             =   1395
         Width           =   540
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Url"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   195
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   4
         Top             =   840
         Width           =   300
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   6
      Top             =   0
      Width           =   0
   End
   Begin VB.Image imgFX 
      Height          =   255
      Left            =   5595
      Top             =   3420
      Width           =   300
   End
   Begin VB.Image imgIE 
      Height          =   255
      Left            =   5265
      Top             =   3420
      Width           =   300
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Browser Compatibility"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   3705
      TabIndex        =   8
      Top             =   3435
      Width           =   1485
   End
End
Attribute VB_Name = "frmAddFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Function AddFavoritos() As Boolean

    Dim src As New cStringBuilder
    
    If txtUrl.Text = "" Then
        MsgBox "You must input the link to add.", vbCritical
        txtUrl.SetFocus
        Exit Function
    End If
    
    If txtText.Text = "" Then
        MsgBox "You must input the favourites text.", vbCritical
        txtUrl.SetFocus
        Exit Function
    End If
    
    If txtTexto.Text = "" Then
        MsgBox "You must input the link text.", vbCritical
        txtTexto.SetFocus
        Exit Function
    End If
    
    src.Append "<html>" & vbNewLine
    src.Append "<head>" & vbNewLine
    src.Append "<meta http-equiv=" & Chr$(34) & "Content-Type" & Chr$(34) & "CONTENT=" & Chr$(34) & "text/html;" & Chr$(34) & ">" & vbNewLine
    src.Append "<title>Testing Page</title>" & vbNewLine
    src.Append "</head>" & vbNewLine
    src.Append "<body>" & vbNewLine
    src.Append "<script language=" & Chr$(34) & "Javascript" & Chr$(34) & " type=" & Chr$(34) & "text/javascript" & Chr$(34) & ">" & vbNewLine
    src.Append "function savelink()" & vbNewLine
    src.Append "{" & vbNewLine
    src.Append "    var url=" & Chr$(34) & txtUrl.Text & Chr$(34) & ";" & vbNewLine
    src.Append "    var title=" & Chr$(34) & txtText.Text & Chr$(34) & ";" & vbNewLine
    src.Append "" & vbNewLine
    src.Append "    // IE cannot bookmark pages saved on hd: use main URL" & vbNewLine
    src.Append "    if(url.indexOf(" & Chr$(34) & "file:" & Chr$(34) & ") > -1)" & vbNewLine
    src.Append "        url=" & Chr$(34) & txtUrl.Text & Chr$(34) & ";" & vbNewLine
    src.Append "" & vbNewLine
    
    If chk(0).Value And chk(1).Value Then
        src.Append "    // add IE favorite" & vbNewLine
        src.Append "    if(window.external) {" & vbNewLine
        src.Append "        external.AddFavorite(url,title);" & vbNewLine
        src.Append "}   else if(window.sidebar && sidebar.addPanel) { // add to FF bookmarks" & vbNewLine
        src.Append "        sidebar.addPanel(title,url,'');" & vbNewLine
        src.Append "}   else {                                        // unknown browser: report user" & vbNewLine
        src.Append "        alert('Failed to recognize your browser, please bookmark the page manually.');" & vbNewLine
        src.Append "    }" & vbNewLine
    ElseIf chk(0).Value Then
        src.Append "    // add IE favorite" & vbNewLine
        src.Append "    if (window.external) {" & vbNewLine
        src.Append "        external.AddFavorite(url,title);" & vbNewLine
        src.Append "    }" & vbNewLine
    Else
        src.Append "    if(window.sidebar && sidebar.addPanel) { // add to FF bookmarks" & vbNewLine
        src.Append "        sidebar.addPanel(title,url,'');" & vbNewLine
        src.Append "    }" & vbNewLine
    End If
    src.Append "}" & vbNewLine
    src.Append "</script>" & vbNewLine
    src.Append "<a href=" & Chr$(34) & "javascript:savelink();" & Chr$(34) & "/>" & txtTexto.Text & "<a/>" & vbNewLine
    src.Append "</body>" & vbNewLine
    src.Append "</html>" & vbNewLine
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(src.ToString)
    End If
    
    Call util.GrabaIni(IniPath, "addtofavorites", "link", txtUrl.Text)
    Call util.GrabaIni(IniPath, "addtofavorites", "text", txtText.Text)
    Call util.GrabaIni(IniPath, "addtofavorites", "caption", txtTexto.Text)
    
    Set src = Nothing
    
    AddFavoritos = True
    
End Function

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If AddFavoritos() Then
            Unload Me
        End If
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

    util.CenterForm Me
    
    txtUrl.Text = util.LeeIni(IniPath, "addtofavorites", "link")
    txtText.Text = util.LeeIni(IniPath, "addtofavorites", "text")
    txtTexto.Text = util.LeeIni(IniPath, "addtofavorites", "caption")
    
    Set imgIE.Picture = LoadResPicture(1007, vbResBitmap)
    Set imgFX.Picture = LoadResPicture(1008, vbResBitmap)
    
    Debug.Print "load : " & Me.Name
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call clear_memory(Me)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmAddFavorites = Nothing
    Debug.Print "unload : " & Me.Name
End Sub


