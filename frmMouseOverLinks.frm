VERSION 5.00
Object = "{246E535D-09D2-4109-80DA-2FF183F4D185}#2.1#0"; "colorpick.ocx"
Begin VB.Form frmMouseOverLinks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MouseOver Wizard"
   ClientHeight    =   3015
   ClientLeft      =   4200
   ClientTop       =   3165
   ClientWidth     =   4065
   Icon            =   "frmMouseOverLinks.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   3
      Index           =   1
      Left            =   2400
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Ok"
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   2040
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   75
      Width           =   3900
      Begin VB.ComboBox cbohoverlink 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1755
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1320
         Width           =   2070
      End
      Begin VB.TextBox txthoverlink 
         Height          =   285
         Left            =   1755
         MaxLength       =   7
         TabIndex        =   2
         Top             =   960
         Width           =   1005
      End
      Begin VB.ComboBox cbotextlink 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1755
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   585
         Width           =   2070
      End
      Begin VB.TextBox txtColorText 
         Height          =   285
         Left            =   1755
         MaxLength       =   7
         TabIndex        =   0
         Top             =   225
         Width           =   1005
      End
      Begin ColorPick.ClrPicker ClrPicker1 
         Height          =   300
         Index           =   0
         Left            =   2775
         TabIndex        =   9
         Top             =   210
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         Color           =   255
         Code            =   "255"
      End
      Begin ColorPick.ClrPicker ClrPicker1 
         Height          =   300
         Index           =   1
         Left            =   2775
         TabIndex        =   10
         Top             =   945
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hover link decoration"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   8
         Top             =   1305
         Width           =   1515
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colour of hover link"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   7
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text link decoration"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colour of text link"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Image imgOP 
      Height          =   255
      Left            =   3705
      Top             =   2685
      Width           =   300
   End
   Begin VB.Image imgNE 
      Height          =   255
      Left            =   3375
      Top             =   2685
      Width           =   300
   End
   Begin VB.Image imgFX 
      Height          =   255
      Left            =   3060
      Top             =   2685
      Width           =   300
   End
   Begin VB.Image imgIE 
      Height          =   255
      Left            =   2745
      Top             =   2685
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
      Index           =   4
      Left            =   1140
      TabIndex        =   11
      Top             =   2700
      Width           =   1485
   End
End
Attribute VB_Name = "frmMouseOverLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function mouseoverlink() As Boolean

    Dim src As New cStringBuilder
    
    src.Append "<HTML>" & vbNewLine
    src.Append "<HEAD>" & vbNewLine
    src.Append "<META http-equiv=" & Chr$(34) & "Content-Type" & Chr$(34) & "CONTENT=" & Chr$(34) & "text/html;" & Chr$(34) & ">" & vbNewLine
    src.Append "<TITLE>Testing Page</TITLE>" & vbNewLine
    src.Append "<style>" & vbNewLine
    src.Append "<!--" & vbNewLine
    src.Append "a {text-decoration: " & cbotextlink.Text & "; color: " & txtColorText.Text & ";}" & vbNewLine
    src.Append "a:hover {text-decoration: " & cbohoverlink.Text & "; color: " & txthoverlink.Text & ";}" & vbNewLine
    src.Append "// -->" & vbNewLine
    src.Append "</style>" & vbNewLine
    src.Append "<BODY>" & vbNewLine
    src.Append "<a href=" & Chr$(34) & "http://www.link1.com" & Chr$(34) & ">Testing Link site 1</a>" & vbNewLine
    src.Append "<a href=" & Chr$(34) & "http://www.link2.com" & Chr$(34) & ">Testing Link site 2</a>" & vbNewLine
    src.Append "<a href=" & Chr$(34) & "http://www.link3.com" & Chr$(34) & ">Testing Link site 3</a>" & vbNewLine
    src.Append "<a href=" & Chr$(34) & "http://www.link4.com" & Chr$(34) & ">Testing Link site 4</a>" & vbNewLine
    src.Append "<a href=" & Chr$(34) & "http://www.link5.com" & Chr$(34) & ">Testing Link site 5</a>" & vbNewLine
    src.Append "</BODY>" & vbNewLine
    src.Append "</HTML>"
    
    Call util.GrabaIni(IniPath, "mouseover", "colour_text_link", txtColorText.Text)
    Call util.GrabaIni(IniPath, "mouseover", "text_link_deco", cbotextlink.ListIndex)
    Call util.GrabaIni(IniPath, "mouseover", "colour_hover_link", txthoverlink.Text)
    Call util.GrabaIni(IniPath, "mouseover", "hover_link_deco", cbohoverlink.ListIndex)
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(src.ToString)
    End If
    
    Set src = Nothing
    
    mouseoverlink = True
    
End Function

Private Sub ClrPicker1_ColorSelected(Index As Integer, m_Color As stdole.OLE_COLOR, m_Code As String)

    If Index = 0 Then
        txtColorText.Text = m_Code
    Else
        txthoverlink.Text = m_Code
    End If
    
End Sub


Private Sub cmd_Click(Index As Integer)

    If Index = 2 Then   'ACEPTAR
        If mouseoverlink() Then
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
        
    Set imgIE.Picture = LoadResPicture(1007, vbResBitmap)
    Set imgFX.Picture = LoadResPicture(1008, vbResBitmap)
    Set imgNE.Picture = LoadResPicture(1009, vbResBitmap)
    Set imgOP.Picture = LoadResPicture(1010, vbResBitmap)
    
    cbotextlink.AddItem "none"
    cbotextlink.AddItem "underline"
    cbotextlink.AddItem "overline"
    cbotextlink.AddItem "line through"
    
    cbohoverlink.AddItem "none"
    cbohoverlink.AddItem "underline"
    cbohoverlink.AddItem "overline"
    cbohoverlink.AddItem "line through"
    
    Dim valor
    
    txtColorText.Text = util.LeeIni(IniPath, "mouseover", "colour_text_link")
    valor = util.LeeIni(IniPath, "mouseover", "text_link_deco")
    If Len(valor) > 0 Then cbotextlink.ListIndex = valor
    txthoverlink.Text = util.LeeIni(IniPath, "mouseover", "colour_hover_link")
    valor = util.LeeIni(IniPath, "mouseover", "hover_link_deco")
    If Len(valor) > 0 Then cbohoverlink.ListIndex = valor
        
    ClrPicker1(0).PathPaleta = App.Path & "\pal\256c.pal"
    ClrPicker1(1).PathPaleta = App.Path & "\pal\256c.pal"
    
    Debug.Print "load"
    
    'DrawXPCtl Me
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call clear_memory(Me)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmMouseOverLinks = Nothing
End Sub


