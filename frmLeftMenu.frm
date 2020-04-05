VERSION 5.00
Object = "{246E535D-09D2-4109-80DA-2FF183F4D185}#2.1#0"; "colorpick.ocx"
Begin VB.Form frmLeftMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Left Menu Wizard"
   ClientHeight    =   5760
   ClientLeft      =   4275
   ClientTop       =   2175
   ClientWidth     =   7650
   Icon            =   "frmLeftMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   20
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   2
      Left            =   5760
      TabIndex        =   19
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Delete Menu Item"
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   18
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Add Menu Item"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   17
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame fra 
      Caption         =   "Menu Settings"
      Height          =   3945
      Index           =   1
      Left            =   30
      TabIndex        =   9
      Top             =   870
      Width           =   7590
      Begin VB.ListBox lstMenus 
         Height          =   2400
         Left            =   90
         TabIndex        =   2
         Top             =   945
         Width           =   2880
      End
      Begin VB.ListBox lstActions 
         Height          =   2400
         Left            =   3030
         TabIndex        =   3
         Top             =   960
         Width           =   4470
      End
      Begin VB.TextBox txtMenu 
         Height          =   285
         Left            =   90
         TabIndex        =   0
         Top             =   390
         Width           =   2910
      End
      Begin VB.TextBox txtAction 
         Height          =   285
         Left            =   3030
         TabIndex        =   1
         Top             =   390
         Width           =   4470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Menu Elements"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Menu Actions"
         Height          =   195
         Left            =   3030
         TabIndex        =   12
         Top             =   735
         Width           =   975
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Menu Caption"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   195
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Menu Action"
         Height          =   195
         Index           =   3
         Left            =   3030
         TabIndex        =   10
         Top             =   195
         Width           =   900
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Color Settings (Click to select a color)"
      Height          =   750
      Index           =   0
      Left            =   30
      TabIndex        =   4
      Top             =   105
      Width           =   7620
      Begin ColorPick.ClrPicker ClrPicker1 
         Height          =   285
         Left            =   2025
         TabIndex        =   14
         Top             =   285
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
      End
      Begin ColorPick.ClrPicker ClrPicker2 
         Height          =   285
         Left            =   5145
         TabIndex        =   15
         Top             =   285
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Menu Color"
         Height          =   195
         Index           =   1
         Left            =   1110
         TabIndex        =   8
         Top             =   330
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HighLited Color"
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   7
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label lblmenucolor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   1005
         Width           =   75
      End
      Begin VB.Label lblhighcolor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4095
         TabIndex        =   5
         Top             =   1005
         Width           =   75
      End
   End
   Begin VB.Image imgOP 
      Height          =   255
      Left            =   7185
      Top             =   5400
      Width           =   300
   End
   Begin VB.Image imgNE 
      Height          =   255
      Left            =   6855
      Top             =   5400
      Width           =   300
   End
   Begin VB.Image imgFX 
      Height          =   255
      Left            =   6540
      Top             =   5400
      Width           =   300
   End
   Begin VB.Image imgIE 
      Height          =   255
      Left            =   6225
      Top             =   5400
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
      Left            =   4620
      TabIndex        =   16
      Top             =   5415
      Width           =   1485
   End
End
Attribute VB_Name = "frmLeftMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function CrearMenuLeft() As Boolean

    Dim k As Integer
    Dim src As New cStringBuilder
    Dim menucolor As String
    Dim highcolor As String
    Dim menuname As String
    
    menucolor = lblmenucolor.Caption
    highcolor = lblhighcolor.Caption
    
    If lstMenus.ListCount - 1 > -1 Then
        src.Append "<HTML>" & vbNewLine
        src.Append "<HEAD>" & vbNewLine
        src.Append "<META http-equiv=" & Chr$(34) & "Content-Type" & Chr$(34) & "CONTENT=" & Chr$(34) & "text/html;" & Chr$(34) & ">" & vbNewLine
        src.Append "<TITLE>Testing Page</TITLE>" & vbNewLine
        src.Append "<BODY>" & vbNewLine
        src.Append "<script>" & vbNewLine
        src.Append "// function to evaluate menus" & vbNewLine
        src.Append "function color (id,color){" & vbNewLine
        src.Append "    el = document.getElementById(id).style;" & vbNewLine
        src.Append "    el.background = color;" & vbNewLine
        src.Append "    el.cursor = 'hand';" & vbNewLine
        src.Append "}" & vbNewLine
        src.Append "</script>" & vbNewLine
        src.Append "" & vbNewLine
        src.Append "<table width=" & Chr$(34) & "15%" & Chr$(34) & " border=" & Chr$(34) & "0" & Chr$(34) & _
                    " cellpadding=" & Chr$(34) & "0" & Chr$(34) & " cellspacing=" & Chr$(34) & "1" & Chr$(34) & _
                    " bgcolor=" & Chr$(34) & "#006699" & Chr$(34) & ">" & vbNewLine
        
        For k = 0 To lstMenus.ListCount - 1
            menuname = "menu" & k + 1
            src.Append "<tr>" & vbNewLine
            If lstActions.List(k) = "" Then
                src.Append "<td bgcolor=" & Chr$(34) & menucolor & Chr$(34) & " id=" & Chr$(34) & menuname & Chr$(34) & _
                       " onmouseover=" & Chr$(34) & "color('" & menuname & "','" & highcolor & "');" & Chr$(34) & _
                       " onMouseOut =" & Chr$(34) & "color('" & menuname & "','" & menucolor & "');" & Chr$(34) & ">" & _
                       "<font color=" & Chr$(34) & "#FFFFFF" & Chr$(34) & " size=" & Chr$(34) & "1" & Chr$(34) & _
                       " face=" & Chr$(34) & "Verdana, Arial, Helvetica, sans-serif" & Chr$(34) & ">" & lstMenus.List(k) & "</font></td>" & vbNewLine
            Else
                src.Append "<td bgcolor=" & Chr$(34) & menucolor & Chr$(34) & " id=" & Chr$(34) & menuname & Chr$(34) & _
                       " onmouseover=" & Chr$(34) & "color('" & menuname & "','" & highcolor & "');" & Chr$(34) & _
                       " onMouseOut =" & Chr$(34) & "color('" & menuname & "','" & menucolor & "');" & Chr$(34) & ">" & _
                       "<font color=" & Chr$(34) & "#FFFFFF" & Chr$(34) & " size=" & Chr$(34) & "1" & Chr$(34) & _
                       " face=" & Chr$(34) & "Verdana, Arial, Helvetica, sans-serif" & Chr$(34) & ">" & _
                       "<a href=" & Chr$(34) & lstActions.List(k) & Chr$(34) & ">" & lstMenus.List(k) & "</a>" & "</font></td>" & vbNewLine
            End If
            src.Append "</tr>" & vbNewLine
        Next k
        src.Append "</table>" & vbNewLine
        src.Append "</body>" & vbNewLine
        src.Append "</html>" & vbNewLine
        
        If frmMain.ActiveForm.Name = "frmEdit" Then
            Call frmMain.ActiveForm.Insertar(src.ToString)
        End If
        
        CrearMenuLeft = True
    Else
        MsgBox "Nothing to do", vbCritical
    End If
    
    Set src = Nothing
    
End Function

Private Sub ClrPicker1_ColorSelected(m_Color As stdole.Ole_Color, m_Code As String)
    lblmenucolor.Caption = m_Code
End Sub


Private Sub ClrPicker2_ColorSelected(m_Color As stdole.Ole_Color, m_Code As String)
    lblhighcolor.Caption = m_Code
End Sub


Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If txtMenu.Text <> "" Then
            lstMenus.AddItem txtMenu.Text
            lstActions.AddItem txtAction.Text
            txtMenu.SetFocus
        End If
    ElseIf Index = 1 Then
        If lstMenus.ListIndex <> -1 Then
            lstMenus.RemoveItem lstMenus.ListIndex
            lstActions.RemoveItem lstMenus.ListIndex
        End If
    ElseIf Index = 3 Then
        If CrearMenuLeft() Then
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
    
    ClrPicker1.PathPaleta = App.Path & "\pal\256c.pal"
    ClrPicker2.PathPaleta = App.Path & "\pal\256c.pal"
    
    Debug.Print "load"
        
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmLeftMenu = Nothing
End Sub


Private Sub txtMenu_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        cmd_Click 0
    End If
    
End Sub


