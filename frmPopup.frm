VERSION 5.00
Begin VB.Form frmPopup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Popup Window Wizard"
   ClientHeight    =   4485
   ClientLeft      =   3075
   ClientTop       =   1965
   ClientWidth     =   7155
   Icon            =   "frmPopup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   29
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   28
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtFuncion 
      Height          =   285
      Left            =   1365
      TabIndex        =   0
      Top             =   120
      Width           =   5715
   End
   Begin VB.Frame fra 
      Caption         =   "Window Position"
      Height          =   1050
      Index           =   2
      Left            =   75
      TabIndex        =   21
      Top             =   1515
      Width           =   7005
      Begin VB.TextBox txtWindowTop 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1335
         MaxLength       =   4
         TabIndex        =   4
         Top             =   615
         Width           =   1170
      End
      Begin VB.TextBox txtWindowWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4395
         MaxLength       =   4
         TabIndex        =   6
         Top             =   615
         Width           =   1170
      End
      Begin VB.TextBox txtWindowHeight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4395
         MaxLength       =   4
         TabIndex        =   5
         Top             =   255
         Width           =   1170
      End
      Begin VB.TextBox txtWindowLeft 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1335
         MaxLength       =   4
         TabIndex        =   3
         Top             =   255
         Width           =   1170
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top"
         Height          =   195
         Index           =   5
         Left            =   900
         TabIndex        =   25
         Top             =   690
         Width           =   285
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         Height          =   195
         Index           =   3
         Left            =   3720
         TabIndex        =   24
         Top             =   690
         Width           =   420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         Height          =   195
         Index           =   2
         Left            =   3720
         TabIndex        =   23
         Top             =   330
         Width           =   465
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left"
         Height          =   195
         Index           =   4
         Left            =   915
         TabIndex        =   22
         Top             =   330
         Width           =   270
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Window Settings"
      Height          =   990
      Index           =   1
      Left            =   75
      TabIndex        =   18
      Top             =   480
      Width           =   7005
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   6480
         TabIndex        =   30
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtLink 
         Height          =   285
         Left            =   1455
         TabIndex        =   2
         Top             =   570
         Width           =   4980
      End
      Begin VB.TextBox txtWindowName 
         Height          =   285
         Left            =   1455
         TabIndex        =   1
         Top             =   225
         Width           =   2790
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Link"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   20
         Top             =   570
         Width           =   300
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Window identifier"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   19
         Top             =   255
         Width           =   1215
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Window Options"
      Height          =   975
      Index           =   0
      Left            =   75
      TabIndex        =   17
      Top             =   2610
      Width           =   7020
      Begin VB.CheckBox chk 
         Caption         =   "Status Bar"
         Height          =   255
         Index           =   9
         Left            =   5535
         TabIndex        =   16
         Top             =   555
         Width           =   1140
      End
      Begin VB.CheckBox chk 
         Caption         =   "Title Bar"
         Height          =   255
         Index           =   8
         Left            =   4185
         TabIndex        =   14
         Top             =   555
         Width           =   1140
      End
      Begin VB.CheckBox chk 
         Caption         =   "FullScreen"
         Height          =   255
         Index           =   7
         Left            =   2820
         TabIndex        =   12
         Top             =   555
         Width           =   1140
      End
      Begin VB.CheckBox chk 
         Caption         =   "Dependent"
         Height          =   255
         Index           =   6
         Left            =   1440
         TabIndex        =   10
         Top             =   555
         Width           =   1140
      End
      Begin VB.CheckBox chk 
         Caption         =   "Directories"
         Height          =   255
         Index           =   5
         Left            =   195
         TabIndex        =   8
         Top             =   555
         Width           =   1140
      End
      Begin VB.CheckBox chk 
         Caption         =   "Resizable"
         Height          =   255
         Index           =   2
         Left            =   2820
         TabIndex        =   11
         Top             =   270
         Width           =   1080
      End
      Begin VB.CheckBox chk 
         Caption         =   "ScrollBars"
         Height          =   255
         Index           =   4
         Left            =   5535
         TabIndex        =   15
         Top             =   270
         Width           =   1095
      End
      Begin VB.CheckBox chk 
         Caption         =   "Location Bar"
         Height          =   240
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Top             =   270
         Width           =   1215
      End
      Begin VB.CheckBox chk 
         Caption         =   "Toolbar"
         Height          =   255
         Index           =   3
         Left            =   4185
         TabIndex        =   13
         Top             =   270
         Width           =   930
      End
      Begin VB.CheckBox chk 
         Caption         =   "Menu Bar"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   7
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.Image imgOP 
      Height          =   255
      Left            =   6795
      Top             =   4170
      Width           =   300
   End
   Begin VB.Image imgNE 
      Height          =   255
      Left            =   6465
      Top             =   4170
      Width           =   300
   End
   Begin VB.Image imgFX 
      Height          =   255
      Left            =   6150
      Top             =   4170
      Width           =   300
   End
   Begin VB.Image imgIE 
      Height          =   255
      Left            =   5835
      Top             =   4170
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
      Index           =   7
      Left            =   4230
      TabIndex        =   27
      Top             =   4185
      Width           =   1485
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Function Name"
      Height          =   195
      Index           =   6
      Left            =   75
      TabIndex        =   26
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function CreatePopup() As Boolean

    Dim src As New cStringBuilder
    Dim menubar As String
    Dim toolbar As String
    Dim location As String
    Dim scrollbar As String
    Dim resizable As String
    Dim directories As String
    Dim dependent As String
    Dim fullscreen As String
    Dim titlebar As String
    Dim statusbar As String
    'Dim Indice As Integer
    'Dim k As Integer
    
    If txtWindowName.Text = "" Then
        MsgBox "Must input the window name.", vbCritical
        txtWindowName.SetFocus
        Exit Function
    End If
    
    If txtLink.Text = "" Then
        MsgBox "Must input the link.", vbCritical
        txtLink.SetFocus
        Exit Function
    End If
    
    If txtFuncion.Text = "" Then
        MsgBox "Must input the function name.", vbCritical
        txtFuncion.SetFocus
        Exit Function
    ElseIf InStr(txtFuncion.Text, "()") Then
        MsgBox "Only input the function name.", vbCritical
        txtFuncion.SetFocus
        Exit Function
    End If
    
    If txtWindowLeft.Text = "" Then
        MsgBox "Must input the window left.", vbCritical
        txtWindowLeft.SetFocus
        Exit Function
    End If
    
    If txtWindowTop.Text = "" Then
        MsgBox "Must input the window top.", vbCritical
        txtWindowTop.SetFocus
        Exit Function
    End If
    
    If txtWindowHeight.Text = "" Then
        MsgBox "Must input the window height.", vbCritical
        txtWindowHeight.SetFocus
        Exit Function
    End If
    
    If txtWindowWidth.Text = "" Then
        MsgBox "Must input the window width.", vbCritical
        txtWindowWidth.SetFocus
        Exit Function
    End If
        
    menubar = "0"
    If chk(0).Value = 1 Then menubar = "1"
    
    toolbar = "0"
    If chk(3).Value = 1 Then toolbar = "1"
    
    location = "0"
    If chk(1).Value = 1 Then location = "1"
    
    resizable = "0"
    If chk(2).Value = 1 Then resizable = "1"
    
    scrollbar = "0"
    If chk(4).Value = 1 Then scrollbar = "1"
    
    directories = "0"
    If chk(5).Value = 1 Then directories = "1"
    
    dependent = "0"
    If chk(6).Value = 1 Then dependent = "1"
    
    fullscreen = "0"
    If chk(7).Value = 1 Then fullscreen = "1"
    
    titlebar = "0"
    If chk(8).Value = 1 Then titlebar = "1"
    
    statusbar = "0"
    If chk(9).Value = 1 Then statusbar = "1"
    
    src.Append "function " & txtFuncion.Text & "()" & vbNewLine
    src.Append "{" & vbNewLine
    src.Append "var mywindow=window.open(" & Chr$(34) & txtLink.Text & Chr$(34) & "," & _
               Chr$(34) & txtWindowName.Text & Chr$(34) & "," & Chr$(34) & _
               "left=" & txtWindowLeft.Text & ",top=" & txtWindowTop.Text & "," & _
               "height=" & txtWindowHeight.Text & "," & "width=" & txtWindowWidth.Text & "," & _
               "menubar=" & menubar & ",toolbar=" & toolbar & ",location=" & location & "," & _
               "directories=" & directories & "," & "dependent=" & dependent & "," & _
               "fullscreen=" & fullscreen & "," & "titlebar=" & titlebar & ",status=" & statusbar & _
               ",scrollbar=" & scrollbar & ",resizable=" & resizable & Chr$(34) & ")" & vbNewLine
    src.Append "}" & vbNewLine
    'src.Append "</script>" & vbNewLine
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(src.ToString)
    End If
    
    Set src = Nothing
    
    CreatePopup = True
    
End Function

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If CreatePopup() Then
            Unload Me
        End If
    ElseIf Index = 1 Then
        Unload Me
    ElseIf Index = 2 Then
        Dim Archivo As String
        Dim glosa As String
        glosa = strGlosa()
        If Cdlg.VBGetOpenFileName(Archivo, , , , , , glosa, , , "Select a file ...", "js") Then
            txtLink.Text = Archivo
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
    
    Set imgIE.Picture = LoadResPicture(1007, vbResBitmap)
    Set imgFX.Picture = LoadResPicture(1008, vbResBitmap)
    Set imgNE.Picture = LoadResPicture(1009, vbResBitmap)
    Set imgOP.Picture = LoadResPicture(1010, vbResBitmap)
        
    util.SetNumber (txtWindowLeft.hwnd)
    util.SetNumber (txtWindowTop.hwnd)
    util.SetNumber (txtWindowHeight.hwnd)
    util.SetNumber (txtWindowWidth.hwnd)
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload : " & Me.Name
    Set frmPopup = Nothing
End Sub


