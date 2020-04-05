VERSION 5.00
Begin VB.Form frmRollover 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Image Rollover Wizard"
   ClientHeight    =   3375
   ClientLeft      =   3195
   ClientTop       =   2400
   ClientWidth     =   6720
   Icon            =   "frmRollover.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   448
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   15
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   2415
      Index           =   0
      Left            =   105
      TabIndex        =   6
      Top             =   105
      Width           =   6555
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   255
         Index           =   3
         Left            =   6000
         TabIndex        =   17
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   6000
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtImgOn 
         Height          =   285
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   4800
      End
      Begin VB.TextBox txtImgOff 
         Height          =   285
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   915
         Width           =   4800
      End
      Begin VB.ComboBox cboFun 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1635
         Width           =   4800
      End
      Begin VB.TextBox txtOnName 
         Height          =   285
         Left            =   1125
         TabIndex        =   1
         Top             =   570
         Width           =   4800
      End
      Begin VB.TextBox txtOffName 
         Height          =   285
         Left            =   1125
         TabIndex        =   3
         Top             =   1275
         Width           =   4800
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1125
         TabIndex        =   5
         Top             =   1995
         Width           =   4800
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Image On"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   12
         Top             =   270
         Width           =   690
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Image Off"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   11
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Function"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   10
         Top             =   1665
         Width           =   615
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Image Name"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   9
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Image Name"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   8
         Top             =   1305
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   7
         Top             =   2025
         Width           =   420
      End
   End
   Begin VB.Image imgOP 
      Height          =   255
      Left            =   6255
      Top             =   3045
      Width           =   300
   End
   Begin VB.Image imgNE 
      Height          =   255
      Left            =   5925
      Top             =   3045
      Width           =   300
   End
   Begin VB.Image imgFX 
      Height          =   255
      Left            =   5610
      Top             =   3045
      Width           =   300
   End
   Begin VB.Image imgIE 
      Height          =   255
      Left            =   5295
      Top             =   3045
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
      Index           =   6
      Left            =   3690
      TabIndex        =   13
      Top             =   3060
      Width           =   1485
   End
   Begin VB.Image picImgOff 
      Height          =   270
      Left            =   750
      Top             =   6735
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image picImgOn 
      Height          =   270
      Left            =   300
      Top             =   6750
      Visible         =   0   'False
      Width           =   345
   End
End
Attribute VB_Name = "frmRollover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub BuscarImg(ByVal Imagen As Integer)

    Dim glosa As String
    Dim Archivo As String
    
    glosa = "CompuServe Graphics Interchange (*.gif)|*.gif|"
    glosa = glosa & "JPG (*.jpg)|*.jpg|"
    glosa = glosa & "All files (*.*)|*.*"
    
    If LastPath = "" Then LastPath = App.Path
    
    If Not Cdlg.VBGetOpenFileName(Archivo, , , , , , glosa, , LastPath, "Open Image", "gif") Then
        Exit Sub
    End If
        
    On Error Resume Next
    If Imagen = 1 Then
        Set picImgOn.Picture = Nothing
        Set picImgOn = LoadPicture(Archivo)
        If Err <> 0 Then
            MsgBox "Error loading image.", vbCritical
            Set picImgOn.Picture = Nothing
            Exit Sub
        End If
        txtImgOn.Text = Archivo
    Else
        Set picImgOff.Picture = Nothing
        Set picImgOff = LoadPicture(Archivo)
        If Err <> 0 Then
            MsgBox "Error loading image.", vbCritical
            Set picImgOff.Picture = Nothing
            Exit Sub
        End If
        txtImgOff.Text = Archivo
    End If
    Err = 0
    
End Sub

Private Function CrearRollover() As Boolean

    Dim str As New cStringBuilder
    Dim funcion As String
    
    If txtImgOn.Text = "" Then
        MsgBox "Select Image On First", vbCritical
        Exit Function
    End If
    
    If txtImgOff.Text = "" Then
        MsgBox "Select Image Off First", vbCritical
        Exit Function
    End If
    
    If txtOnName.Text = "" Then
        MsgBox "Must select a name for image on", vbCritical
        txtOnName.SetFocus
        Exit Function
    End If
    
    If txtOffName.Text = "" Then
        MsgBox "Must select a name for image off", vbCritical
        txtOffName.SetFocus
        Exit Function
    End If
    
    If txtName.Text = "" Then
        MsgBox "Select a name for this rollover", vbCritical
        txtName.SetFocus
        Exit Function
    End If
    
    funcion = "javascript:"
    If cboFun.Text = "" Then
        funcion = funcion & "void()"
    Else
        funcion = funcion & cboFun.Text
    End If
    
    str.Append "<html>" & vbNewLine
    str.Append "<head>" & vbNewLine
    str.Append "<meta http-equiv=" & Chr$(34) & "Content-Type" & Chr$(34) & "CONTENT=" & Chr$(34) & "text/html;" & Chr$(34) & ">" & vbNewLine
    str.Append "<title>Testing Page</title>" & vbNewLine
    str.Append "</head>" & vbNewLine
    str.Append "<body>" & vbNewLine
    
    str.Append "<script language=""JavaScript"">" & vbNewLine
    'str.Append "<!--" & vbNewLine
    str.Append vbTab & txtOnName & "=new Image(" & picImgOn.Height & "," & picImgOn.Width & ");" & vbNewLine
    str.Append vbTab & txtOnName & ".src=" & Chr$(34) & util.VBArchivoSinPath(txtImgOn.Text) & Chr$(34) & ";" & vbNewLine
    str.Append vbTab & txtOffName & "=new Image(" & picImgOff.Height & "," & picImgOff.Width & ");" & vbNewLine
    str.Append vbTab & txtOffName & ".src=" & Chr$(34) & util.VBArchivoSinPath(txtImgOff.Text) & Chr$(34) & ";" & vbNewLine
    'str.Append "-->" & vbNewLine
    str.Append "</script>" & vbNewLine
    str.Append "<a href=" & Chr$(34) & funcion & Chr$(34) & " onMouseOver=" & Chr$(34) & "document." & txtName.Text & ".src = " & txtOnName & ".src" & Chr$(34)
    str.Append " onMouseOut=" & Chr$(34) & "document." & txtName.Text & ".src = " & txtOffName.Text & ".src" & Chr$(34) & ">" & vbNewLine
    str.Append "<img name=" & Chr$(34) & txtName.Text & Chr$(34) & " border=" & Chr$(34) & "0" & Chr$(34) & " src=" & Chr$(34) & _
               util.VBArchivoSinPath(txtImgOff.Text) & Chr$(34) & " height=" & Chr$(34) & picImgOff.Height & Chr$(34) & _
               " width=" & Chr$(34) & picImgOff.Width & Chr$(34) & ">" & vbNewLine
    str.Append "</a>" & vbNewLine

    str.Append "</body>" & vbNewLine
    str.Append "</html>" & vbNewLine
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(str.ToString)
    End If
    
    Set str = Nothing
    
    CrearRollover = True
    
End Function

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If CrearRollover() Then
            Unload Me
        End If
    ElseIf Index = 1 Then
        Unload Me
    ElseIf Index = 2 Then
        Call BuscarImg(1)
    ElseIf Index = 3 Then
        Call BuscarImg(2)
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
        
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call clear_memory(Me)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload : " & Me.Name
    Set frmRollover = Nothing
End Sub


