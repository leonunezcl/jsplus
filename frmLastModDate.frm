VERSION 5.00
Begin VB.Form frmLastModDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Last Modified Date"
   ClientHeight    =   2055
   ClientLeft      =   4515
   ClientTop       =   2955
   ClientWidth     =   5955
   Icon            =   "frmLastModDate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   1020
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   5835
      Begin VB.TextBox txtMsg 
         Height          =   285
         Left            =   150
         TabIndex        =   0
         Top             =   480
         Width           =   5550
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Input text to display:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   255
         Width           =   1410
      End
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
      Left            =   3030
      TabIndex        =   3
      Top             =   1710
      Width           =   1485
   End
   Begin VB.Image imgIE 
      Height          =   255
      Left            =   4635
      Top             =   1695
      Width           =   300
   End
   Begin VB.Image imgFX 
      Height          =   255
      Left            =   4950
      Top             =   1695
      Width           =   300
   End
   Begin VB.Image imgNE 
      Height          =   255
      Left            =   5265
      Top             =   1695
      Width           =   300
   End
   Begin VB.Image imgOP 
      Height          =   255
      Left            =   5595
      Top             =   1695
      Width           =   300
   End
End
Attribute VB_Name = "frmLastModDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function CrearFecha() As Boolean

    Dim texto As String
    Dim src As New cStringBuilder
    texto = txtMsg.Text

    src.Append "<html>" & vbNewLine
    src.Append "<head>" & vbNewLine
    src.Append "<meta http-equiv=" & Chr$(34) & "Content-Type" & Chr$(34) & "CONTENT=" & Chr$(34) & "text/html;" & Chr$(34) & ">" & vbNewLine
    src.Append "<title>Testing Page</title>" & vbNewLine
    src.Append "</head>" & vbNewLine
    src.Append "<body>" & vbNewLine
    
    src.Append "<script language=" & Chr$(34) & "Javascript" & Chr$(34) & " type=" & Chr$(34) & "text/javascript" & Chr$(34) & ">" & vbNewLine
    src.Append "document.write(" & Chr$(34) & texto & Chr$(34) & "+document.lastModified);" & vbNewLine
    src.Append "</script>" & vbNewLine
    
    src.Append "</body>" & vbNewLine
    src.Append "</html>" & vbNewLine
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(src.ToString)
    End If
    
    Set src = Nothing
    
    CrearFecha = True
    
End Function

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If CrearFecha() Then
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
    
    Debug.Print "load : " & Me.Name
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload : " & Me.Name
    Set frmLastModDate = Nothing
End Sub


