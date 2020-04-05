VERSION 5.00
Begin VB.Form frmNew 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New document "
   ClientHeight    =   2205
   ClientLeft      =   5550
   ClientTop       =   3240
   ClientWidth     =   4215
   Icon            =   "frmNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   4905
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2790
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Frame fra 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select"
      ForeColor       =   &H00FF8080&
      Height          =   1545
      Index           =   0
      Left            =   75
      TabIndex        =   7
      Top             =   90
      Width           =   4050
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "New Empty Document"
         ForeColor       =   &H00404040&
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   255
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "New Html Document"
         ForeColor       =   &H00404040&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   735
         Width           =   1770
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "New from predefined template"
         ForeColor       =   &H00404040&
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   975
         Width           =   2580
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "New from user template"
         ForeColor       =   &H00404040&
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   2145
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "New Frame"
         ForeColor       =   &H00404040&
         Height          =   240
         Index           =   4
         Left            =   240
         TabIndex        =   1
         Top             =   495
         Width           =   1230
      End
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   2
      Left            =   2505
      TabIndex        =   6
      Top             =   1695
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
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   1695
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
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub NuevoDocumento()

    Dim Indice As Integer
    
    Call frmMain.newEdit
            
    If opt(1).Value Then
        Dim sBuffer As New cStringBuilder
                
        sBuffer.Append "<!DOCTYPE HTML PUBLIC " & Chr$(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr$(34) & ">" & vbNewLine
        sBuffer.Append "<HTML>" & vbNewLine
        sBuffer.Append "<HEAD>" & vbNewLine
        sBuffer.Append "<TITLE></TITLE>" & vbNewLine
        sBuffer.Append "<META NAME=" & Chr$(34) & "Generator" & Chr$(34) & " CONTENT=" & Chr$(34) & "Javascript Plus!" & Chr$(34) & ">" & vbNewLine
        sBuffer.Append "<META NAME=" & Chr$(34) & "Author" & Chr$(34) & " CONTENT=" & Chr$(34) & Chr$(34) & ">" & vbNewLine
        sBuffer.Append "<META NAME=" & Chr$(34) & "Keywords" & Chr$(34) & " CONTENT=" & Chr$(34) & Chr$(34) & ">" & vbNewLine
        sBuffer.Append "<META NAME=" & Chr$(34) & "Description" & Chr$(34) & " CONTENT=" & Chr$(34) & Chr$(34) & ">" & vbNewLine
        sBuffer.Append "</HEAD>" & vbNewLine
        sBuffer.Append "<BODY>" & vbNewLine
        sBuffer.Append "<SCRIPT LANGUAJE=""JavaScript"">" & vbNewLine
        sBuffer.Append vbNewLine
        sBuffer.Append vbNewLine
        sBuffer.Append "</SCRIPT>" & vbNewLine
        sBuffer.Append "</BODY>" & vbNewLine
        sBuffer.Append "</HTML>"
        
        frmMain.ActiveForm.txtCode.Text = sBuffer.ToString
        frmMain.ActiveForm.txtCode.Modified = True
        Set sBuffer = Nothing
    ElseIf opt(2).Value Then
        frmPreTemplates.Show vbModal
    ElseIf opt(3).Value Then
        frmUserTemplate.Show vbModal
    ElseIf opt(4).Value Then
        sBuffer.Append "<FRAMESET ROWS=" & Chr$(34) & ", " & Chr$(34) & " COLS=" & Chr$(34) & "," & Chr$(34) & ">" & vbNewLine
        sBuffer.Append "<FRAME SRC="""" NAME="""">" & vbNewLine
        sBuffer.Append "<FRAME SRC="""" NAME="""">" & vbNewLine
        sBuffer.Append "</FRAMESET>" & vbNewLine
        frmMain.ActiveForm.txtCode.Text = sBuffer.ToString
        frmMain.ActiveForm.txtCode.Modified = True
        Set sBuffer = Nothing
    End If
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 1 Then
        Call NuevoDocumento
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

    Util.CenterForm Me
    set_color_form Me
    
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    cmd(0).Refresh
    cmd(1).Refresh
    
    Debug.Print "load"
    DrawXPCtl Me
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmNew = Nothing
End Sub


