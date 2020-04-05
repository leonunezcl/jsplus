VERSION 5.00
Object = "{866F095F-113F-4DC1-B803-F4CF4AFC96EE}#1.0#0"; "vbspgbbar.ocx"
Begin VB.Form frmInfoAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   6255
   ClientLeft      =   4260
   ClientTop       =   1950
   ClientWidth     =   6660
   ControlBox      =   0   'False
   Icon            =   "frmInfoAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdComprar 
      Caption         =   "&Buy"
      Height          =   375
      Left            =   3840
      TabIndex        =   15
      Tag             =   "http://www.regnow.com/softsell/nph-softsell.cgi?item=12453-1"
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   5760
      Width           =   1215
   End
   Begin vbsprgbar.ucProgressBar pgb 
      Height          =   210
      Left            =   135
      TabIndex        =   13
      Top             =   5040
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   370
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   12937777
      Max             =   20
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3330
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmInfoAbout.frx":000C
      Top             =   1275
      Width           =   6525
   End
   Begin VB.Label lbluser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Agatha Nurse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   4770
      Width           =   2070
   End
   Begin VB.Label lblreg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This software is registered to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   4755
      Width           =   2535
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "UNREGISTERED VERSION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   1
      Left            =   4320
      TabIndex        =   10
      Top             =   960
      Width           =   2265
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "All rights reserved."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   105
      TabIndex        =   9
      Top             =   300
      Width           =   1605
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2002-2009 Luis Nunez"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   16
      Left            =   105
      TabIndex        =   8
      Top             =   75
      Width           =   2910
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   105
      Picture         =   "frmInfoAbout.frx":0012
      Top             =   945
      Width           =   315
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "VBSoftware"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   105
      TabIndex        =   7
      Top             =   510
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Santiago, Chile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   105
      TabIndex        =   6
      Top             =   720
      Width           =   1305
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   6045
      Picture         =   "frmInfoAbout.frx":00A4
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "JavaScript Plus!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   4095
      TabIndex        =   5
      Top             =   30
      Width           =   1920
   End
   Begin VB.Label lblv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4230
      TabIndex        =   4
      Top             =   330
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Beta 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4230
      TabIndex        =   3
      Top             =   555
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   0
      Top             =   -15
      Width           =   6645
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   6765
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.vbsoftware.cl"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3840
      MouseIcon       =   "frmInfoAbout.frx":03AE
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Tag             =   "http://www.vbsoftware.cl"
      Top             =   5310
      Width           =   2175
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product and Company Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   1
      Top             =   5310
      Width           =   2475
   End
End
Attribute VB_Name = "frmInfoAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fexpired As Boolean
Private Sub cmd_Click()
    Unload Me
End Sub

Private Sub cmdComprar_Click()
    util.ShellFunc cmdComprar.Tag, vbNormalFocus
End Sub

Private Sub Form_Activate()
    If fexpired Then
        frmMain.fexpired = True
        frmTriExp.Show vbModal
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()

    Dim src As New cStringBuilder
    Dim ini As String
    Dim dat As String
    Dim F As String
    Dim fi As String
    Dim us As String
    Dim na As String
    Dim linea As String
    Dim nFreeFile As Long
    Dim C As Integer
    Dim cnt As String
    Dim val As String
    
    util.CenterForm Me
        
    lblv.Caption = App.Major & "." & App.Minor & "." & App.Revision
    
lit:
    #If LITE = 1 Then
        If frmMain.paulina Then
            lbl(1).Caption = Chr$(85) & Chr$(78) & Chr$(82) & Chr$(69) & Chr$(71) & Chr$(73) & Chr$(83) & Chr$(84) & Chr$(69) & Chr$(82) & Chr$(69) & Chr$(68) & Chr$(32) & Chr$(86) & Chr$(69) & Chr$(82) & Chr$(83) & Chr$(73) & Chr$(79) & Chr$(78)
            F = Chr$(50) & Chr$(48)
            lblreg.Caption = "Unregistered copy"
            lbluser.Caption = vbNullString
            lbluser.Visible = False
            pgb.Value = CInt(F) - frmMain.palic
            lblreg.Caption = "Your free " & F & "-usage trial period has " & pgb.Value & " uses remaining."
        Else
            cnt = Base64Encode(Chr$(99) & Chr$(110) & Chr$(116))
            val = Base64Encode(Chr$(118) & Chr$(97) & Chr$(108))
            lbl(1).Caption = Chr$(85) & Chr$(78) & Chr$(82) & Chr$(69) & Chr$(71) & Chr$(73) & Chr$(83) & Chr$(84) & Chr$(69) & Chr$(82) & Chr$(69) & Chr$(68) & Chr$(32) & Chr$(86) & Chr$(69) & Chr$(82) & Chr$(83) & Chr$(73) & Chr$(79) & Chr$(78)
            'lblreg.Visible = False
            lblreg.Caption = "Unregistered copy"
            lbluser.Caption = vbNullString
            lbluser.Visible = False
            fi = Base64Encode(Chr$(117) & Chr$(115) & Chr$(101) & Chr$(114) & Chr$(51) & Chr$(50) & Chr$(46) & Chr$(100) & Chr$(97) & Chr$(116))
            dat = Base64Encode(util.StripPath(util.SysDir)) & fi
            F = Chr$(50) & Chr$(48)
            nFreeFile = FreeFile
            C = 1
            If Not ArchivoExiste2(Base64Decode(dat)) Then
                
                pgb.Value = CInt(F)
                pgb.Max = CInt(F)
                
                lblreg.Caption = "Your free " & F & "-usage trial period has " & CInt(F) & " uses remaining."
                
                Open Base64Decode(dat) For Output As #nFreeFile
                    Print #nFreeFile, cnt
                    Print #nFreeFile, val & Base64Encode(Chr$(61) & Chr$(48))
                Close #nFreeFile
            Else
                Open Base64Decode(dat) For Input As #nFreeFile
                    Do While Not EOF(nFreeFile)
                        Line Input #nFreeFile, linea
                    Loop
                Close #nFreeFile
                    
                If Len(Explode(Base64Decode(linea), 2, Chr$(61))) > 0 Then
                    If CInt(Explode(Base64Decode(linea), 2, Chr$(61))) >= CInt(F) Then
                        fexpired = True
                    Else
                        If CInt(Explode(Base64Decode(linea), 2, Chr$(61))) <= CInt(F) Then
                            pgb.Value = CInt(F) - CInt(Explode(Base64Decode(linea), 2, Chr$(61)))
                            lblreg.Caption = "Your free " & F & "-usage trial period has " & CInt(F) - CInt(Explode(Base64Decode(linea), 2, Chr$(61))) & " uses remaining."
                        Else
                            pgb.Value = 0
                            lblreg.Caption = "Your free " & F & "-usage trial period has " & Chr$(48) & " uses remaining."
                        End If
                    End If
                Else
                    MsgBox "Failed to start Javascript Plus!", vbCritical
                    End
                End If
            End If
        End If
    #Else
        fi = Base64Encode(Chr$(114) & Chr$(101) & Chr$(103) & Chr$(117) & Chr$(115) & Chr$(101) & Chr$(114) & Chr$(46) & Chr$(105) & Chr$(110) & Chr$(105))
        ini = Base64Encode(util.StripPath(App.Path)) & fi
        If ArchivoExiste2(Base64Decode(ini)) Then
            cmdComprar.Visible = False
            lbl(1).Caption = Chr$(82) & Chr$(69) & Chr$(71) & Chr$(73) & Chr$(83) & Chr$(84) & Chr$(69) & Chr$(82) & Chr$(69) & Chr$(68) & Chr$(32) & Chr$(86) & Chr$(69) & Chr$(82) & Chr$(83) & Chr$(73) & Chr$(79) & Chr$(78)
            us = Chr$(117) & Chr$(115) & Chr$(101) & Chr$(114)
            na = Chr$(110) & Chr$(97) & Chr$(109) & Chr$(101)
            lbluser.Caption = util.LeeIni(Base64Decode(ini), us, na)
            lbluser.Visible = True
            pgb.Visible = False
            Label1.Visible = False
        Else
            GoTo lit:
        End If
    #End If
    
    src.Append "Dear Customer," & vbNewLine
    src.Append "" & vbNewLine
    src.Append "Thank you for using this software. I hope it helps you getting "
    src.Append "your job done faster and easier." & vbNewLine
    src.Append "" & vbNewLine
    src.Append "If you encounter any problems using this software, please, "
    src.Append "feel free to contact me via e-mail to "
    src.Append "contact@vbsoftware.cl so we can help you." & vbNewLine
    src.Append "" & vbNewLine
    src.Append "Feature suggestions, improvement ideas and other contributions "
    src.Append "are very welcome. Please send them to contact@vbsoftware.cl" & vbNewLine
    src.Append "" & vbNewLine
    src.Append "Thank you very much!" & vbNewLine
    src.Append "" & vbNewLine
    src.Append "Sincerely," & vbNewLine
    src.Append "Luis Nunez" & vbNewLine
    src.Append "The Author"

    txtInfo.Text = src.ToString
    
    Image1.Picture = LoadResPicture(1003, vbResBitmap)
    
    Debug.Print "load"
    
    Set src = Nothing
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim dat As String
    Dim ini As String
    Dim fi As String
    Dim V As Variant
    Dim linea As String
    Dim nFreeFile As Long
    Dim cnt As String
    Dim val As String
    
    frmMain.paulina = True
lit:
    #If LITE = 1 Then
        If Not started Then
            If Not fexpired Then
                fi = Base64Encode(Chr$(117) & Chr$(115) & Chr$(101) & Chr$(114) & Chr$(51) & Chr$(50) & Chr$(46) & Chr$(100) & Chr$(97) & Chr$(116))
                dat = Base64Encode(util.StripPath(util.SysDir)) & fi
                
                nFreeFile = FreeFile
                Open Base64Decode(dat) For Input As #nFreeFile
                    Do While Not EOF(nFreeFile)
                        Line Input #nFreeFile, linea
                    Loop
                Close #nFreeFile
                
                If CInt(Explode(Base64Decode(linea), 2, Chr$(61))) > 0 Then
                    frmMain.palic = CInt(Explode(Base64Decode(linea), 2, Chr$(61)))
                    V = CInt(Explode(Base64Decode(linea), 2, Chr$(61))) + 1
                Else
                    V = Chr$(49)
                    frmMain.palic = V
                End If
                
                cnt = Base64Encode(Chr$(99) & Chr$(110) & Chr$(116))
                val = Base64Encode(Chr$(118) & Chr$(97) & Chr$(108))
        
                Open Base64Decode(dat) For Output As #nFreeFile
                    Print #nFreeFile, cnt
                    Print #nFreeFile, Base64Encode("val=" & V)
                Close #nFreeFile
                
                set_file_time
            End If
        End If
    #Else
        fi = Base64Encode(Chr$(114) & Chr$(101) & Chr$(103) & Chr$(117) & Chr$(115) & Chr$(101) & Chr$(114) & Chr$(46) & Chr$(105) & Chr$(110) & Chr$(105))
        ini = Base64Encode(util.StripPath(App.Path)) & fi
        If Not ArchivoExiste2(Base64Decode(ini)) Then
            GoTo lit
        End If
    #End If
    started = True
    
    Call clear_memory(Me)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmInfoAbout = Nothing
End Sub





Private Sub lblURL_Click()
    util.ShellFunc lblURL.Tag, vbNormalFocus
End Sub


