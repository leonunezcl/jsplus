VERSION 5.00
Begin VB.Form frmPau 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unlock JavaScript Plus!"
   ClientHeight    =   1380
   ClientLeft      =   2910
   ClientTop       =   3000
   ClientWidth     =   4440
   Icon            =   "frmPau.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1920
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4710
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   3600
      TabIndex        =   4
      Top             =   390
      Width           =   780
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   2730
      TabIndex        =   3
      Top             =   390
      Width           =   780
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1860
      TabIndex        =   2
      Top             =   390
      Width           =   780
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   990
      TabIndex        =   1
      Top             =   390
      Width           =   780
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   390
      Width           =   780
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   1
      Left            =   2430
      TabIndex        =   6
      Top             =   840
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
      Index           =   0
      Left            =   540
      TabIndex        =   5
      Top             =   840
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
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Input the register code sent to email account:"
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
      TabIndex        =   7
      Top             =   105
      Width           =   3900
   End
End
Attribute VB_Name = "frmPau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Grk() As Boolean

    Dim creg As New cRegistry
    Dim s As String
    Dim nfreefile As Long
    Dim Archivo As String
    
    s = Base64Encode(txt(0).Text & "-" & txt(1).Text & "-" & txt(2).Text & "-" & txt(3).Text & "-" & txt(4).Text)
    
    creg.ClassKey = HKEY_CURRENT_USER
    creg.SectionKey = "Software\jspad"
    
    If Not creg.KeyExists Then
        If creg.CreateKey Then
            creg.ValueType = REG_EXPAND_SZ
            creg.ValueKey = "Startup"
            creg.Value = s
            
            Grk = True
        End If
    Else
        creg.ValueType = REG_EXPAND_SZ
        creg.ValueKey = "Startup"
        creg.Value = s
        
        Grk = True
    End If
    
    nfreefile = FreeFile
    
    Archivo = Util.StripPath(App.Path) & "licencia.dat"
    
    Open Archivo For Output As #nfreefile
        Print #nfreefile, s
    Close #nfreefile
    
End Function

Private Function Unl() As Boolean

    Dim k As Integer
    Dim arr_k() As String
    Dim ret As Boolean
    ReDim arr_k(12)
    
    arr_k(1) = "Q7UZ-MLZD-S7NV-9UBS-5D"
    arr_k(2) = "HLPZ-VNPK-5CCA-E4WL-E0"
    arr_k(3) = "L5QB-JJBQ-RXTD-LDG3-08"
    arr_k(4) = "4CWX-PQVP-HNLB-YX7Q-C0"
    arr_k(5) = "7NDJ-5CN3-GWX7-CE46-5B"
    arr_k(6) = "UFWV-MA2J-KDNX-F4Y8-19"
    arr_k(7) = "6QKT-3YMZ-VRX7-LTMW-68"
    arr_k(8) = "NRVT-5L76-G8LM-T3L6-66"
    arr_k(9) = "XRBU-BG2N-PX5S-74LS-00"
    arr_k(10) = "MP7U-VD6L-7LVK-X3N8-1D"
    arr_k(11) = "JD2Q-ZJZD-8A24-8GGJ-47"
    arr_k(12) = "U34Q-HJ5S-TKL5-J58Q-CE"

    For k = 0 To 4
        If Len(Trim$(txt(k).Text)) = 0 Then
            txt(k).SetFocus
            Exit Function
        End If
    Next k
    
    Dim pt1
    Dim pt2
    Dim pt3
    Dim pt4
    Dim pt5
    
    pt1 = txt(0).Text
    pt2 = txt(1).Text
    pt3 = txt(2).Text
    pt4 = txt(3).Text
    pt5 = txt(4).Text
    
    For k = 1 To 12
        If pt1 = Explode(arr_k(k), 1, "-") Then
            If pt2 = Explode(arr_k(k), 2, "-") Then
                If pt3 = Explode(arr_k(k), 3, "-") Then
                    If pt4 = Explode(arr_k(k), 4, "-") Then
                        If pt5 = Explode(arr_k(k), 5, "-") Then
                            ret = True
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next k
    
    Unl = ret
    
End Function

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If Unl() Then
            If Grk() Then
                MsgBox "Thank you! for register JavaScript Plus!.", vbInformation
                Unload Me
            End If
        Else
            MsgBox "Invalid code. Please try again.", vbCritical
            txt(0).SetFocus
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    Util.CenterForm Me
    
    set_color_form Me
    
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    cmd(0).Refresh
    cmd(1).Refresh
    
    DrawXPCtl Me
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmPau = Nothing
End Sub


