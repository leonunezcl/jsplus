VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Viewer / Extractor"
   ClientHeight    =   6630
   ClientLeft      =   4275
   ClientTop       =   2850
   ClientWidth     =   6720
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   6720
   Begin VB.Frame fra 
      Caption         =   "Navigation"
      Height          =   975
      Index           =   2
      Left            =   2640
      TabIndex        =   55
      Top             =   5400
      Width           =   3975
      Begin VB.CommandButton cmdPrev 
         Caption         =   "<--"
         Enabled         =   0   'False
         Height          =   375
         Left            =   165
         TabIndex        =   64
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   2745
         TabIndex        =   57
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "-->"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1455
         TabIndex        =   56
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComctlLib.ProgressBar Prog1 
      Height          =   150
      Left            =   4560
      TabIndex        =   54
      Top             =   6450
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   53
      Top             =   6375
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Caption         =   " Extracted Icons from selected File"
      Height          =   5295
      Index           =   1
      Left            =   2640
      TabIndex        =   4
      Top             =   75
      Width           =   3975
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   47
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   52
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   46
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   51
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   45
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   50
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   44
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   49
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   43
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   48
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   42
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   47
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   41
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   46
         Top             =   3960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   40
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   45
         Top             =   3960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   39
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   44
         Top             =   3960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   38
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   43
         Top             =   3960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   37
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   42
         Top             =   3960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   36
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   41
         Top             =   3960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   35
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   40
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   34
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   39
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   33
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   38
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   32
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   37
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   31
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   36
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   30
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   35
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   29
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   34
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   28
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   33
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   27
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   32
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   26
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   31
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   25
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   30
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   24
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   29
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   23
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   28
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   22
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   27
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   21
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   26
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   20
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   25
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   19
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   24
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   18
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   23
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   12
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   22
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   13
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   21
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   14
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   20
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   15
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   19
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   16
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   18
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   17
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   17
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   10
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   16
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   11
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   15
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   14
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   8
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   12
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   9
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   11
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   3240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   510
         TabIndex        =   9
         Top             =   360
         Width           =   510
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   1440
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   840
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fra 
      Caption         =   " Select a File "
      Height          =   6315
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   2535
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Icon"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1545
         TabIndex        =   63
         ToolTipText     =   "Copy Icon to Clipboard"
         Top             =   5535
         Width           =   870
      End
      Begin VB.PictureBox PicSelec 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   135
         ScaleHeight     =   495
         ScaleWidth      =   510
         TabIndex        =   62
         Top             =   5295
         Width           =   510
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         Height          =   1665
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   2295
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   1785
         Left            =   120
         Pattern         =   "*.exe;*.dll"
         TabIndex        =   2
         Top             =   3165
         Width           =   2295
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   555
         Width           =   2295
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Current Icon Selection"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   61
         Top             =   5025
         Width           =   1575
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Folders"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   60
         Top             =   945
         Width           =   510
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Drives"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   59
         Top             =   300
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Select a file to view contents"
         Height          =   180
         Left            =   120
         TabIndex        =   58
         Top             =   2925
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Cdlg As New cCommonDialog
Private fHayMas As Boolean
Private nTotal As Integer
Private nSiguiente As Integer

Private Sub HabilitaBotones()

    If nSiguiente > 1 Then
        If (nSiguiente - 1) Mod 47 > 0 Then
            cmdPrev.Enabled = True
        Else
            cmdPrev.Enabled = False
        End If
    Else
        cmdPrev.Enabled = False
    End If
    
    cmdNext.Enabled = fHayMas
        
End Sub

Private Sub cmdPrev_Click()

    Call clrPicbox
    NumberOfIcon = ExtractIconAndShow(Dir1.Path, File1.fileName, -1, "D", fHayMas, nTotal, nSiguiente)
    Call HabilitaBotones
    
End Sub

Private Sub cmdSave_Click()

    Dim Archivo As String
    
    If Not PicSelec.Picture Is Nothing Then
        If Not Cdlg.VBGetSaveFileName(Archivo, , , , , App.Path, "Save As ...", "ico") Then
            Exit Sub
        End If
    End If
    
    SavePicture PicSelec.Picture, Archivo
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMain = Nothing
End Sub


Private Sub Pic1_Click(Index As Integer)
    PicSelec.Height = Pic1(Index).Height
    PicSelec.Width = Pic1(Index).Width
    PicSelec.Picture = Pic1(Index).Image
    cmdSave.Enabled = True
End Sub

'Clear all Picture Box contents
Private Sub clrPicbox()
    
    PicSelec.Cls
    
    For I = 0 To 47
        Pic1(I).Cls
        Pic1(I).AutoRedraw = True
        Pic1(I).Picture = Nothing
    Next I
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    Call clrPicbox
    NumberOfIcon = ExtractIconAndShow(Dir1.Path, File1.fileName, -1, "A", fHayMas, nTotal, nSiguiente)
    Call HabilitaBotones
End Sub


Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    'lblInfo.Caption = "Status : You are in " & Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()

    Call clrPicbox
        
    If File1.fileName <> "" Then
        fHayMas = False
        nTotal = 0
        nSiguiente = 0
        NumberOfIcon = ExtractIconAndShow(Dir1.Path, File1.fileName, -1, "A", fHayMas, nTotal, nSiguiente)
        Call HabilitaBotones
    End If
    
End Sub

