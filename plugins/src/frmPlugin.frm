VERSION 5.00
Begin VB.Form frmPlugin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VBSoftware - Get Files from folder"
   ClientHeight    =   8025
   ClientLeft      =   3360
   ClientTop       =   2460
   ClientWidth     =   7755
   Icon            =   "frmPlugin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "Remove All"
      Height          =   495
      Index           =   5
      Left            =   4860
      TabIndex        =   16
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Remove"
      Height          =   495
      Index           =   4
      Left            =   3435
      TabIndex        =   15
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Add All"
      Height          =   495
      Index           =   3
      Left            =   2010
      TabIndex        =   14
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Add"
      Height          =   495
      Index           =   2
      Left            =   615
      TabIndex        =   13
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Files Selected"
      Height          =   3195
      Left            =   105
      TabIndex        =   11
      Top             =   4755
      Width           =   6105
      Begin VB.CheckBox chk 
         Caption         =   "Select All"
         Height          =   210
         Left            =   150
         TabIndex        =   17
         Top             =   2895
         Width           =   1080
      End
      Begin VB.ListBox lstFiles 
         Height          =   2535
         Left            =   135
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   255
         Width           =   5865
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Cancel"
      Height          =   495
      Index           =   1
      Left            =   6345
      TabIndex        =   10
      Top             =   825
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Ok"
      Height          =   495
      Index           =   0
      Left            =   6330
      TabIndex        =   9
      Top             =   225
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      Height          =   3990
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6105
      Begin VB.CheckBox chkAddPath 
         Caption         =   "Add Path + File ?"
         Height          =   195
         Left            =   3300
         TabIndex        =   18
         Top             =   3510
         Width           =   1635
      End
      Begin VB.TextBox txtFilter 
         Height          =   315
         Left            =   105
         TabIndex        =   8
         Text            =   "*.*"
         Top             =   3540
         Width           =   2145
      End
      Begin VB.FileListBox File1 
         Height          =   2235
         Left            =   105
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         Top             =   1050
         Width           =   2160
      End
      Begin VB.DirListBox Dir1 
         Height          =   2790
         Left            =   2370
         TabIndex        =   4
         Top             =   465
         Width           =   3615
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   105
         TabIndex        =   2
         Top             =   465
         Width           =   2175
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Filter"
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
         Index           =   3
         Left            =   105
         TabIndex        =   7
         Top             =   3315
         Width           =   435
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Files"
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
         Index           =   2
         Left            =   105
         TabIndex        =   5
         Top             =   825
         Width           =   405
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Folders"
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
         Index           =   1
         Left            =   2370
         TabIndex        =   3
         Top             =   255
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Drives"
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
         Index           =   0
         Left            =   105
         TabIndex        =   1
         Top             =   255
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Util As New cLibrary


Private Sub LockBtn(ByVal estado As Boolean)

   Dim k As Integer
   
   For k = 0 To cmd.Count - 1
      cmd(k).Enabled = estado
   Next k
   
End Sub

Private Sub chk_Click()

   Dim ret As Boolean
   Dim k As Integer
   
   If chk.Value = 1 Then
      ret = True
   End If
   
   For k = 0 To lstFiles.ListCount - 1
      lstFiles.Selected(k) = ret
   Next k
   
End Sub

Private Sub cmd_Click(Index As Integer)

   Dim k As Integer
   
   Select Case Index
      Case 0   'ok
         Util.Hourglass hWnd, True
         LockBtn False
         For k = lstFiles.ListCount - 1 To 0 Step -1
            If lstFiles.Selected(k) Then
               glbOutputString.Append lstFiles.List(k) & vbNewLine
            End If
         Next k
         Util.Hourglass hWnd, False
         Unload Me
      Case 1   'cancel
         Unload Me
      Case 2   'add
         Util.Hourglass hWnd, True
         LockBtn False
         For k = 0 To File1.ListCount - 1
            If File1.Selected(k) Then
               If chkAddPath.Value = 0 Then
                  lstFiles.AddItem File1.List(k)
               Else
                  lstFiles.AddItem StripPath(Dir1.Path) & File1.List(k)
               End If
               lstFiles.Selected(lstFiles.NewIndex) = True
            End If
         Next k
         LockBtn True
         Util.Hourglass hWnd, False
      Case 3   'add all
         Util.Hourglass hWnd, True
         LockBtn False
         For k = 0 To File1.ListCount - 1
            If chkAddPath.Value = 0 Then
               lstFiles.AddItem File1.List(k)
            Else
               lstFiles.AddItem StripPath(Dir1.Path) & File1.List(k)
            End If
            lstFiles.Selected(lstFiles.NewIndex) = True
         Next k
         LockBtn True
         Util.Hourglass hWnd, False
      Case 4   'remove
         Util.Hourglass hWnd, True
         LockBtn False
         For k = lstFiles.ListCount - 1 To 0 Step -1
            If lstFiles.Selected(k) Then
               lstFiles.RemoveItem k
            End If
         Next k
         LockBtn True
         Util.Hourglass hWnd, False
      Case 5   'remove all
         Util.Hourglass hWnd, True
         lstFiles.Clear
         Util.Hourglass hWnd, False
   End Select
   
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()

   On Error Resume Next
   
   Dir1.Path = Drive1.Drive
   
   Err = 0
   
End Sub


Private Sub Form_Load()
   Util.CenterForm Me
   DrawXPCtl Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set frmPlugin = Nothing
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
      If Trim$(txtFilter.Text) <> "" Then
         File1.Pattern = txtFilter.Text
      End If
   End If
   
End Sub


