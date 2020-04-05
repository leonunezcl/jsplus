VERSION 5.00
Begin VB.Form frmSetTreeExp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configure Extensions"
   ClientHeight    =   3750
   ClientLeft      =   4560
   ClientTop       =   4710
   ClientWidth     =   3960
   Icon            =   "frmSetTreeExp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   3555
      Left            =   105
      TabIndex        =   2
      Top             =   90
      Width           =   2460
      Begin VB.CommandButton cmd 
         Caption         =   "&Add"
         Height          =   285
         Index           =   2
         Left            =   1905
         TabIndex        =   6
         Top             =   270
         Width           =   450
      End
      Begin VB.TextBox txtFilExt 
         Height          =   285
         Left            =   1125
         TabIndex        =   0
         Top             =   270
         Width           =   735
      End
      Begin VB.ListBox lstExt 
         Height          =   2760
         Left            =   75
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   585
         Width           =   2295
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "File Extension:"
         Height          =   195
         Left            =   75
         TabIndex        =   3
         Top             =   300
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmSetTreeExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IniFileTree As String
Public origen As Integer
Private Sub cmd_Click(Index As Integer)

   Dim k As Integer
   Dim j As Integer
   
   Select Case Index
      Case 0   'ok
         BorrarArchivo IniFileTree
         
         j = 1
         For k = 0 To lstExt.ListCount - 1
            If lstExt.Selected(k) Then
               Call GrabaIni(IniFileTree, "tree", "ext" & j, lstExt.List(k))
               j = j + 1
            End If
         Next k
         
         Unload Me
      Case 1   'cancel
         Unload Me
      Case 2   'add
         Dim ext As String
         Dim found As Boolean
         
         ext = Trim$(txtFilExt.Text)
         
         If Len(ext) > 0 Then
            For k = 0 To lstExt.ListCount - 1
               If lstExt.List(k) = ext Then
                  found = True
                  Exit For
               End If
            Next k
            
            If Not found Then
               lstExt.AddItem ext
               lstExt.Selected(lstExt.NewIndex) = True
            Else
               MsgBox "File extension '" & ext & "' already exists", vbCritical
            End If
         End If
   End Select
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()

   Dim arr_ext() As String
   Dim k As Integer
   
   Hourglass hwnd, True
   
   CenterForm Me
   
   If origen = 0 Then
      IniFileTree = StripPath(App.Path) & "treeext.ini"
      get_info_section "tree", arr_ext, IniFileTree
   Else
      IniFileTree = StripPath(App.Path) & "batch.ini"
      get_info_section "tree", arr_ext, IniFileTree
   End If
   
   For k = 1 To UBound(arr_ext)
      lstExt.AddItem arr_ext(k)
      lstExt.Selected(lstExt.NewIndex) = True
   Next k
   
   Hourglass hwnd, False
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Debug.Print "unload"
   Set frmSetTreeExp = Nothing
End Sub


