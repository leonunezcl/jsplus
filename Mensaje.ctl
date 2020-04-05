VERSION 5.00
Object = "{E910F8E1-8996-4EE9-90F1-3E7C64FA9829}#1.1#0"; "vbaListView6.ocx"
Begin VB.UserControl Mensaje 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CommandButton cmd 
      Caption         =   "+"
      Height          =   225
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   270
   End
   Begin vbalListViewLib6.vbalListViewCtl lvwmsg 
      Height          =   1260
      Left            =   1260
      TabIndex        =   0
      Top             =   735
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   2223
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   1
      MultiSelect     =   -1  'True
      LabelEdit       =   0   'False
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
      AutoArrange     =   0   'False
      Appearance      =   0
      FlatScrollBar   =   -1  'True
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
End
Attribute VB_Name = "Mensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Expand(ByVal Estado As String)

Private Sub cmd_Click()
    If cmd.Caption = "+" Then
        cmd.Caption = "-"
    Else
        cmd.Caption = "+"
    End If
    RaiseEvent Expand(cmd.Caption)
End Sub


Private Sub UserControl_Initialize()

    With lvwmsg
        .Columns.Add , "k1", "Nº", , 500
        .Columns.Add , "k2", "Description", , 8000
    End With
    
End Sub


Private Sub UserControl_Resize()

    On Error Resume Next
    lvwmsg.Move cmd.Height + 50, 0, UserControl.Width - cmd.Width - 50, UserControl.Height
    Err = 0
    
End Sub


