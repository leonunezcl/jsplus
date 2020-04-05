VERSION 5.00
Begin VB.Form frmPreview 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Image Preview"
   ClientHeight    =   7995
   ClientLeft      =   2475
   ClientTop       =   2970
   ClientWidth     =   9630
   Icon            =   "frmImagePreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3435
      Left            =   30
      ScaleHeight     =   3405
      ScaleWidth      =   5985
      TabIndex        =   2
      Top             =   15
      Width           =   6015
      Begin VB.PictureBox picImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   420
         ScaleHeight     =   525
         ScaleWidth      =   585
         TabIndex        =   3
         Top             =   165
         Width           =   585
         Begin VB.Label lblShape 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   915
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   2205
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3315
      Left            =   6120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   390
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3855
      Width           =   3375
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    HScroll1.Visible = False
    VScroll1.SmallChange = 150
    VScroll1.LargeChange = 1500
    HScroll1.SmallChange = 150
    HScroll1.LargeChange = 1500
    
    Dim resp As Long
    
    picImage.Height = frmImageEffect.picSave.Height
    picImage.Width = frmImageEffect.picSave.Width
    
    GPX_BitBlt picImage.hdc, 0, 0, frmImageEffect.picSave.ScaleWidth, frmImageEffect.picSave.ScaleHeight, frmImageEffect.picSave.hdc, 0, 0, vbSrcCopy, resp
    picImage.Refresh
    DoEvents
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set picImage.Picture = Nothing
End Sub


Private Sub Form_Resize()

    Dim dy As Long
    If HScroll1.Visible Then dy = HScroll1.Height
    If WindowState = vbMinimized Then Exit Sub
    picContainer.Move 0, 0, ScaleWidth, ScaleHeight
    picImage.Move 0, 0
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmPreview = Nothing
End Sub


Private Sub HScroll1_Change()
    picImage.Left = -HScroll1.Value
End Sub


Private Sub picContainer_Resize()

    Dim nMax As Long, dy As Long
    On Error Resume Next
    nMax = picImage.Height - picContainer.Height
    If nMax < 0 Then nMax = 0
    VScroll1.Max = nMax
    If nMax > 0 Then
       VScroll1.Visible = True
       picContainer.Width = ScaleWidth - VScroll1.Width
       VScroll1.Move picContainer.Left + picContainer.Width, picContainer.Top, VScroll1.Width, picContainer.Height
    Else
       VScroll1.Visible = False
    End If
    nMax = picImage.Width - picContainer.Width
    If nMax < 0 Then nMax = 0
    HScroll1.Max = nMax
    If nMax > 0 Then
       HScroll1.Visible = True
       picContainer.Height = ScaleHeight - HScroll1.Height
       HScroll1.Move picContainer.Left, picContainer.Top + picContainer.Height, picContainer.Width
       dy = HScroll1.Height
    Else
       HScroll1.Visible = False
    End If
   
End Sub


Private Sub picImage_Resize()
    picContainer_Resize
End Sub


Private Sub VScroll1_Change()
    picImage.Top = -VScroll1.Value
End Sub


