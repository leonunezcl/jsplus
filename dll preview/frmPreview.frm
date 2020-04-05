VERSION 5.00
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   Caption         =   "Presentación Preliminar"
   ClientHeight    =   7500
   ClientLeft      =   330
   ClientTop       =   1425
   ClientWidth     =   8355
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   8355
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picControl 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7575
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.PictureBox picPages 
         Height          =   315
         Left            =   4440
         ScaleHeight     =   255
         ScaleWidth      =   1155
         TabIndex        =   14
         Top             =   60
         Width           =   1215
         Begin VB.Label lblStatus 
            Caption         =   "Pág:"
            Height          =   255
            Left            =   30
            TabIndex        =   15
            Top             =   15
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Siguiente"
         Height          =   315
         Left            =   3480
         TabIndex        =   13
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Previo"
         Height          =   315
         Left            =   2520
         TabIndex        =   12
         Top             =   60
         Width           =   855
      End
      Begin VB.ComboBox cboZoom 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   60
         Width           =   1815
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Imprimir"
         Height          =   315
         Left            =   5760
         TabIndex        =   2
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Cerrar"
         Height          =   315
         Left            =   6720
         TabIndex        =   1
         Top             =   60
         Width           =   855
      End
      Begin VB.Label lblView 
         Caption         =   "Ver"
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   90
         Width           =   495
      End
   End
   Begin VB.PictureBox picScroll 
      Height          =   6705
      Left            =   0
      ScaleHeight     =   6645
      ScaleWidth      =   8205
      TabIndex        =   3
      Top             =   600
      Width           =   8265
      Begin VB.VScrollBar vsPreview 
         Height          =   1215
         Left            =   3000
         TabIndex        =   9
         Top             =   840
         Width           =   255
      End
      Begin VB.HScrollBar hsPreview 
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   4560
         Width           =   1725
      End
      Begin VB.PictureBox picShow 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   240
         ScaleHeight     =   3615
         ScaleWidth      =   3540
         TabIndex        =   4
         Top             =   480
         Width           =   3540
         Begin VB.PictureBox picHold 
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   240
            ScaleHeight     =   1815
            ScaleWidth      =   2175
            TabIndex        =   6
            Top             =   120
            Width           =   2175
            Begin VB.PictureBox picDoc 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1215
               Left            =   240
               ScaleHeight     =   1215
               ScaleWidth      =   1695
               TabIndex        =   7
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.PictureBox picNormal 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   480
            ScaleHeight     =   1215
            ScaleWidth      =   1695
            TabIndex        =   5
            Top             =   2040
            Visible         =   0   'False
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================
' Project:      Enhance Print object
' Author:       edward moth
' Copyright:    © 2000 qbd software ltd
'
' ==============================================================
' Module:       frmPreview
' Purpose:      Display Print Preview
' ==============================================================


Option Explicit
Private mDocument As jsPrinter
Private bScrollCode As Boolean
Private sZoom As Single
Private lPage As Integer
Private lPageMax As Integer


Public Property Set Document(ByVal vNewValue As jsPrinter)

Set mDocument = vNewValue

End Property



Private Sub cboZoom_Click()

Dim iEvents As Integer

If Not bScrollCode Then
  If cboZoom.ListIndex >= 0 Then
' Because the Zoom_Check procedure can take some time
' the following line will close the dropdown
    iEvents = DoEvents
    If cboZoom.ItemData(cboZoom.ListIndex) <> sZoom Then
      sZoom = cboZoom.ItemData(cboZoom.ListIndex)
      Zoom_Check
    End If
  End If
End If

End Sub

Private Sub cmdClose_Click()

Set mDocument = Nothing
Unload Me

End Sub

Private Sub cmdNext_Click()

lPage = lPage + 1
Preview_Display lPage

End Sub

Private Sub cmdPrevious_Click()

lPage = lPage - 1
Preview_Display lPage


End Sub

Private Sub cmdPrint_Click()

Dim lPrintStart As Integer
Dim lPrintEnd As Integer

' Display the Print Options form
' Pass current page details
frmPrint.PageCurrent = lPage
frmPrint.PageMax = lPageMax
frmPrint.Show vbModal
lPrintStart = frmPrint.PageStart
lPrintEnd = frmPrint.PageEnd

If frmPrint.PrintDoc Then
  lblStatus.Caption = "Imprimiendo ..."
  lblStatus.Refresh
  mDocument.PrintDoc lPrintStart, lPrintEnd
  lblStatus.Caption = "Página: " & lPage & " / " & lPageMax
End If

Unload frmPrint

End Sub

Private Sub Form_Load()
sZoom = 100
With cboZoom
  .AddItem "100 %"
  .ItemData(.ListCount - 1) = 100
  .AddItem "75 %"
  .ItemData(.ListCount - 1) = 75
  .AddItem "50 %"
  .ItemData(.ListCount - 1) = 50
  .AddItem "Full"
  .ItemData(.ListCount - 1) = 0
  .AddItem "Full Ancho"
  .ItemData(.ListCount - 1) = -1
  bScrollCode = True
  .ListIndex = 0
  bScrollCode = False
End With
sZoom = 100
lPage = 1
Form_Resize

lPageMax = mDocument.Pages
Preview_Display lPage

End Sub

Public Sub Preview_Display(ByVal iPage As Integer)

Dim iMin As Integer
Dim iMax As Integer
Screen.MousePointer = vbHourglass
picNormal.Cls
mDocument.PreviewPage picNormal, iPage
Preview_Status

Zoom_Check
Screen.MousePointer = vbDefault
End Sub
Private Sub Zoom_Check()

Dim sSizeX As Single
Dim sSizeY As Single
Dim sRatio As Single
Dim spImage As StdPicture
Dim sWidth As Single
Dim sHeight As Single
Dim bScroll As Byte
Dim bOldScroll As Byte
Screen.MousePointer = vbHourglass

sWidth = picScroll.ScaleWidth
sHeight = picScroll.ScaleHeight
' Check the height and width to determine whether scroll bars
' are required.  This is in a loop because if a scroll bar is
' required it will affect the opposite dimension of the page
' display
Do
  bOldScroll = bScroll
  If sZoom = 0 Then
    sRatio = (sHeight - 480) / picNormal.Height
  ElseIf sZoom = -1 Then
    sRatio = (sWidth - 480) / picNormal.Width
  Else
    sRatio = sZoom / 100
  End If
  sSizeX = picNormal.Width * sRatio
  sSizeY = picNormal.Height * sRatio
  If sSizeX > sWidth And (bScroll And 1) <> 1 Then
    sHeight = sHeight - hsPreview.Height
    bScroll = bScroll + 1
  End If
  If sSizeY > sHeight And (bScroll And 2) <> 2 Then
    sWidth = sWidth - vsPreview.Width
    bScroll = bScroll + 2
  End If
Loop While bOldScroll <> bScroll

vsPreview.Height = sHeight
hsPreview.Width = sWidth

picShow.Move 0, 0, sWidth, sHeight
picDoc.Move 240, 240, sSizeX, sSizeY
picDoc.Cls
picDoc.PaintPicture picNormal.Image, 0, 0, sSizeX, sSizeY

' Display scroll bars if required
bScrollCode = True
picHold.Move 0, 0, sSizeX + 480, sSizeY + 480
If (bScroll And 2) = 2 Then
  vsPreview.Visible = True
  vsPreview.Max = (picHold.ScaleHeight - picShow.ScaleHeight) / 14.4 + 1
  vsPreview.Min = 0
  vsPreview.SmallChange = 14
  vsPreview.LargeChange = picShow.ScaleHeight / 14.4
  vsPreview.Value = vsPreview.Min
Else
  vsPreview.Visible = False
End If

If (bScroll And 1) = 1 Then
  hsPreview.Visible = True
  hsPreview.Max = (picHold.ScaleWidth - picShow.ScaleWidth) / 14.4 + 1
  hsPreview.Min = 0
  hsPreview.SmallChange = 14
  hsPreview.LargeChange = picShow.ScaleWidth / 14.4
  hsPreview.Value = hsPreview.Min
Else
  hsPreview.Visible = False
End If
bScrollCode = False
Screen.MousePointer = vbDefault

End Sub



Private Sub Form_Resize()

    If WindowState <> vbMinimized Then
        picScroll.Move 0, picControl.Height, Me.ScaleWidth, Me.ScaleHeight - picControl.Height
        vsPreview.Move picScroll.ScaleWidth - vsPreview.Width, 0, vsPreview.Width, picScroll.ScaleHeight - hsPreview.Height
        hsPreview.Move 0, picScroll.ScaleHeight - hsPreview.Height, picScroll.ScaleWidth - vsPreview.Width
        picShow.Move 0, 0, picScroll.ScaleWidth, picScroll.ScaleHeight
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPreview = Nothing
End Sub

Private Sub hsPreview_Change()

If Not bScrollCode Then
  picHold.Left = -hsPreview.Value * 14.4
End If

End Sub



Private Sub vsPreview_Change()
If Not bScrollCode Then
  picHold.Top = -vsPreview.Value * 14.4
End If

End Sub

Public Sub Preview_Status()
cmdPrevious.Enabled = CBool(lPage > 1)
cmdNext.Enabled = CBool(lPage < lPageMax)
lblStatus.Caption = "Página: " & lPage & " / " & lPageMax

End Sub
