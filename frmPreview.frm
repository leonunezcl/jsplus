VERSION 5.00
Begin VB.Form frmPreview 
   Caption         =   "Print Preview"
   ClientHeight    =   6570
   ClientLeft      =   705
   ClientTop       =   2910
   ClientWidth     =   9300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   9300
   Begin VB.Timer tmrProgress 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8640
      Top             =   1320
   End
   Begin VB.PictureBox picControl 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   8625
      TabIndex        =   0
      Top             =   0
      Width           =   8625
      Begin VB.CommandButton cmdEnd 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3600
         Picture         =   "frmPreview.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   60
         Width           =   315
      End
      Begin VB.CommandButton cmdStart 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         Picture         =   "frmPreview.frx":28EC
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   60
         Width           =   315
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go"
         Height          =   315
         Left            =   5640
         TabIndex        =   15
         Top             =   60
         Width           =   315
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   6000
         TabIndex        =   14
         Top             =   60
         Width           =   975
      End
      Begin VB.TextBox txtPage 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4455
         MaxLength       =   5
         TabIndex        =   13
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3240
         Picture         =   "frmPreview.frx":2A36
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   315
      End
      Begin VB.CommandButton cmdPrevious 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         Picture         =   "frmPreview.frx":2B80
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   315
      End
      Begin VB.ComboBox cboZoom 
         Height          =   315
         ItemData        =   "frmPreview.frx":2CCA
         Left            =   600
         List            =   "frmPreview.frx":2CCC
         TabIndex        =   9
         Text            =   "cboZoom"
         Top             =   60
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   315
         Left            =   7440
         TabIndex        =   1
         Top             =   60
         Width           =   855
      End
      Begin VB.Label lblTotal 
         Caption         =   "/ 0"
         Height          =   195
         Left            =   5280
         TabIndex        =   17
         Top             =   90
         Width           =   645
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pag:"
         Height          =   195
         Left            =   4050
         TabIndex        =   16
         Top             =   90
         Width           =   330
      End
      Begin VB.Label lblView 
         AutoSize        =   -1  'True
         Caption         =   "View:"
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   90
         Width           =   390
      End
   End
   Begin VB.PictureBox picScroll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5595
      ScaleWidth      =   8430
      TabIndex        =   2
      Top             =   465
      Width           =   8490
      Begin VB.PictureBox picProgress 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   3360
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   225
         TabIndex        =   20
         Top             =   360
         Width           =   3375
         Begin VB.Image imgPage 
            Height          =   990
            Index           =   3
            Left            =   120
            Picture         =   "frmPreview.frx":2CCE
            Top             =   30
            Width           =   720
         End
         Begin VB.Image imgPage 
            Height          =   990
            Index           =   2
            Left            =   120
            Picture         =   "frmPreview.frx":5230
            Top             =   30
            Width           =   720
         End
         Begin VB.Image imgPage 
            Height          =   990
            Index           =   1
            Left            =   120
            Picture         =   "frmPreview.frx":7792
            Top             =   30
            Width           =   720
         End
         Begin VB.Label lblProgTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2400
            TabIndex        =   22
            Top             =   375
            Width           =   765
         End
         Begin VB.Label lblProgPage 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Page:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   960
            TabIndex        =   21
            Top             =   360
            Width           =   720
         End
         Begin VB.Image imgPage 
            Height          =   990
            Index           =   0
            Left            =   120
            Picture         =   "frmPreview.frx":9CF4
            Top             =   30
            Width           =   720
         End
      End
      Begin VB.VScrollBar vsPreview 
         Height          =   1215
         Left            =   3000
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
      End
      Begin VB.HScrollBar hsPreview 
         Height          =   255
         Left            =   720
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1725
      End
      Begin VB.PictureBox picShow 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   240
         ScaleHeight     =   3615
         ScaleWidth      =   3540
         TabIndex        =   3
         Top             =   480
         Width           =   3540
         Begin VB.PictureBox picHold 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   240
            ScaleHeight     =   1815
            ScaleWidth      =   2175
            TabIndex        =   5
            Top             =   120
            Width           =   2175
            Begin VB.PictureBox picDoc 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   1215
               Left            =   240
               ScaleHeight     =   1215
               ScaleWidth      =   1695
               TabIndex        =   6
               Top             =   240
               Visible         =   0   'False
               Width           =   1695
            End
         End
         Begin VB.PictureBox picNormal 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   480
            ScaleHeight     =   1215
            ScaleWidth      =   1695
            TabIndex        =   4
            TabStop         =   0   'False
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
' Author:       Edward Moth
' Copyright:    © 2000-2002 qbd software ltd
'
' ==============================================================
' Module:       frmPreview
' Purpose:      Display Print Preview
' ==============================================================
Option Explicit
Private bPreview As Boolean
Private qPreview As qcPrinter
Private bScrollCode As Boolean
Private sZoom As Single
Private lPage As Integer
Private lPageMax As Integer
Private bDisplayPage As Boolean
Private mintLanguageOffset As String
Private mblnLanguageReverse As Boolean
Private mintCharset As Integer
Private msngWidth As Single, msngHeight As Single, msngTop As Single, msngLeft As Single
Private bShown As Boolean

Public Property Get WindowWidth() As Single
  WindowWidth = msngWidth
End Property

Public Property Get WindowHeight() As Single
  WindowHeight = msngHeight
End Property

Public Property Get WindowLeft() As Single
  WindowLeft = msngLeft
End Property

Public Property Get WindowTop() As Single
  WindowTop = msngTop
End Property

Public Property Get CurrentPage() As Integer
  CurrentPage = lPage
End Property

Public Property Let CurrentPage(ByVal nNewPage As Integer)
  lPage = nNewPage
End Property

Public Property Set Document(ByVal vNewValue As qcPrinter)
  Set qPreview = vNewValue
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

Private Sub cboZoom_KeyPress(KeyAscii As Integer)
  Dim sNewZoom As Single

  If KeyAscii = 13 Then
    sNewZoom = Val(cboZoom.Text)

    If sNewZoom > 0 And sNewZoom <= 200 Then
      cboZoom.Text = sNewZoom & " %"

      If sNewZoom = sZoom Then
        Exit Sub
      End If

      sZoom = sNewZoom
      Zoom_Check
    Else

      If cboZoom.ListIndex >= 0 Then
        cboZoom.Text = cboZoom.List(cboZoom.ListIndex)
      Else
        cboZoom.Text = sZoom & " %"
      End If

    End If

  End If

End Sub

Private Sub cmdClose_Click()
  Set qPreview = Nothing
  Me.Hide
End Sub

Private Sub cmdGo_Click()
  Dim vPage As Variant
  Dim lGo As Long
  vPage = txtPage.Text

  If IsNumeric(vPage) Then
    lGo = CLng(vPage)

    If lGo >= 1 And lGo <= lPageMax Then
      lPage = lGo
      Preview_Display lPage
    End If

  End If

End Sub

Private Sub cmdNext_Click()
  lPage = lPage + 1
  Preview_Display lPage
End Sub

Private Sub cmdPrevious_Click()
  lPage = lPage - 1
  Preview_Display lPage
End Sub

Private Sub cmdEnd_Click()
  lPage = lPageMax
  Preview_Display lPage
End Sub

Private Sub cmdSave_Click()
'  Dim strFile As String
'  On Error Resume Next
'
'  With cDialog
'    .Filter = qPreview.SaveFileDescription & "(*." & qPreview.SaveFileExtension & ")|*." & qPreview.SaveFileExtension & "|All Files (*.*)|*.*"
'    .DefaultExt = "*." & qPreview.SaveFileExtension
'    .DialogTitle = "Save " & qPreview.SaveFileDescription
'    .CancelError = True
'    .ShowSave
'
'    If Err.Number = cdlCancel Then
'      Exit Sub
'    End If
'
'    strFile = .Filename
'    Me.MousePointer = vbHourglass
'    qPreview.Document.Save strFile
'    Me.MousePointer = vbNormal
'  End With

End Sub

Private Sub cmdStart_Click()
  lPage = 1
  Preview_Display lPage
End Sub

Private Sub cmdPrint_Click()
  Dim lPrintStart As Integer
  Dim lPrintEnd As Integer
  Dim iCopies As Integer
  Dim bCollate As Boolean
  Dim iPrinter As Integer
  ' Display the Print Options form
  ' Pass current page details
  frmPrint.Flags = qPreview.PrintOptions
  frmPrint.SaveInfo qPreview.ShowSaveButton, qPreview.SaveFileExtension, qPreview.SaveFileDescription
  frmPrint.PageCurrent = lPage
  frmPrint.PageMax = lPageMax
  'Call frmPrint.SetCaptions(mintLanguageOffset, mblnLanguageReverse, mintCharset)

  Do
    frmPrint.SaveDoc = False
    frmPrint.Show vbModal

    If frmPrint.SaveDoc Then
      qPreview.Document.Save frmPrint.SaveFile
    End If

  Loop While frmPrint.SaveDoc

  lPrintStart = frmPrint.PageStart
  lPrintEnd = frmPrint.PageEnd
  iCopies = frmPrint.Copies
  bCollate = frmPrint.Collate
  iPrinter = frmPrint.PrinterNumber

  If frmPrint.PrintDoc Then
    lblTotal.Caption = "Printing..."
    lblTotal.Refresh
    qPreview.PrintDoc lPrintStart, lPrintEnd, iCopies, bCollate, iPrinter
    lblTotal.Caption = " / " & lPageMax
  End If

  Unload frmPrint
End Sub

Private Sub Form_Activate()
  Dim sTime As Single
  Me.Refresh
  bDisplayPage = True

  If lPage = 0 Then
    lPage = 1
  End If

  '    Preview_Display lPage
  Me.MousePointer = vbHourglass

  With picProgress
    .AutoRedraw = True
    picProgress.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), vbButtonShadow, B
    picProgress.Line (0, 0)-(.ScaleWidth - 1, 0), vb3DHighlight
    picProgress.Line (0, 0)-(0, .ScaleHeight - 1), vb3DHighlight
    .Visible = True
    .AutoRedraw = False
  End With

  DoEvents
  sTime = Timer
  tmrProgress.Enabled = True
  lPageMax = qPreview.Document.Pages
  picProgress.Visible = False
  tmrProgress.Enabled = False
  Preview_Status
  sTime = Timer - sTime
  Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  bDisplayPage = True

  If lPage = 0 Then
    lPage = 1
  End If

    Util.SetNumber (txtPage.hWnd)
  sZoom = 100
  Preview_Display lPage
End Sub

Public Sub Preview_Display(ByVal iPage As Integer)
  Dim iMin As Integer
  Dim iMax As Integer
  Screen.MousePointer = vbHourglass
  picNormal.Cls
  picDoc.Visible = False
  qPreview.PreviewPage picNormal, iPage
  Preview_Status
  Zoom_Check True
  Screen.MousePointer = vbDefault
End Sub

Private Sub Zoom_Check(Optional ByVal ForcePaint As Boolean = False)
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

  If (sSizeX <> picDoc.Width And sSizeY <> picDoc.Height) Or ForcePaint Then
    picDoc.Move 240, 240, sSizeX, sSizeY
    picDoc.Cls
    picDoc.PaintPicture picNormal.Image, 0, 0, sSizeX, sSizeY
  End If

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

  If bDisplayPage Then
    picDoc.Visible = True
  End If

End Sub

Private Sub Form_Resize()

  If Me.WindowState = vbMinimized Then
    Exit Sub
  End If

  If Me.ScaleHeight > 600 Then
    picScroll.Move 0, 435, Me.ScaleWidth, Me.ScaleHeight - 435
  End If

  If Me.WindowState = vbNormal And Me.Visible Then
    msngWidth = Me.Width
    msngHeight = Me.Height
    msngTop = Me.Top
    msngLeft = Me.Left
  End If

  picControl.Width = Me.ScaleWidth
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If UnloadMode <> vbFormCode Then
    Me.Hide
    Cancel = 1
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  tmrProgress.Enabled = False
  bDisplayPage = False
End Sub

Private Sub hsPreview_Change()

  If Not bScrollCode Then
    picHold.Left = -hsPreview.Value * 14.4
  End If

End Sub

Private Sub picDoc_Click()
  picScroll.SetFocus
End Sub

Private Sub picDoc_GotFocus()
  picScroll.SetFocus
End Sub

Private Sub picScroll_Resize()
  vsPreview.Move picScroll.ScaleWidth - vsPreview.Width, 0, 255, picScroll.ScaleHeight
  hsPreview.Move 0, picScroll.ScaleHeight - hsPreview.Height, picScroll.ScaleWidth
  Zoom_Check
  picProgress.Move (picDoc.Width - picProgress.Width) / 2, 240
End Sub

Private Sub tmrProgress_Timer()
  Static iImage As Integer
  '#BUGFIX FABIAN ;)

  If qPreview Is Nothing Then
    tmrProgress.Enabled = False
    Exit Sub
  End If

  '#END BUGFIX
  lblTotal.Caption = "/ " & qPreview.Document.PageProgress
  lblProgTotal = qPreview.Document.PageProgress
  imgPage(iImage).ZOrder 0
  iImage = (iImage + 1) Mod 4
End Sub

Private Sub txtPage_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
    cmdGo_Click
    KeyAscii = 0
  End If

End Sub

Private Sub vsPreview_Change()

  If Not bScrollCode Then
    picHold.Top = -vsPreview.Value * 14.4
  End If

End Sub

Public Sub Preview_Status()
  cmdStart.Enabled = CBool(lPage > 1)
  cmdPrevious.Enabled = CBool(lPage > 1)
  cmdNext.Enabled = CBool(lPage < lPageMax)
  cmdEnd.Enabled = CBool(lPage < lPageMax)
  lblTotal.Caption = " / " & lPageMax
  lblTotal.Visible = True
  txtPage.Text = lPage
  txtPage.SelStart = 0
  txtPage.SelLength = 5
  txtPage.Enabled = CBool(lPageMax > 1)
  cmdGo.Enabled = CBool(lPageMax > 1)
End Sub

Public Sub SetCaptions(ByVal intLanguageOffset As Integer, ByVal blnLanguageRev As Boolean, ByVal intCharset As Integer)
  Dim sComboSize As Single
  mintLanguageOffset = intLanguageOffset
  mblnLanguageReverse = blnLanguageRev
  mintCharset = intCharset
  Me.Font.Charset = intCharset
  'Me.Caption = LoadResString(intLanguageOffset)
  Me.cmdClose.Font.Charset = intCharset
  'Me.cmdClose.Caption = LoadResString(intLanguageOffset + 1)
  Me.cmdPrint.Font.Charset = intCharset
  'Me.cmdPrint.Caption = LoadResString(intLanguageOffset + 2)
  Me.cmdGo.Font.Charset = intCharset
  'Me.cmdGo.Caption = LoadResString(intLanguageOffset + 3)
  Me.lblStatus.Font.Charset = intCharset
  'Me.lblStatus.Caption = LoadResString(intLanguageOffset + 4)
  Me.lblProgPage.Font.Charset = intCharset
  'Me.lblProgPage.Caption = LoadResString(intLanguageOffset + 4)
  Me.lblView.Font.Charset = intCharset
  'Me.lblView.Caption = LoadResString(intLanguageOffset + 5)
  cboZoom.Font.Charset = intCharset
  sZoom = 100

  With cboZoom
    .Clear
    .AddItem "100 %"
    .ItemData(.ListCount - 1) = 100
    .AddItem "75 %"
    .ItemData(.ListCount - 1) = 75
    .AddItem "50 %"
    .ItemData(.ListCount - 1) = 50
    .AddItem "Page Completed" 'LoadResString(intLanguageOffset + 6) ' Full Page
    .ItemData(.ListCount - 1) = 0
    .AddItem "Page Width" 'LoadResString(intLanguageOffset + 7) ' Full Width
    .ItemData(.ListCount - 1) = -1
    bScrollCode = True
    .ListIndex = 0
    bScrollCode = False
    'Check combobox width

    If picControl.TextWidth(.List(4)) > picControl.TextWidth(.List(5)) Then
      sComboSize = picControl.TextWidth(.List(4)) + 360
    Else
      sComboSize = picControl.TextWidth(.List(5)) + 360
    End If

    If sComboSize < 1455 Then
      sComboSize = 1455
    End If

    cboZoom.Width = sComboSize
  End With

  sZoom = 100
  txtPage.Font.Charset = intCharset
  picControl.Font.Charset = intCharset
  Controls_Size
End Sub

Private Sub Controls_Size()
  Dim sCurX As Single, sWidth As Single

  If mblnLanguageReverse Then
    sCurX = 120
    Control_MoveItem cmdClose, sCurX, 120, 855

    'If qPreview.ShowSaveButton Then
    '  Control_MoveItem cmdSave, sCurX, 120, 315
    'End If

    'cmdSave.Visible = qPreview.ShowSaveButton
    Control_MoveItem cmdPrint, sCurX, 120, 855
    Control_MoveItem cmdGo, sCurX, 60, 315
    lblTotal.Caption = "/ 9999 "
    Control_MoveItem lblTotal, sCurX, 60
    lblTotal.Caption = "/ 0"
    Control_MoveItem txtPage, sCurX, 60
    Control_MoveItem lblStatus, sCurX, 120
    Control_MoveItem cmdStart, sCurX, 60
    Control_MoveItem cmdPrevious, sCurX, 60
    Control_MoveItem cmdNext, sCurX, 60
    Control_MoveItem cmdEnd, sCurX, 120
    Control_MoveItem cboZoom, sCurX, 120
    Control_MoveItem lblView, sCurX, 0
    lblTotal.Alignment = 1
  Else
    sCurX = lblView.Left + lblView.Width + 120
    Control_MoveItem cboZoom, sCurX, 120
    Control_MoveItem cmdStart, sCurX, 60
    Control_MoveItem cmdPrevious, sCurX, 60
    Control_MoveItem cmdNext, sCurX, 60
    Control_MoveItem cmdEnd, sCurX, 120
    Control_MoveItem lblStatus, sCurX, 60
    Control_MoveItem txtPage, sCurX, 60
    lblTotal.Caption = "/ 9999 "
    Control_MoveItem lblTotal, sCurX, 60
    lblTotal.Caption = "/ 0"
    Control_MoveItem cmdGo, sCurX, 240, 315

    If qPreview.ShowSaveButton Then
    '  Control_MoveItem cmdSave, sCurX, 120, 315
    End If

    'cmdSave.Visible = qPreview.ShowSaveButton
    Control_MoveItem cmdPrint, sCurX, 120, 855
    Control_MoveItem cmdClose, sCurX, 0, 855
    lblTotal.Alignment = 0
  End If

End Sub

Private Sub Control_MoveItem(ByRef ctlItem As Control, _
                             ByRef sCurrentPos As Single, _
                             ByVal sGap As Single, _
                             Optional sAutoMinWidth As Single = 0)
  Dim sWidth As Single
  Dim sStrip As String
  ctlItem.Left = sCurrentPos

  If sAutoMinWidth > 0 Then
    sStrip = String_StripAmpersand(ctlItem.Caption)
    sWidth = picControl.TextWidth(sStrip) + 240

    If sWidth < sAutoMinWidth Then
      sWidth = sAutoMinWidth
    End If

    ctlItem.Width = sWidth
  Else
    sWidth = ctlItem.Width
  End If

  sCurrentPos = sCurrentPos + sWidth + sGap
End Sub

Private Function String_StripAmpersand(ByVal sText As String) As String
  String_StripAmpersand = Replace(sText, "&", "")
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim vsStep As Integer
  Dim hsStep As Integer
  vsStep = vsPreview.Max / 3
  hsStep = hsPreview.Max / 3

  If (KeyCode = vbKeyPageDown) Or (KeyCode = vbKeyDown) Then

    If (vsPreview.Value + vsStep) > vsPreview.Max Then
      vsPreview.Value = vsPreview.Max
    Else
      vsPreview.Value = vsPreview.Value + vsStep
    End If

    picHold.Top = -vsPreview.Value * 14.4
  End If

  If (KeyCode = vbKeyPageUp) Or (KeyCode = vbKeyUp) Then

    If (vsPreview.Value - vsStep) < vsPreview.Min Then
      vsPreview.Value = vsPreview.Min
    Else
      vsPreview.Value = vsPreview.Value - vsStep
    End If

    picHold.Top = -vsPreview.Value * 14.4
  End If

  If (KeyCode = vbKeyRight) Then

    If (hsPreview.Value + hsStep) > hsPreview.Max Then
      hsPreview.Value = hsPreview.Max
    Else
      hsPreview.Value = hsPreview.Value + hsStep
    End If

    picHold.Left = -hsPreview.Value * 14.4
  End If

  If (KeyCode = vbKeyLeft) Then

    If (hsPreview.Value - hsStep) < hsPreview.Min Then
      hsPreview.Value = hsPreview.Min
    Else
      hsPreview.Value = hsPreview.Value - hsStep
    End If

    picHold.Left = -hsPreview.Value * 14.4
  End If

End Sub

