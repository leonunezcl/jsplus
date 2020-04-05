VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmPageSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "No Printers Installed"
   ClientHeight    =   5055
   ClientLeft      =   4500
   ClientTop       =   2610
   ClientWidth     =   5385
   Icon            =   "PageSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrinter 
      Caption         =   "Printer..."
      Height          =   375
      Left            =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4590
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4275
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4590
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3330
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4590
      Width           =   900
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   4455
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Measurements"
      TabPicture(0)   =   "PageSetup.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSize"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "picShadow"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "picThumb"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Document"
      TabPicture(1)   =   "PageSetup.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFooter"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtHeader"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label6"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.TextBox txtFooter 
         Height          =   315
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "PageSetup.frx":05C2
         Top             =   1550
         Width           =   4695
      End
      Begin VB.TextBox txtHeader 
         Height          =   315
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "PageSetup.frx":05DB
         Top             =   900
         Width           =   4695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Margins (millimeters)"
         Height          =   1335
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   3255
         Begin VB.TextBox txtLeftMargin 
            Height          =   315
            Left            =   720
            TabIndex        =   12
            Text            =   "25"
            Top             =   345
            Width           =   375
         End
         Begin VB.TextBox txtRightMargin 
            Height          =   315
            Left            =   720
            TabIndex        =   11
            Text            =   "25"
            Top             =   810
            Width           =   375
         End
         Begin VB.TextBox txtTopMargin 
            Height          =   315
            Left            =   2280
            TabIndex        =   10
            Text            =   "25"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtBottomMargin 
            Height          =   315
            Left            =   2280
            TabIndex        =   9
            Text            =   "25"
            Top             =   810
            Width           =   375
         End
         Begin VB.Label lblBottomMargin 
            Caption         =   "Bottom:"
            Height          =   285
            Left            =   1680
            TabIndex        =   16
            Top             =   840
            Width           =   840
         End
         Begin VB.Label lblTopMargin 
            Caption         =   "Top:"
            Height          =   195
            Left            =   1680
            TabIndex        =   15
            Top             =   405
            Width           =   615
         End
         Begin VB.Label lblRightMargin 
            Caption         =   "Right:"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   855
            Width           =   570
         End
         Begin VB.Label lblLeftMargin 
            Caption         =   "Left:"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   405
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Orientation"
         Height          =   1335
         Left            =   3480
         TabIndex        =   5
         Top             =   3000
         Width           =   1575
         Begin VB.OptionButton optPortrait 
            Caption         =   "P&ortrait"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optLandscape 
            Caption         =   "L&andscape"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.PictureBox picThumb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   1920
         ScaleHeight     =   1935
         ScaleWidth      =   1335
         TabIndex        =   4
         Top             =   840
         Width           =   1335
         Begin VB.Shape shpMarg 
            BorderColor     =   &H00808080&
            BorderStyle     =   3  'Dot
            Height          =   1695
            Left            =   120
            Top             =   120
            Width           =   1095
         End
         Begin VB.Image imgText 
            Height          =   1950
            Left            =   180
            Picture         =   "PageSetup.frx":05E0
            Top             =   180
            Width           =   1950
         End
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   2040
         ScaleHeight     =   1935
         ScaleWidth      =   1335
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "\n: Carriage return character"
         Height          =   255
         Left            =   -74775
         TabIndex        =   27
         Top             =   4050
         Width           =   4215
      End
      Begin VB.Label lblSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Size"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   4680
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "^N: Document Name"
         Height          =   255
         Left            =   -74775
         TabIndex        =   25
         Top             =   2790
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Page Footer"
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Page Header"
         Height          =   255
         Left            =   -74760
         TabIndex        =   23
         Top             =   660
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "^P: Document Name and Path"
         Height          =   255
         Left            =   -74775
         TabIndex        =   22
         Top             =   3030
         Width           =   4215
      End
      Begin VB.Label Label5 
         Caption         =   "^D: Today's Date (Default format)"
         Height          =   255
         Left            =   -74775
         TabIndex        =   21
         Top             =   3270
         Width           =   4215
      End
      Begin VB.Label Label6 
         Caption         =   "^T: Current Time (Default format)"
         Height          =   255
         Left            =   -74775
         TabIndex        =   20
         Top             =   3510
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmPageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pWidth As Integer   ' current printer paper width
Public pHeight As Integer  ' current printer paper height

Private Sub cmdPrinter_Click()

'Purpose:   To allow the user to preset printer defaults (plus things like orientation)
    
    On Error GoTo errHandler

    With frmMain.CD
        .PrinterDefault = True 'tells the system to set all printer driver properties
        .CancelError = True
        If Me.optLandscape = True Then
            .Orientation = cdlLandscape
        Else
            .Orientation = cdlPortrait
        End If
        .Flags = cdlPDPrintSetup Or cdlPDReturnDC
        .ShowPrinter
    End With
    ' Setting papersize
    pPaperSize = Printer.papersize
    lblSize.Caption = getPrintSize(pPaperSize)
    
    ' Setting printer orientation
    Printer.Orientation = frmMain.CD.Orientation
    
    ' Setting paperorientation
    If Printer.Orientation = 1 Then     ' portrait
        optPortrait.Value = True
    ElseIf Printer.Orientation = 2 Then ' landscape
        optLandscape.Value = True
    End If
    
    ' Setting textboxmargins from registry
    txtLeftMargin.Text = gLeftMargin
    txtRightMargin.Text = gRightMargin
    txtTopMargin.Text = gTopMargin
    txtBottomMargin.Text = gBottomMargin

    ' setting printer paper height in mm
    pHeight = Round(Printer.Height / 57)
    ' setting printer paper width in mm
    pWidth = Round(Printer.Width / 57)
    Call makeThumb("Left")
    Call makeThumb("Right")
    Call makeThumb("Top")
    Call makeThumb("Bottom")

    Exit Sub
errHandler:
    Exit Sub

End Sub

Private Sub Form_Load()
'SetFont Me
    On Error GoTo ErrorHandler
    
    util.CenterForm Me
    
    util.SetNumber txtLeftMargin.hWnd
    util.SetNumber txtRightMargin.hWnd
    util.SetNumber txtTopMargin.hWnd
    util.SetNumber txtBottomMargin.hWnd
    
    pOrientation = Printer.Orientation  ' 1 = Portrait, 2 = Landscape
    pPaperSize = Printer.papersize
    If pPaperSize = 0 Then  ' No printer installed
        MsgBox "No printer found!", vbCritical
        lblSize.Caption = "No printer found..."
        imgText.Visible = False
        ' disabling printer spesific items
        Me.optLandscape.Enabled = False
        Me.optPortrait.Enabled = False
        Me.cmdPrinter.Enabled = False
    Else
        ' Setting form caption to printername
        Me.Caption = Printer.DeviceName '"Page Setup - " & Printer.DeviceName
        lblSize.Caption = getPrintSize(pPaperSize)
        ' setting printer paper height in mm
        pHeight = Round(Printer.Height / 57)
        ' setting printer paper width in mm
        pWidth = Round(Printer.Width / 57)
    End If
    
    If pOrientation = 2 Then        ' Landscape
        optLandscape.Value = True
        picThumb.Width = 1935
        picThumb.Height = 1335
        picShadow.Width = 1935
        picShadow.Height = 1335
    ElseIf pOrientation = 1 Then    ' Portrait
        optPortrait.Value = True
        picThumb.Width = 1335
        picThumb.Height = 1935
        picShadow.Width = 1335
        picShadow.Height = 1935
    End If
    ' Centring picThumb
    picThumb.Left = lblSize.Left + (lblSize.Width / 2) - (picThumb.Width / 2)
    picShadow.Left = lblSize.Left + (lblSize.Width / 2) - (picThumb.Width / 2) + 110
 
    txtLeftMargin.Text = gLeftMargin
    txtRightMargin.Text = gRightMargin
    txtTopMargin.Text = gTopMargin
    txtBottomMargin.Text = gBottomMargin
    Call makeThumb("Left")
    Call makeThumb("Right")
    Call makeThumb("Top")
    Call makeThumb("Bottom")

ErrorHandler:
    Exit Sub

End Sub

Private Sub cmdOK_Click()
    sPrintHeader = txtHeader.Text
    sPrintFooter = txtFooter.Text
        
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPageSetup = Nothing
End Sub


Private Sub optLandscape_Click()
    On Error GoTo ErrorHandler

    picThumb.Width = 1935
    picThumb.Height = 1335
    picShadow.Width = 1935
    picShadow.Height = 1335
    picThumb.Left = lblSize.Left + (lblSize.Width / 2) - (picThumb.Width / 2)
    picShadow.Left = lblSize.Left + (lblSize.Width / 2) - (picThumb.Width / 2) + 110
    
    ' Set printer orientation
    Printer.Orientation = 2  ' 1 = Portrait, 2 = Landscape
    pOrientation = 2
    ' Read papersize
    pPaperSize = Printer.papersize
    
    If pPaperSize = 0 Then  ' No printer installed
        ' nothing
    Else
        lblSize.Caption = getPrintSize(pPaperSize)
        ' setting printer paper height in mm
        pHeight = Round(Printer.Height / 57)
        ' setting printer paper width in mm
        pWidth = Round(Printer.Width / 57)
    End If
    
    Call makeThumb("Left")
    Call makeThumb("Right")
    Call makeThumb("Top")
    Call makeThumb("Bottom")

ErrorHandler:
    Exit Sub

End Sub

Private Sub optPortrait_Click()
    On Error GoTo ErrorHandler
    
    picThumb.Width = 1335
    picThumb.Height = 1935
    picShadow.Width = 1335
    picShadow.Height = 1935
    picThumb.Left = lblSize.Left + (lblSize.Width / 2) - (picThumb.Width / 2)
    picShadow.Left = lblSize.Left + (lblSize.Width / 2) - (picThumb.Width / 2) + 110
    
    
    ' Set printer orientation
    Printer.Orientation = 1  ' 1 = Portrait, 2 = Landscape
    pOrientation = 1
    ' Read papersize
    pPaperSize = Printer.papersize

    If pPaperSize = 0 Then  ' No printer installed
        ' nothing
    Else
        lblSize.Caption = getPrintSize(pPaperSize)
        ' setting printer paper height in mm
        pHeight = Round(Printer.Height / 57)
        ' setting printer paper width in mm
        pWidth = Round(Printer.Width / 57)
    End If

    Call makeThumb("Left")
    Call makeThumb("Right")
    Call makeThumb("Top")
    Call makeThumb("Bottom")

ErrorHandler:
    Exit Sub

End Sub

Private Sub txtBottomMargin_Change()
    Call makeThumb("Bottom")
End Sub

Private Sub txtBottomMargin_KeyPress(KeyAscii As Integer)
    ' making shure only text or del or backspace is pressed
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 127 Or KeyAscii = 8 Then
        ' nothing
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtFooter_LostFocus()
If txtFooter.Text = "" Then Exit Sub
If Left(txtFooter.Text, 4) <> "\n\n" Then txtFooter.Text = "\n\n" & txtFooter.Text
End Sub

Private Sub txtHeader_LostFocus()
If txtHeader.Text = "" Then Exit Sub
If Right(txtHeader.Text, 4) <> "\n\n" Then txtHeader.Text = txtHeader.Text & "\n\n"
End Sub

Private Sub txtLeftMargin_Change()
    Call makeThumb("Left")
End Sub

Private Sub txtLeftMargin_KeyPress(KeyAscii As Integer)
    ' making shure only text or del or backspace is pressed
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 127 Or KeyAscii = 8 Then
        ' nothing
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtRightMargin_Change()
    Call makeThumb("Right")
End Sub

Private Sub txtRightMargin_KeyPress(KeyAscii As Integer)
    ' making shure only text or del or backspace is pressed
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 127 Or KeyAscii = 8 Then
        ' nothing
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtTopMargin_Change()
    Call makeThumb("Top")
End Sub

Private Sub txtTopMargin_KeyPress(KeyAscii As Integer)
    ' making shure only text or del or backspace is pressed
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 127 Or KeyAscii = 8 Then
        ' nothing
    Else
        KeyAscii = 0
    End If
End Sub

Public Sub makeThumb(marg As String)
On Error Resume Next
    ' this sub calculates how many percent the margins selected accounts for on the selected
    ' printers selected papertype.
    ' It then draws a thumb of the paper with margins the way it will look when printed.
    
    Dim lMarg As Integer    ' the actual left margin in mm
    Dim rMarg As Integer    ' the actual right margin in mm
    Dim tMarg As Integer    ' the actual top margin in mm
    Dim bMarg As Integer    ' the actual bottom margin in mm

    Dim lPercent As Integer ' the left margin in percentage of actual papersize
    Dim rPercent As Integer ' the right margin in percentage of actual papersize
    Dim tPercent As Integer ' the top margin in percentage of actual papersize
    Dim bPercent As Integer ' the bottom margin in percentage of actual papersize
    
    ' Making shure the margins are set to "0" if any of the textboxes are empty
    If txtLeftMargin.Text = "" Then
        lMarg = 0
    Else
        lMarg = Int(txtLeftMargin.Text)
    End If
    If txtRightMargin.Text = "" Then
        rMarg = 0
    Else
        rMarg = Int(txtRightMargin.Text)
    End If
    If txtTopMargin.Text = "" Then
        tMarg = 0
    Else
        tMarg = Int(txtTopMargin.Text)
    End If
    If txtBottomMargin.Text = "" Then
        bMarg = 0
    Else
        bMarg = Int(txtBottomMargin.Text)
    End If
    
    ' Setting the percentages from margins
    lPercent = Round(lMarg * 100 / pWidth)
    rPercent = Round(rMarg * 100 / pWidth)
    tPercent = Round(tMarg * 100 / pHeight)
    bPercent = Round(bMarg * 100 / pHeight)
    
    ' Adjusting the rectagle on the paperthumb
    shpMarg.Left = (picThumb.Width / 100) * (lPercent)
    shpMarg.Width = picThumb.Width - shpMarg.Left - (picThumb.Width / 100 * rPercent)
    shpMarg.Top = (picThumb.Height / 100) * tPercent
    shpMarg.Height = picThumb.Height - shpMarg.Top - (picThumb.Height / 100 * bPercent)
    ' Adjusting the text on the paperthumb
    imgText.Left = shpMarg.Left
    imgText.Top = shpMarg.Top
    imgText.Width = shpMarg.Width
    imgText.Height = shpMarg.Height

    ' Checking the size of the left/right and top/bottom margins
    Select Case marg
        Case "Left"
            If lMarg > Round(pWidth / 100 * 70) / 2 Then    ' 35 % of page width
                MsgBox marg & " margin bigger than 35% of page width or overlapping! Adjusting.", vbInformation
                txtLeftMargin.Text = Round(pWidth / 100 * 70) / 2
                gLeftMargin = txtLeftMargin.Text
                Exit Sub
            Else
                ' if the textbox is empty set it to "0"
                If txtLeftMargin.Text = "" Then
                    gLeftMargin = 0
                Else
                    gLeftMargin = txtLeftMargin.Text
                End If
            End If
        Case "Right"
            If rMarg > Round(pWidth / 100 * 70) / 2 Then    ' 35 % of page width
                MsgBox marg & " margin bigger than 35% of page width or overlapping! Adjusting.", vbInformation
                txtRightMargin.Text = Round(pWidth / 100 * 70) / 2
                gRightMargin = txtRightMargin.Text
                Exit Sub
            Else
                ' if the textbox is empty set it to "0"
                If txtRightMargin.Text = "" Then
                    gRightMargin = 0
                Else
                    gRightMargin = txtRightMargin.Text
                End If
            End If
        Case "Top"
            If tMarg > Round(pHeight / 100 * 70) / 2 Then    ' 35 % of page width
                MsgBox marg & " margin bigger than 35% of page height or overlapping! Adjusting.", vbInformation
                txtTopMargin.Text = Round(pHeight / 100 * 70) / 2
                gTopMargin = txtTopMargin.Text
                Exit Sub
            Else
                ' if the textbox is empty set it to "0"
                If txtTopMargin.Text = "" Then
                    gTopMargin = 0
                Else
                    gTopMargin = txtTopMargin.Text
                End If
            End If
        Case "Bottom"
            If bMarg > Round(pHeight / 100 * 70) / 2 Then    ' 35 % of page width
                MsgBox marg & " margin bigger than 35% of page height or overlapping! Adjusting.", vbInformation
                txtBottomMargin.Text = Round(pHeight / 100 * 70) / 2
                gBottomMargin = txtBottomMargin.Text
                Exit Sub
            Else
                ' if the textbox is empty set it to "0"
                If txtBottomMargin.Text = "" Then
                    gBottomMargin = 0
                Else
                    gBottomMargin = txtBottomMargin.Text
                End If
            End If
    End Select
End Sub
