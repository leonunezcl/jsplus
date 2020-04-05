VERSION 5.00
Begin VB.Form frmPrinting 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Printing ...."
   ClientHeight    =   1620
   ClientLeft      =   4560
   ClientTop       =   4605
   ClientWidth     =   3120
   Icon            =   "frmPrinting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrCancel 
      Interval        =   20
      Left            =   1200
      Top             =   1080
   End
   Begin VB.Timer tmrPrint 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   1080
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.Label lblPrinting 
         BackStyle       =   0  'Transparent
         Caption         =   "Preparing ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   3435
      End
      Begin VB.Label lblCopyTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "/ 1"
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
         Left            =   1800
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblPropCopy 
         BackStyle       =   0  'Transparent
         Caption         =   "Copy:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   915
      End
      Begin VB.Label lblCopy 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label lblPageTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "/ 1"
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
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblProgPage 
         BackStyle       =   0  'Transparent
         Caption         =   "Page:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   780
      End
      Begin VB.Label lblPage 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private qPrint As qcPrinter
Private mblnCancel As Boolean
Private mintStartPage As Integer
Private mintEndPage As Integer
Private mintCopies As Integer
Private mblnCollate As Boolean
Private meDuplex As enumDuplexPrintOptions
Private mintPage As Integer
Private mintCopy As Integer
Private mintPageAdd As Integer
Private mblnPrinting As Boolean

Public Property Set qPrintObject(ByVal vNewValue As qcPrinter)
  Set qPrint = vNewValue
End Property

Public Sub PrintInformation(ByVal StartPage As Integer, _
                            ByVal EndPage As Integer, _
                            ByVal Copies As Integer, _
                            ByVal Collate As Boolean, _
                            ByVal DuplexInfo As enumDuplexPrintOptions)
  mintStartPage = StartPage
  mintEndPage = EndPage
  mintCopies = Copies
  mblnCollate = Collate
  meDuplex = DuplexInfo
End Sub

Private Sub PrintDocument()
  ' Procedure sets the Target in PrintText to Printer
  ' printing the pages defined by StartPage and EndPage
  ' Check current printer
  Dim iStart As Integer
  Dim iEnd As Integer
  'Dim iCopy As Integer
  Dim iPage As Integer
  'Dim bPrint As Boolean
  lblPrinting.Caption = "Preparing Document ..."
  DoEvents
  iPage = qPrint.Document.Pages

  If mintStartPage = 0 Then
    iStart = 1
  Else
    iStart = mintStartPage
  End If

  If mintEndPage = 0 Then
    iEnd = iPage
  Else
    iEnd = mintEndPage

    If iEnd > iPage Then
      iEnd = iPage
    End If

  End If

  If iStart > iEnd Then
    iEnd = iStart
  End If

  Select Case meDuplex
    Case enumDuplexPrintOptions.duplexOdd

      If iStart Mod 2 = 0 Then
        iStart = iStart + 1
      End If

      If iEnd Mod 2 = 0 Then
        iEnd = iEnd - 1
      End If

      If iStart > iEnd Then
        ' No Odd pages
        GoTo Print_cancel
      End If

      mintPageAdd = 2
    Case enumDuplexPrintOptions.duplexEven

      If iStart Mod 2 = 1 Then
        iStart = iStart + 1
      End If

      If iEnd Mod 2 = 1 Then
        iEnd = iEnd - 1
      End If

      If iStart > iEnd Then
        ' No Even pages
        GoTo Print_cancel
      End If

      mintPageAdd = 2
    Case enumDuplexPrintOptions.duplexAll
      mintPageAdd = 1
  End Select

  If mintCopies < 1 Then
    mintCopies = 1
  End If

  lblCopyTotal.Caption = "/ " & mintCopies
  lblPageTotal.Caption = "/ " & iEnd
  lblCopy.Caption = "1"
  lblPage.Caption = iStart
  lblPrinting.Caption = "Printing Document ..."
  lblPage.Visible = True
  lblPageTotal.Visible = True
  lblCopy.Visible = True
  lblCopyTotal.Visible = True
  DoEvents
  Me.MousePointer = vbArrowHourglass
  Printer.DrawMode = vbCopyPen
  mintCopy = 1
  mintStartPage = iStart
  mintEndPage = iEnd
  mintPage = iStart
  tmrPrint.Enabled = True
  Exit Sub
Print_cancel:
  KillDocument
End Sub

Private Sub KillDocument()
  Printer.KillDoc
  tmrPrint.Enabled = False
  tmrCancel.Enabled = False
  Me.Hide
End Sub



Private Sub cmdClose_Click()
  mblnCancel = True
End Sub

Private Sub Form_Activate()
  PrintDocument
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set qPrint = Nothing
End Sub

Private Sub tmrCancel_Timer()
  DoEvents

  If mblnCancel Then
    gblnCancelDocument = True
    KillDocument
    tmrCancel.Enabled = False
  End If

End Sub

Private Sub tmrPrint_Timer()
  'Dim intPageAdd As Integer
  Dim blnComplete As Boolean

  If mblnCancel = True Then
    KillDocument
    Exit Sub
  End If

  mblnPrinting = True
  tmrPrint.Enabled = False
  lblCopy.Caption = mintCopy
  lblPage.Caption = mintPage
  DoEvents
  Printer.Print " ";
  Printer.CurrentX = 0
  qPrint.PrintText Printer, mintPage, mintPage

  If mblnCollate Then
    mintPage = mintPage + mintPageAdd

    If mintPage > mintEndPage Then
      mintPage = mintStartPage
      mintCopy = mintCopy + 1

      If mintCopy > mintCopies Then
        blnComplete = True
      End If

    End If

  Else
    mintCopy = mintCopy + 1

    If mintCopy > mintCopies Then
      mintPage = mintPage + mintPageAdd

      If mintPage > mintEndPage Then
        blnComplete = True
      End If

    End If

  End If

  If Not blnComplete Then
    DoEvents

    If mblnCancel Then
      KillDocument
    End If

    Printer.NewPage
    DoEvents
    tmrPrint.Enabled = True
  End If

  If Not mblnCancel And blnComplete Then
    tmrPrint.Enabled = False
    Printer.EndDoc
    Me.Hide
  End If

  If mblnCancel Then
    KillDocument
  End If

End Sub

