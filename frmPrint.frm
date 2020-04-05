VERSION 5.00
Begin VB.Form frmPrint 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Printing Setup"
   ClientHeight    =   4140
   ClientLeft      =   1380
   ClientTop       =   2460
   ClientWidth     =   5955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   315
      Left            =   3585
      TabIndex        =   19
      Top             =   3705
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   315
      Left            =   4680
      TabIndex        =   18
      Top             =   3705
      Width           =   975
   End
   Begin VB.Frame fraPrinter 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Printer"
      ForeColor       =   &H00FF8080&
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   5760
      Begin VB.CommandButton cmdProperties 
         Caption         =   "Printer Setup"
         Height          =   315
         Left            =   4425
         TabIndex        =   17
         Top             =   240
         Width           =   1200
      End
      Begin VB.ComboBox cboPrinter 
         Height          =   315
         Left            =   720
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lblName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   465
      End
   End
   Begin VB.Frame fraCopies 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copies"
      ForeColor       =   &H00FF8080&
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   5760
      Begin VB.CheckBox chkCollate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Collate"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4365
         TabIndex        =   12
         Top             =   780
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox txtCopies 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4365
         TabIndex        =   8
         Text            =   "1"
         Top             =   225
         Width           =   975
      End
      Begin VB.PictureBox picCopies 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   2055
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pages"
      ForeColor       =   &H00FF8080&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5760
      Begin VB.TextBox txtStart 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4380
         TabIndex        =   14
         Top             =   630
         Width           =   975
      End
      Begin VB.TextBox txtEnd 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4380
         TabIndex        =   11
         Top             =   990
         Width           =   975
      End
      Begin VB.OptionButton optPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Range"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2415
      End
      Begin VB.OptionButton optPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "All"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton optPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Current Page"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblEnd 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "End:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3120
         TabIndex        =   16
         Top             =   1020
         Width           =   330
      End
      Begin VB.Label lblStart 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Start:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3120
         TabIndex        =   15
         Top             =   660
         Width           =   420
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pág:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3120
         TabIndex        =   4
         Top             =   300
         Width           =   330
      End
      Begin VB.Label lblPages 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0 / 0"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4425
         TabIndex        =   3
         Top             =   300
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmPrint"
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
' Module:       frmPrint
' Purpose:      User options for printing
' ==============================================================
' Credits:      Thanks to Jarry of Jacsoft for his contributions
'               toward improving the functionality of this
'               module.
' ==============================================================
Option Explicit
Private mvarCurrent As Integer
Private mvarMax As Integer
Private mvarStart As Integer
Private mvarEnd As Integer
Private mvarPrint As Boolean
Private mvarCollate As Boolean
Private mvarFlags As qePrintOptionFlags
Private mblnShowSave As Boolean
Private mblnSave As Boolean
Private mstrSaveExt As String
Private mstrFileDescription As String
Private mstrSaveName As String
Private mvarPrinter As Integer
Private mvarCopies As Integer
Private bInternal As Boolean
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function PrinterProperties Lib "winspool.drv" (ByVal hWnd As Long, ByVal hPrinter As Long) As Long
Private Type PRINTER_DEFAULTS
  pDatatype As Long ' String
  pDevMode As Long
  pDesiredAccess As Long
End Type
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
                                    PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

Friend Property Let Flags(ByVal eFlag As qePrintOptionFlags)
  mvarFlags = eFlag
End Property

Friend Property Get SaveDoc() As Boolean
  SaveDoc = mblnSave
End Property

Friend Property Let SaveDoc(ByVal newSave As Boolean)
  mblnSave = False
End Property

Friend Property Get SaveFile() As String
  SaveFile = mstrSaveName
End Property

Friend Sub SaveInfo(ByVal ShowSave As Boolean, ByVal DefaultExtension As String, ByVal DefaultDescription As String)
  mblnShowSave = ShowSave
  mstrSaveExt = DefaultExtension
  mstrFileDescription = DefaultDescription
End Sub

Private Sub cboPrinter_Click()

  If Not bInternal And cboPrinter.ListIndex > -1 Then
    mvarPrinter = cboPrinter.ItemData(cboPrinter.ListIndex)
  End If

End Sub

Private Sub chkCollate_Click()
  mvarCollate = CBool(chkCollate.Value = vbChecked)
  Copies_ShowImage
End Sub

Private Sub Copies_ShowImage()
  Dim sX As Single
  Dim sY As Single
  Dim iPage As Integer
  Dim iNum As Integer
  picCopies.Cls
  picCopies.FontSize = 8

  If mvarCollate Then
    sX = 1400: sY = 0
    iNum = 2

    Do While iNum > 0
      iPage = 3

      Do While iPage > 0
        picCopies.Line (sX, sY)-Step(300, 420), vbWhite, BF
        picCopies.Line (sX, sY)-Step(300, 420), vbBlack, B
        picCopies.CurrentX = sX + 300 - picCopies.TextWidth(iPage) - 60
        picCopies.CurrentY = sY + 420 - picCopies.TextHeight(iPage)
        picCopies.Print iPage
        sX = sX - 150
        sY = sY + 210
        iPage = iPage - 1
      Loop

      sX = sX - 400
      sY = 0
      iNum = iNum - 1
    Loop

  Else
    sX = 1400: sY = 105
    iNum = 3

    Do While iNum > 0
      iPage = 2

      Do While iPage > 0
        picCopies.Line (sX, sY)-Step(300, 420), vbWhite, BF
        picCopies.Line (sX, sY)-Step(300, 420), vbBlack, B
        picCopies.CurrentX = sX + 300 - picCopies.TextWidth(iNum) - 60
        picCopies.CurrentY = sY + 420 - picCopies.TextHeight(iNum)
        picCopies.Print iNum
        sX = sX - 150
        sY = sY + 210
        iPage = iPage - 1
      Loop

      sX = sX - 200
      sY = 105
      iNum = iNum - 1
    Loop

  End If

End Sub

Private Sub cmdCancel_Click()
  mvarPrint = False
  Me.Hide
End Sub

Private Sub cmdPrint_Click()
  Dim lStart As Integer, lEnd As Integer
  Dim bEnable As Boolean
  bEnable = True
  lStart = Val(txtStart.Text)
  lEnd = Val(txtEnd.Text)

  If lStart = 0 Or lEnd = 0 Then
    bEnable = False
  End If

  If lStart > lEnd Then
    bEnable = False
  End If

  If lStart <> CInt(lStart) Then
    bEnable = False
  End If

  If lEnd <> CInt(lEnd) Then
    bEnable = False
  End If

  If optPrint(0).Value Then
    bEnable = True
    mvarStart = mvarCurrent
    mvarEnd = mvarCurrent
  ElseIf optPrint(1).Value Then
    bEnable = True
    mvarStart = 1
    mvarEnd = mvarMax
  ElseIf optPrint(2).Value Then
    mvarStart = lStart
    mvarEnd = lEnd
  End If

  mvarCopies = Val(txtCopies.Text)

  If Not bEnable Then
    MsgBox "Please enter a valid page range.", vbOKOnly, "Range Error"
    mvarPrint = False
  Else
    mvarPrint = True
    Me.Hide
  End If

End Sub

Public Property Get PrintDoc() As Boolean
  PrintDoc = mvarPrint
End Property

Public Property Let PageCurrent(ByVal vNewValue As Integer)
  mvarCurrent = vNewValue
End Property

Public Property Get PageStart() As Integer
  PageStart = mvarStart
End Property

Public Property Get PageEnd() As Integer
  PageEnd = mvarEnd
End Property

Public Property Get Collate() As Boolean
  Collate = mvarCollate
End Property

Public Property Get Copies() As Integer
  Copies = mvarCopies
End Property

Public Property Get PrinterNumber() As Integer
  PrinterNumber = mvarPrinter
End Property

Public Property Let PageMax(ByVal vNewValue As Integer)
  mvarMax = vNewValue
End Property

Private Sub cmdProperties_Click()
  Dim sPrinterName As String
  Dim hPrinter As Long
  Dim pdDefault As PRINTER_DEFAULTS
  Dim lReturn As Long
  sPrinterName = Printers(mvarPrinter).DeviceName

  With pdDefault
    .pDatatype = 0 ' vbNullString
    .pDesiredAccess = PRINTER_ALL_ACCESS
    .pDevMode = 0
  End With

  lReturn = OpenPrinter(sPrinterName, hPrinter, pdDefault)

  If lReturn = 0 Then
    ' Not an admin, try reduced privileges
    pdDefault.pDesiredAccess = STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_USE
    lReturn = OpenPrinter(sPrinterName, hPrinter, pdDefault)

    If lReturn = 0 Then
      MsgBox "Unable to open selected Printer's properties.", vbOKOnly + vbExclamation, "Warning"
    End If

  End If

  If hPrinter Then
    lReturn = PrinterProperties(Me.hWnd, hPrinter)
    lReturn = ClosePrinter(hPrinter)
  End If

End Sub

Private Sub cmdSave_Click()
'  Dim strFile As String
'  On Error Resume Next
'
'  With cDialog
'    .Filter = mstrFileDescription & "(*." & mstrSaveExt & ")|*." & mstrSaveExt & "|All Files (*.*)|*.*"
'    .DefaultExt = "*." & mstrSaveExt
'    .DialogTitle = "Save " & mstrFileDescription
'    .CancelError = True
'    .ShowSave
'
'    If Err.Number = cdlCancel Then
'      Exit Sub
'    End If
'
'    mstrSaveName = .Filename
'  End With
'
'  mblnSave = True
'  mvarPrint = False
'  Me.Hide
End Sub

Private Sub Form_Load()
  Dim prtPrinter As Printer
  Dim iPrinter As Integer
  Dim sPrinter As String
  Dim sY As Single
  ' Change options dependent on flags and default values
  bInternal = True
  'cmdSave.Visible = mblnShowSave
  ' Display Printer List

    Util.SetNumber (txtStart.hWnd)
    Util.SetNumber (txtEnd.hWnd)
    Util.SetNumber (txtCopies.hWnd)
    
  If CBool(mvarFlags And ShowPrinter_po) Or mvarFlags = ChoosePrinterOnly_po Then
    fraPrinter.Visible = True
    sY = sY + fraPrinter.Height + 120
    cboPrinter.Clear
    iPrinter = 0

    For Each prtPrinter In Printers
      sPrinter = prtPrinter.DeviceName & " on "

      If VBA.Right$(prtPrinter.Port, 1) = ":" Then
        sPrinter = sPrinter & Left$(prtPrinter.Port, Len(prtPrinter.Port) - 1)
      Else
        sPrinter = sPrinter & prtPrinter.Port
      End If

      cboPrinter.AddItem sPrinter
      cboPrinter.ItemData(cboPrinter.NewIndex) = iPrinter

      If prtPrinter.DeviceName = Printer.DeviceName And Printer.Port = prtPrinter.Port Then
        cboPrinter.ListIndex = cboPrinter.NewIndex
        mvarPrinter = iPrinter
      End If

      iPrinter = iPrinter + 1
    Next

  Else
    fraPrinter.Visible = False
  End If

  If mvarFlags = ChoosePrinterOnly_po Then
    fraOptions.Visible = False
    fraCopies.Visible = False
  Else
    ' Display Page range print options
    optPrint(1).Value = True
    optPrint(2).Enabled = CBool(mvarMax > 1)
    lblPages.Caption = mvarCurrent & " / " & mvarMax
    fraOptions.Top = sY
    sY = sY + fraOptions.Height + 120
    txtStart.Enabled = CBool(mvarMax > 1)
    txtEnd.Enabled = CBool(mvarMax > 1)
    txtStart.Text = "1"
    txtEnd.Text = mvarMax
    ' Display copy and collation options

    If CBool(mvarFlags And ShowCopies_po) Then
      fraCopies.Visible = True
      mvarCollate = True
      Copies_ShowImage
      fraCopies.Top = sY
      sY = sY + fraCopies.Height + 120
    Else
      fraCopies.Visible = False
    End If

  End If

  ' Adjust form size
  'cmdSave.Top = sY
  cmdPrint.Top = sY
  cmdCancel.Top = sY
  sY = sY + cmdPrint.Height + 120
  sY = (Me.Height - Me.ScaleHeight) + sY
  Me.Height = sY
  bInternal = False
End Sub

Private Sub optPrint_Click(Index As Integer)
  txtStart.Locked = CBool(Index <> 2)
  txtEnd.Locked = CBool(Index <> 2)

  If Index = 0 Then
    txtStart.Text = mvarCurrent
    txtEnd.Text = mvarCurrent
  Else
    txtStart.Text = 1
    txtEnd.Text = mvarMax
  End If

End Sub

Private Sub txtEnd_GotFocus()
  optPrint(2).Value = True
End Sub

Private Sub txtStart_GotFocus()
  optPrint(2).Value = True
End Sub

Public Sub SetCaptions(ByVal intLanguageOffset As Integer, ByVal blnLanguageRev As Boolean, intCharset As Integer)
  Me.Font.Charset = intCharset
  Me.Caption = LoadResString(intLanguageOffset + 8)
  Me.chkCollate.Font.Charset = intCharset
  Me.chkCollate.Caption = LoadResString(intLanguageOffset + 9)
  Me.cmdCancel.Font.Charset = intCharset
  Me.cmdCancel.Caption = LoadResString(intLanguageOffset + 10)
  Me.cmdProperties.Font.Charset = intCharset
  Me.cmdProperties.Caption = LoadResString(intLanguageOffset + 11)
  Me.cmdPrint.Font.Charset = intCharset
  Me.cmdPrint.Caption = LoadResString(intLanguageOffset + 12)
  Me.fraCopies.Font.Charset = intCharset
  Me.fraCopies.Caption = LoadResString(intLanguageOffset + 13)
  Me.fraOptions.Font.Charset = intCharset
  Me.fraOptions.Caption = LoadResString(intLanguageOffset + 14)
  Me.fraPrinter.Font.Charset = intCharset
  Me.fraPrinter.Caption = LoadResString(intLanguageOffset + 15)
  Me.lblEnd.Font.Charset = intCharset
  Me.lblEnd.Caption = LoadResString(intLanguageOffset + 16)
  Me.lblName.Font.Charset = intCharset
  Me.lblName.Caption = LoadResString(intLanguageOffset + 17)
  Me.lblStart.Font.Charset = intCharset
  Me.lblStart.Caption = LoadResString(intLanguageOffset + 18)
  Me.lblStatus.Font.Charset = intCharset
  Me.lblStatus.Caption = LoadResString(intLanguageOffset + 19)
  Me.optPrint(0).Font.Charset = intCharset
  Me.optPrint(0).Caption = LoadResString(intLanguageOffset + 20)
  Me.optPrint(1).Font.Charset = intCharset
  Me.optPrint(1).Caption = LoadResString(intLanguageOffset + 21)
  Me.optPrint(2).Font.Charset = intCharset
  Me.optPrint(2).Caption = LoadResString(intLanguageOffset + 22)
  txtStart.Font.Charset = intCharset
  txtEnd.Font.Charset = intCharset
  txtCopies.Font.Charset = intCharset
  cboPrinter.Font.Charset = intCharset
  Controls_Size blnLanguageRev
End Sub

Private Sub Controls_Size(ByVal bReverse As Boolean)
  Dim sCurX As Single
  Dim sStrip As String
  ' Resize fraProperties
  sStrip = String_StripAmpersand(cmdProperties.Caption)
  sCurX = Me.TextWidth(sStrip) + 240

  If sCurX < 975 Then
    sCurX = 975
  End If

  cmdProperties.Width = sCurX
  sCurX = sCurX + lblName.Width
  cboPrinter.Width = 5295 - sCurX - 240

  If bReverse Then
    lblName.Left = 5415 - lblName.Width
    cmdProperties.Left = 120
    cboPrinter.Left = cmdProperties.Width + 240
  Else
    lblName.Left = 120
    cmdProperties.Left = 5415 - cmdProperties.Width
    cboPrinter.Left = lblName.Width + 240
  End If

  ' Resize fraPages

  If bReverse Then
    lblStatus.Left = 4425
    lblStart.Left = 4425
    lblEnd.Left = 4425
    lblPages.Left = 3120
    txtStart.Left = 3120
    txtEnd.Left = 3120
  End If

  sStrip = String_StripAmpersand(cmdPrint.Caption)
  sCurX = Me.TextWidth(sStrip) + 240
  sStrip = String_StripAmpersand(cmdCancel.Caption)

  If sCurX > Me.TextWidth(sStrip) + 240 Then
    sCurX = Me.TextWidth(sStrip) + 240
  End If

  If sCurX < 975 Then
    sCurX = 975
  End If

  If bReverse Then
    cmdCancel.Left = 120
    cmdCancel.Width = sCurX
    cmdPrint.Left = sCurX + 240
    cmdPrint.Width = sCurX
  Else
    cmdCancel.Width = sCurX
    cmdCancel.Left = 5655 - sCurX
    cmdPrint.Width = sCurX
    cmdPrint.Left = 5535 - sCurX * 2
  End If

End Sub

Private Function String_StripAmpersand(ByVal sText As String) As String
  String_StripAmpersand = Replace(sText, "&", "")
End Function

