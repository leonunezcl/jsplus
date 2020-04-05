VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrintPreview"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_FileName As String


Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal pFileName As String)
    m_FileName = pFileName
End Property
Public Sub PageSetup()

    On Error GoTo error_setup
    
    frmPageSetup.Show vbModal
    
    Exit Sub
    
error_setup:
    Err = 0
    
End Sub

Public Sub StartPreview()

    sPrintHeader = util.LeeIni(IniPath, "print", "Header")
    If sPrintHeader = "" Then sPrintHeader = "^N"
    
    sPrintFooter = util.LeeIni(IniPath, "print", "Footer")
    If sPrintFooter = "" Then sPrintFooter = "^D"
    
    gLeftMargin = IIf(util.LeeIni(IniPath, "print", "gLeftMargin") = "", 25, util.LeeIni(IniPath, "print", "gLeftMargin"))
    
    gRightMargin = IIf(util.LeeIni(IniPath, "print", "gRightMargin") = "", 25, util.LeeIni(IniPath, "print", "gRightMargin"))
    
    gTopMargin = IIf(util.LeeIni(IniPath, "print", "gTopMargin") = "", 25, util.LeeIni(IniPath, "print", "gTopMargin"))
    
    gBottomMargin = IIf(util.LeeIni(IniPath, "print", "gBottomMargin") = "", 25, util.LeeIni(IniPath, "print", "gBottomMargin"))

    'MsgBox "1"
    
    frmMain.FileName = m_FileName
    Load frmMain
    
    'MsgBox "2"
    
    'setup header and footer
    frmMain.RTF1.Text = frmMain.txtCode.Text
    sPrintText = frmMain.RTF1.Text
    sHeader = SetPrintLine(sPrintHeader)
    sFooter = SetPrintLine(sPrintFooter)
    
    'MsgBox "3"
    
    sPrintText = sHeader & vbCrLf & vbCrLf & sPrintText & vbCrLf & vbCrLf & sFooter
    frmMain.rtfTmp.Text = sPrintText

    'MsgBox "4"
    
    frmDocPreview.Show vbModal
    
    Unload frmMain
    
End Sub


Private Sub Class_Initialize()

    gLeftMargin = 25
    gRightMargin = 25
    gTopMargin = 25
    gBottomMargin = 25
    
End Sub


