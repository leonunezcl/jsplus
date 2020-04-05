Attribute VB_Name = "modPrint"
' Module for printing from RTFbox

Option Explicit

Public Const WM_USER = &H400
Public Const EM_FORMATRANGE As Long = WM_USER + 57
Public Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Public Const PHYSICALOFFSETX As Long = 112
Public Const PHYSICALOFFSETY As Long = 113
Private Const C_INI = "jsplus.ini"

Public util As New cLibrary

Public sPrintHeader As String   ' used with frmPagesetup and the print options
Public sPrintFooter As String   ' used with frmPagesetup and the print options
Public sPrintText As String     ' used with frmPagesetup and the print options
Public sHeader As String        ' used with frmPagesetup and the print options
Public sFooter As String        ' used with frmPagesetup and the print options
Public gPrint As Boolean
Public gLeftMargin As Integer           ' Print Preview
Public gRightMargin As Integer          ' Print Preview
Public gTopMargin As Integer            ' Print Preview
Public gBottomMargin As Integer         ' Print Preview

Global pPaperSize As String     ' Holds the current printer papersize
Global papersize As String
Global pOrientation As PrinterOrientationConstants   ' Holds the current printer paper orientation

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CharRange
    cpMin As Long ' First character of range (0 For start of doc)
    cpMax As Long ' Last character of range (-1 For End of doc)
End Type

Private Type FormatRange
    hdc As Long ' Actual DC to draw on
    hdcTarget As Long ' Target DC For determining text formatting
    rc As Rect ' Region of the DC to draw to (in twips)
    rcPage As Rect ' Region of the entire DC (page size) (in twips)
    chrg As CharRange ' Range of text to draw (see above declaration)
End Type

Private Declare Function GetDeviceCaps Lib "gdi32" ( _
    ByVal hdc As Long, ByVal nIndex As Long) As Long


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, _
    lp As Any) As Long


Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
    (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
    ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
                      
Public Function IniPath() As String

    IniPath = util.StripPath(App.Path) & C_INI
    
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PrintRTF - Prints the contents of a RichTextBox control using the provided margins
'
' RTF - A RichTextBox control to print
'
' LeftMarginWidth - Width of desired left margin in twips
'
' TopMarginHeight - Height of desired top margin in twips
'
' RightMarginWidth - Width of desired right margin in twips
'
' BottomMarginHeight - Height of desired bottom margin in twips
'
' Notes - If you are also using WYSIWYG_RTF() on the provided RTF
' parameter you should specify the same LeftMarginWidth and
' RightMarginWidth that you used to call WYSIWYG_RTF()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, TopMarginHeight, RightMarginWidth, BottomMarginHeight)
    
    On Error GoTo ErrorHandler
    Dim LeftOffset As Long, TopOffset As Long
    Dim LeftMargin As Long, TopMargin As Long
    Dim RightMargin As Long, BottomMargin As Long
    Dim fr As FormatRange
    Dim rcDrawTo As Rect
    Dim rcPage As Rect
    Dim TextLength As Long
    Dim NextCharPosition As Long
    Dim r As Long
    ' Start a print job to get a valid Printer.hDC
    Printer.Print Space(1)
    Printer.ScaleMode = vbTwips
    ' Get the offsett to the printable area on the page in twips
    LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
    PHYSICALOFFSETX), vbPixels, vbTwips)
    TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, _
    PHYSICALOFFSETY), vbPixels, vbTwips)
    ' Calculate the Left, Top, Right, and Bottom margins
    LeftMargin = LeftMarginWidth - LeftOffset
    TopMargin = TopMarginHeight - TopOffset
    RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
    BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset
    ' Set printable area rect
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = Printer.ScaleWidth
    rcPage.Bottom = Printer.ScaleHeight
    ' Set rect in which to print (relative to printable area)
    rcDrawTo.Left = LeftMargin
    rcDrawTo.Top = TopMargin
    rcDrawTo.Right = RightMargin
    rcDrawTo.Bottom = BottomMargin
    ' Set up the print instructions
    fr.hdc = Printer.hdc ' Use the same DC For measuring and rendering
    fr.hdcTarget = Printer.hdc ' Point at printer hDC
    fr.rc = rcDrawTo ' Indicate the area On page to draw to
    fr.rcPage = rcPage ' Indicate entire size of page
    fr.chrg.cpMin = 0 ' Indicate start of text through
    fr.chrg.cpMax = -1 ' End of the text
    ' Get length of text in RTF
    
    TextLength = Len(RTF.Text)
    ' Loop printing each page until done
    Do
        ' Print the page by sending EM_FORMATRANGE message
        NextCharPosition = SendMessage(RTF.hWnd, EM_FORMATRANGE, True, fr)
        If NextCharPosition >= TextLength Then Exit Do 'If done then exit
        fr.chrg.cpMin = NextCharPosition ' Starting position For next page
        Printer.NewPage ' Move On to Next page
        Printer.Print Space(1) ' Re-initialize hDC
        fr.hdc = Printer.hdc
        fr.hdcTarget = Printer.hdc
    Loop
    ' Commit the print job
    Printer.EndDoc
    ' Allow the RTF to free up memory
    r = SendMessage(RTF.hWnd, EM_FORMATRANGE, False, ByVal CLng(0))
ErrorHandler:
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WYSIWYG_RTF - Sets an RTF control to display itself the same as it
'               would print on the default printer
'
' RTF - A RichTextBox control to set for WYSIWYG display.
'
' LeftMarginWidth - Width of desired left margin in twips
'
' RightMarginWidth - Width of desired right margin in twips
'
' Returns - The length of a line on the printer in twips
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub WYSIWYG_RTF(RTF As RichTextBox, LeftMarginWidth As Long, RightMarginWidth As Long, TopMarginWidth As Long, BottomMarginWidth As Long, PrintableWidth As Long, PrintableHeight As Long)
   Dim LeftOffset As Long
   Dim LeftMargin As Long
   Dim RightMargin As Long
   Dim TopOffset As Long
   Dim TopMargin As Long
   Dim BottomMargin As Long
   Dim PrinterhDC As Long
   Dim r As Long

   ' Start a print job to initialize printer object
   Printer.Print Space(1)
   Printer.ScaleMode = vbTwips
   
   ' Get the left offset to the printable area on the page in twips
   LeftOffset = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
   LeftOffset = Printer.ScaleX(LeftOffset, vbPixels, vbTwips)
   
   ' Calculate the Left, and Right margins
   LeftMargin = LeftMarginWidth - LeftOffset
   RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
   
   ' Calculate the line width
   PrintableWidth = RightMargin - LeftMargin
   
   ' Get the top offset to the printable area on the page in twips
   TopOffset = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)
   TopOffset = Printer.ScaleX(TopOffset, vbPixels, vbTwips)
   
   ' Calculate the Left, and Right margins
   TopMargin = TopMarginWidth - TopOffset
   BottomMargin = (Printer.Height - BottomMarginWidth) - TopOffset
   
   ' Calculate the line width
   PrintableHeight = BottomMargin - TopMargin
    
   
   ' Create an hDC on the Printer pointed to by the Printer object
   ' This DC needs to remain for the RTF to keep up the WYSIWYG display
   PrinterhDC = CreateDC(Printer.DriverName, Printer.DeviceName, 0, 0)

   ' Tell the RTF to base it's display off of the printer
   '    at the desired line width
   r = SendMessage(RTF.hWnd, EM_SETTARGETDEVICE, PrinterhDC, ByVal PrintableWidth)

   ' Abort the temporary print job used to get printer info
   Printer.KillDoc
End Sub

Public Function SetPrintLine(ByVal sText As String) As String
'Purpose: Works with the Page Header and Page Footer and is used to
'Replace Control keys with actual text
'Demonstrates the use of the new VB6 Replace Funtion
    sText = Replace(sText, "^N", GetFile(frmMain.Caption))
    sText = Replace(sText, "^n", GetFile(frmMain.Caption))
    sText = Replace(sText, "^P", frmMain.Caption)
    sText = Replace(sText, "^p", frmMain.Caption)
    sText = Replace(sText, "^D", Format(Date, "dd mmmm yyyy"))
    sText = Replace(sText, "^d", Format(Date, "dd mmmm yyyy"))
    sText = Replace(sText, "^T", Format(Time, "hh:mm"))
    sText = Replace(sText, "^t", Format(Time, "hh:mm"))
    sText = Replace(sText, "\n", vbCrLf)
    SetPrintLine = sText
End Function

Function GetFile(sPath As String) As String
    'Returns only file title
    Dim i, j As Integer
    i = InStr(1, Reverse(sPath), "\")
    If i = 0 Then i = InStr(1, Reverse(sPath), "/")
    If i = 0 Then GetFile = sPath: Exit Function
    GetFile = Right(sPath, i - 1)
End Function
Public Function Reverse(sString As String) As String
'VB6 has this as an in-built function called
'StrReverse(String) but I am not sure of VB5.
Dim i As Integer, s As String
For i = 1 To Len(sString)
s = s & Mid(sString, Len(sString) + 1 - i, 1)
Next i
Reverse = s
End Function
Public Function getPrintSize(ByVal pPaperSize As String)
'Purpose: Works with the printer.papersize and is used to
'Replace Control keys with actual text
    Select Case pPaperSize
        Case 1
            pPaperSize = "Letter, 8.5 x 11 in."
        Case 2
            pPaperSize = "Letter Small, 8.5 x 11 in."
        Case 3
            pPaperSize = "Tabloid, 11 x 17 in."
        Case 4
            pPaperSize = "Ledger, 17 x 11 in."
        Case 5
            pPaperSize = "Legal, 8.5 x 14 in."
        Case 6
            pPaperSize = "Statement, 5.5 x 8.5 in."
        Case 7
            pPaperSize = "Executive, 7.5 x 10.5 in."
        Case 8
            pPaperSize = "A3, 297 x 420 mm"
        Case 9
            pPaperSize = "A4, 210 x 297 mm"
        Case 10
            pPaperSize = "A4 Small, 210 x 297 mm"
        Case 11
            pPaperSize = "A5, 148 x 210 mm"
        Case 12
            pPaperSize = "B4, 250 x 354 mm"
        Case 13
            pPaperSize = "B5, 182 x 257 mm"
        Case 14
            pPaperSize = "Folio, 8.5 x 13 in."
        Case 15
            pPaperSize = "Quarto, 215 x 275 mm"
        Case 16
            pPaperSize = "10 x 14 in."
        Case 17
            pPaperSize = "11 x 17 in."
        Case 18
            pPaperSize = "Note, 8.5 x 11 in."
        Case 19
            pPaperSize = "Envelope #9, 3.875 x 8.875 in."
        Case 20
            pPaperSize = "Envelope #10, 4.125 x 9.5 in."
        Case 21
            pPaperSize = "Envelope #11, 4.5 x 10.375 in."
        Case 22
            pPaperSize = "Envelope #12, 4.5 x 11 in."
        Case 23
            pPaperSize = "Envelope #14, 5 x 11.5 in."
        Case 24
            pPaperSize = "C size sheet"
        Case 25
            pPaperSize = "D size sheet"
        Case 26
            pPaperSize = "E size sheet"
        Case 27
            pPaperSize = "Envelope DL, 110 x 220 mm"
        Case 28
            pPaperSize = "Envelope C3, 324 x 458 mm"
        Case 29
            pPaperSize = "Envelope C4, 229 x 324 mm"
        Case 30
            pPaperSize = "Envelope C5, 162 x 229 mm"
        Case 31
            pPaperSize = "Envelope C6, 114 x 162 mm"
        Case 32
            pPaperSize = "Envelope C65, 114 x 229 mm"
        Case 33
            pPaperSize = "Envelope B4, 250 x 353 mm"
        Case 34
            pPaperSize = "Envelope B5, 176 x 250 mm"
        Case 35
            pPaperSize = "Envelope B6, 176 x 125 mm"
        Case 36
            pPaperSize = "Envelope, 110 x 230 mm"
        Case 37
            pPaperSize = "Envelope Monarch, 3.875 x 7.5 in."
        Case 38
            pPaperSize = "Envelope, 3.625 x 6.5 in."
        Case 39
            pPaperSize = "U.S. Standard Fanfold, 14.875 x 11 in."
        Case 40
            pPaperSize = "German Standard Fanfold, 8.5 x 12 in."
        Case 41
            pPaperSize = "German Legal Fanfold, 8.5 x 13 in."
        Case 42 To 49
            pPaperSize = "Unknown format"
        Case 50
            pPaperSize = "Letter Extra"
        Case 51
            pPaperSize = "Legal Extra"
        Case 53
            pPaperSize = "A4 Extra"
        Case 54 To 64
            pPaperSize = "Unknown format"
        Case 64
            pPaperSize = "A5 Extra"
        Case 65
            pPaperSize = "B5 Extra"
        Case 66 To 255
            pPaperSize = "Unknown format"
        Case 256
            pPaperSize = "User-defined"
        Case 257
            pPaperSize = "Commercial, 210 x 270 mm"
        Case 258
            pPaperSize = "Foolscap, 203 x  330 mm"
        Case 259
            pPaperSize = "Legal Small, 8,5 x 13 in."
        Case 260
            pPaperSize = "Tabloid Extra"
        Case 261
            pPaperSize = "A5 Tranverse"
        Case 262
            pPaperSize = "B5 Tranverse Extra"
        Case 263
            pPaperSize = "B5 Tranverse"
        Case 264
            pPaperSize = "Unknown format"
        Case 265
            pPaperSize = "A4 Tranverse"
        Case 266
            pPaperSize = "A4 Tranverse Extra"
        Case 267
            pPaperSize = "Large"
        End Select
        
    getPrintSize = pPaperSize
End Function
