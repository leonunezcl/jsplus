VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "jsPrinter"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ==============================================================
' Project:      Enhance Print object
' Author:       edward moth
' Copyright:    © 2000 qbd software ltd
'
' ==============================================================
' Module:       qcPrinter
' Purpose:      Class holding enhanced Print object
' ==============================================================

Option Explicit
' ALIGNMENT Enum
Public Enum qePrinterAlign
  eLeft
  eCentre
  eRight
  eJustify
End Enum
' END LINE CHARACTER Enum
Public Enum qePrinterChar
  eNone
  eSpace
  eDash
  eLine
  eOops
End Enum
' PAPERSIZE Enum
Public Enum qePrinterPaperSize
  eA3size = 8
  eA4size = 9
  eA5size = 11
  eB4size = 12
End Enum
' SCALEMODE Enum
Public Enum qePrinterScale
  eTwip = 1
  eInch = 5
  eMillimetre = 6
  eCentimetre = 7
End Enum
Private Type qePrinterPageInfo
  Width As Single
  Height As Single
  AvailWidth As Single
  AvailHeight As Single
  LeftM As Single
  RightM As Single
  TopM As Single
  BottomM As Single
End Type


Private mvarPageSize As qePrinterPaperSize
Private mvarScaleMode As qePrinterScale
Private mvarMarginLeft As Single
Private mvarMarginRight As Single
Private mvarMarginTop As Single
Private mvarMarginBottom As Single
Private mvarPages As Integer
Private bPageChange As Boolean

Private qcolText As New qcTextItemCol

Private qPage As qePrinterPageInfo

Friend Sub PreviewPage(picBox As PictureBox, Page As Integer)

GetPageSize
With picBox
  .Width = qPage.Width
  .Height = qPage.Height
End With
PrintText picBox, Page, Page

End Sub
Private Sub PrintText(Target As Object, _
                      StartPage As Integer, _
                      EndPage As Integer)

' Procedure prints the text items that fall between StartPage
' and EndPage.  In the case of preview a single page
' The procedure moves through the collection of text items
' and calculates there position on each page

Dim sWidth As Single, sHeight As Single
Dim sSizeX As Single, sSizeY As Single
Dim sText As String, sChar2 As String
Dim sLine As String, sChar As String, sWord As String
Dim iLines As Integer, iCount As Integer, iSpace As Integer
Dim iMarker As Integer
Dim sCharWidth As Single, sLineHeight As Single
Dim sCurrentPosY As Single, sCurrentPosX As Single
Dim eCharType As qePrinterChar, eEnd As qePrinterChar
Dim bNewLine As Boolean, bFinish As Boolean, bPrint As Boolean
Dim bFound As Boolean
Dim iItems As Integer, iCurrentPage As Integer
Dim sCurrentHeight As Single
Dim qText As qcTextItem
Dim sIndentL As Single, sIndentR As Single
If qPage.AvailHeight <= 0 Or qPage.AvailWidth <= 0 Then
  Exit Sub
End If
  
  
iCurrentPage = 1
iItems = 1
If iCurrentPage >= StartPage And iCurrentPage <= EndPage Then
  bPrint = True
Else
  bPrint = False
End If
bFinish = False
sCurrentHeight = 0
' Get page height
sHeight = qPage.AvailHeight
    
' Loop for each text item
Do While iItems <= qcolText.Count And Not bFinish
  
' Get font attributes for current text item
  Set qText = qcolText(iItems)
  sText = qText.Text
  With Target.Font
    .Name = qText.FontName
    .Size = qText.FontSize
    .Bold = qText.FontBold
    .Italic = qText.FontItalic
    .Underline = qText.FontUnderline
  End With
  Target.ForeColor = qText.FontColor

' Calculate page width including any indents
  sIndentL = ConvertToTwip(qText.IndentScale, qText.IndentLeft)
  sIndentR = ConvertToTwip(qText.IndentScale, qText.IndentRight)
    
  sWidth = qPage.AvailWidth - sIndentL - sIndentR
  If sWidth <= 0 Then
    Exit Sub
  End If
    
  With Target
    If iItems = 1 Then
'Make sure Top margin is accounted for on first page
      .CurrentY = qPage.TopM
    End If
    sLineHeight = .TextHeight("H")
    sLine = ""
    sWord = ""
    eEnd = eNone
    bNewLine = False
    iCount = 1
' Loop through text to create lines
    Do
' Check whether new line will break page borders
' and whether page is relevant to print process
      If sLineHeight + sCurrentHeight > sHeight Then
        iCurrentPage = iCurrentPage + 1
        sCurrentHeight = 0
        If bPrint And TypeOf Target Is Printer Then
          .NewPage
          .CurrentY = qPage.TopM
        End If
        If iCurrentPage >= StartPage And iCurrentPage <= EndPage Then
          bPrint = True
        ElseIf iCurrentPage > EndPage Then
          bFinish = True
          GoTo PrintTextEndLine
        End If
      End If
' Find suitable line break
      Do
        eCharType = eNone
        sChar = Mid$(sText, iCount, 1)
' CHARACTER CHECK: Look for potential line breaks or where text
'                  width is greater than boundary
        Select Case sChar
          Case " "
            eCharType = eSpace
          Case "-"
            sSizeX = .TextWidth(sLine & sWord & sChar)
            If sSizeX > sWidth Then
              eCharType = eOops
            Else
              eCharType = eDash
            End If
          Case vbLf
            sChar = ""
            eCharType = eLine
          Case vbCr
            If iCount < Len(sText) Then
              If Mid$(sText, iCount + 1, 1) = vbLf Then
                iCount = iCount + 1
              End If
            End If
            sChar = ""
            eCharType = eLine
          Case Else
' CHARACTER CHECK: See if addition of character makes line too long
            sSizeX = .TextWidth(sLine & sWord & sChar)
            If sSizeX > sWidth Then
              eCharType = eOops
            Else
              sWord = sWord & sChar
            End If
        End Select
        iCount = iCount + 1
      Loop While iCount <= Len(sText) And eCharType = eNone
' LINE SPLIT: Examine potential line break
      If iCount > Len(sText) Then
        eCharType = eLine
      End If
      Select Case eCharType
        Case qePrinterChar.eNone
          sLine = sLine & sWord
          eEnd = eLine
        Case qePrinterChar.eOops
          If eEnd = eNone Then
            sLine = sWord
            sWord = sChar
          Else
            sLine = Trim$(sLine)
            sWord = sWord & sChar
          End If
          bNewLine = True
        Case qePrinterChar.eDash, qePrinterChar.eSpace
          eEnd = eCharType
          sLine = sLine & sWord & sChar
          sWord = ""
        Case qePrinterChar.eLine
          sLine = sLine & sWord
          eEnd = eLine
          sWord = ""
          bNewLine = True
      End Select
' LINE SPLIT: Add new line if required
      If bNewLine And bPrint Then
        .CurrentX = qPage.LeftM + sIndentL
' ALIGNMENT: Calculate position of line dependent on alignment setting
        Select Case qText.Alignment
          Case qePrinterAlign.eLeft
            Target.Print sLine

          Case qePrinterAlign.eCentre
            sSizeX = (sWidth - .TextWidth(sLine)) / 2 + qPage.LeftM + sIndentL
            .CurrentX = sSizeX
            Target.Print sLine

          Case qePrinterAlign.eRight
            sSizeX = qPage.Width - qPage.RightM - .TextWidth(sLine) - sIndentR
            .CurrentX = sSizeX
            Target.Print sLine

          Case qePrinterAlign.eJustify
' ALIGNMENT: Full justification is more complex.  Find spaces
'            and calculate extra spacing required
' NEXT LINE: Check to see if line has an line break
            If eEnd <> eLine Then
              sCharWidth = .TextWidth(" ")
              iSpace = 0
              iMarker = 0
              Do
                iMarker = InStr(iMarker + 1, sLine, " ")
                If iMarker > 0 Then
                  iSpace = iSpace + 1
                End If
              Loop While iMarker > 0
              sSizeX = sWidth - .TextWidth(sLine)
              bFound = False
' ALIGNMENT: Check number of spaces and extra size, if too large
'            use character justification as well as word justification
              If iSpace > 0 Then
                If sSizeX \ iSpace > sCharWidth * 3 Then
                  bFound = True
                End If
              Else
                bFound = True
              End If
              If bFound Then
                sSizeY = Len(sLine) - 1 + (iSpace * 2)
                sSizeY = sSizeX / sSizeY
                sSizeX = sSizeY * 3
              Else
                sSizeX = sSizeX / iSpace
                sSizeY = 0
              End If
              iMarker = 1
              .CurrentX = qPage.LeftM + sIndentL
              Do While iMarker <= Len(sLine)
                sChar2 = Mid$(sLine, iMarker, 1)
                sCurrentPosX = .CurrentX + .TextWidth(sChar2)
                sCurrentPosY = .CurrentY
                Target.Print sChar2
                If sChar2 = " " Then
                  sCurrentPosX = sCurrentPosX + sSizeX
                Else
                  sCurrentPosX = sCurrentPosX + sSizeY
                End If
                .CurrentX = sCurrentPosX
                .CurrentY = sCurrentPosY
                iMarker = iMarker + 1
              Loop
              Target.Print ""
            Else
              Target.Print sLine
            End If
        End Select
      End If
      If bNewLine Then
        sCurrentHeight = sCurrentHeight + sLineHeight
        sLine = ""
        eEnd = eNone
        bNewLine = False
      End If
PrintTextEndLine:
              
    Loop While iCount <= Len(sText) And Not bFinish
  End With
      
  iItems = iItems + 1
Loop
' If printing to Printer use the EndDoc method to
' start print
    
If TypeOf Target Is Printer Then
  Printer.EndDoc
End If
      
      
End Sub
                  
Private Sub GetPageSize()
With qPage
  .Width = Printer.ScaleWidth
  .Height = Printer.ScaleHeight
  .LeftM = ConvertToTwip(mvarScaleMode, mvarMarginLeft)
  .RightM = ConvertToTwip(mvarScaleMode, mvarMarginRight)
  .TopM = ConvertToTwip(mvarScaleMode, mvarMarginTop)
  .BottomM = ConvertToTwip(mvarScaleMode, mvarMarginBottom)
  .AvailWidth = .Width - .LeftM - .RightM
  .AvailHeight = .Height - .TopM - .BottomM
End With
End Sub

Public Property Get TextItem() As qcTextItemCol
Set TextItem = qcolText
End Property

Public Property Let TextItem(ByVal vData As qcTextItemCol)
Set qcolText = vData
bPageChange = True
End Property

Public Property Get ItemCount() As Long
ItemCount = qcolText.Count
End Property

Public Property Let MarginBottom(ByVal vData As Single)
If mvarMarginBottom <> vData Then
  bPageChange = True
  mvarMarginBottom = vData
  qPage.BottomM = ConvertToTwip(mvarScaleMode, vData)
End If
End Property

Public Property Get MarginBottom() As Single
MarginBottom = mvarMarginBottom
End Property

Public Property Let MarginTop(ByVal vData As Single)
If mvarMarginTop <> vData Then
  bPageChange = True
  mvarMarginTop = vData
  qPage.TopM = ConvertToTwip(mvarScaleMode, vData)
End If
End Property

Public Property Get MarginTop() As Single
MarginTop = mvarMarginTop
End Property


Public Property Let MarginRight(ByVal vData As Single)
If mvarMarginRight <> vData Then
  bPageChange = True
  mvarMarginRight = vData
  qPage.RightM = ConvertToTwip(mvarScaleMode, vData)
End If
End Property

Public Property Get MarginRight() As Single
MarginRight = mvarMarginRight
End Property

Public Property Let MarginLeft(ByVal vData As Single)
If mvarMarginLeft <> vData Then
  bPageChange = True
  mvarMarginLeft = vData
  qPage.LeftM = ConvertToTwip(mvarScaleMode, vData)
End If
End Property

Public Property Get MarginLeft() As Single
MarginLeft = mvarMarginLeft
End Property

Public Sub PrintDoc(ByVal StartPage As Integer, _
                    ByVal EndPage As Integer)

' Procedure sets the Target in PrintText to Printer
' printing the pages defined by StartPage and EndPage
PrintText Printer, StartPage, EndPage

End Sub

Public Sub Preview()
Dim fShow As Form

' Display the preview form
Set frmPreview.Document = Me
Set fShow = frmPreview
fShow.Show vbModal


End Sub

Public Property Let ScaleMode(ByVal vData As qePrinterScale)
Dim iCount As Integer

If vData <> mvarScaleMode Then
  mvarMarginBottom = ConvertFromTwip(vData, qPage.BottomM)
  mvarMarginTop = ConvertFromTwip(vData, qPage.TopM)
  mvarMarginLeft = ConvertFromTwip(vData, qPage.LeftM)
  mvarMarginRight = ConvertFromTwip(vData, qPage.RightM)
  mvarScaleMode = vData
  iCount = 1
  Do While iCount <= qcolText.Count
    qcolText(iCount).IndentScale = mvarScaleMode
    iCount = iCount + 1
  Loop
End If
    
End Property


Public Property Get ScaleMode() As qePrinterScale
ScaleMode = mvarScaleMode
End Property


Public Property Get Pages() As Integer

Dim iCount As Integer
Dim lPages As Integer
Dim sHeight As Single
Dim sTextH As Single
Dim qText As qcTextItem

' Calculate the number of pages only if a change has taken
' place - if not return the last page calculation
If bPageChange Then
  lPages = 1
  iCount = 1
  GetPageSize
  Do While iCount <= qcolText.Count
    Set qText = qcolText.Item(iCount)
    qText.GetSize qPage.LeftM, qPage.RightM
    sTextH = qText.LineHeight * qText.Lines

    Do While sHeight + sTextH > qPage.AvailHeight
      lPages = lPages + 1
      sTextH = sTextH - (((qPage.AvailHeight - sHeight) \ qText.LineHeight) * qText.LineHeight)
      sHeight = 0
    Loop
    sHeight = sHeight + sTextH
    iCount = iCount + 1
  Loop
  mvarPages = lPages
End If

Pages = mvarPages

End Property


Public Property Let PageSize(ByVal vData As qePrinterPaperSize)
If mvarPageSize <> vData Then
  bPageChange = True
  mvarPageSize = vData
End If
End Property


Public Property Get PageSize() As qePrinterPaperSize
PageSize = mvarPageSize
End Property

Public Function AddText(ByVal Text As String, _
                        ByVal FontName As String, _
                        ByVal FontSize As Single, _
                        Optional FontBold As Boolean = False, _
                        Optional FontItalic As Boolean = False, _
                        Optional FontUnderline As Boolean = False, _
                        Optional FontColor As Long = &H0, _
                        Optional Align As qePrinterAlign = eLeft, _
                        Optional IndentLeft As Single = 0, _
                        Optional IndentRight As Single = 0, _
                        Optional Key As String) As Boolean


Dim qText As qcTextItem

Set qText = New qcTextItem
With qText
  .Text = Text
  .FontName = FontName
  .FontSize = FontSize
  .FontBold = FontBold
  .FontItalic = FontItalic
  .FontUnderline = FontUnderline
  .FontColor = FontColor
  .Alignment = Align
  .IndentScale = mvarScaleMode
  .IndentLeft = IndentLeft
  .IndentRight = IndentRight
End With

qcolText.Add qText, Key

Set qText = Nothing
bPageChange = True

End Function

Private Sub Class_Initialize()
mvarScaleMode = eTwip
mvarPageSize = Printer.PaperSize
mvarScaleMode = eTwip
mvarMarginLeft = 1440
mvarMarginRight = 1440
mvarMarginTop = 1440
mvarMarginBottom = 1440
mvarPages = 1
bPageChange = False
GetPageSize

End Sub

Private Sub Class_Terminate()
Set qcolText = Nothing

End Sub
Private Function ConvertToTwip(ByVal eScale As qePrinterScale, _
                               ByVal sValue As Single) As Single

Dim sNewValue As Single
' Convert value to Twips
Select Case eScale
  Case qePrinterScale.eTwip
    sNewValue = sValue
  Case qePrinterScale.eInch
    sNewValue = sValue * 1440
  Case qePrinterScale.eCentimetre
    sNewValue = sValue * 567
  Case qePrinterScale.eMillimetre
    sNewValue = sValue * 56.7
End Select
ConvertToTwip = sNewValue

End Function

Private Function ConvertFromTwip(ByVal eScale As qePrinterScale, _
                                 ByVal sValue As Single) As Single

Dim sNewValue As Single
'Convert value from Twips
Select Case eScale
  Case qePrinterScale.eTwip
    sNewValue = sValue
  Case qePrinterScale.eInch
    sNewValue = sValue / 1440
  Case qePrinterScale.eCentimetre
    sNewValue = sValue / 567
  Case qePrinterScale.eMillimetre
    sNewValue = sValue / 56.7
End Select

ConvertFromTwip = sNewValue

End Function


Public Sub ResetItems()

Set qcolText = Nothing
Set qcolText = New qcTextItemCol

End Sub

Public Function RemoveItem(IndexKey As Variant) As Boolean

On Error GoTo RemoveItemError
' Remove Item from TextItem collection.  Returns True if
' successful, otherwise returns false

qcolText.Remove IndexKey
bPageChange = True
RemoveItem = True
Exit Function

RemoveItemError:
RemoveItem = False

End Function
