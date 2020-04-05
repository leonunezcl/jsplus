VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "qcTextItem"
Attribute VB_GlobalNameSpace = False
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
' Module:       qcTextItem
' Purpose:      Holds details of each Text Item in Document
' ==============================================================

Option Explicit

Private mvarText As String
Private mvarLineHeight As Single
Private mvarLines As Integer
Private mvarFontName As String
Private mvarFontSize As Single
Private mvarFontBold As Boolean
Private mvarFontItalic As Boolean
Private mvarFontUnderline As Boolean
Private mvarFontColor As Long
Private mvarIndentRight As Single
Private mvarIndentLeft As Single
Private mvarAlignment As Currency

Private mvarIndentScale As qePrinterScale
Friend Property Let IndentScale(ByVal vData As qePrinterScale)
mvarIndentScale = vData
End Property


Friend Property Get IndentScale() As qePrinterScale
IndentScale = mvarIndentScale
End Property
Friend Sub GetSize(MarginLeft As Single, _
                   MarginRight As Single, _
                   Optional ScaleMode As qePrinterScale)

Dim sHeight As Single
Dim sWidth As Single
Dim sSizeX As Single
Dim sLine As String, sChar As String, sWord As String
Dim iLines As Integer, iCount As Integer
Dim eCharType As qePrinterChar, eEnd As qePrinterChar
Dim bNewLine As Boolean
Dim sfFont As New StdFont

With Printer.Font
  .Name = mvarFontName
  .Size = mvarFontSize
  .Bold = mvarFontBold
  .Italic = mvarFontItalic
  .Underline = mvarFontUnderline
End With
mvarLineHeight = Printer.TextHeight("H")

sWidth = Printer.ScaleWidth - MarginLeft - MarginRight - mvarIndentLeft - mvarIndentRight

If sWidth <= 0 Then
  Exit Sub
End If

With Printer
  sLine = ""
  sWord = ""
  eEnd = eNone
  iCount = 1

  Do
    bNewLine = False

    Do
      Do
        eCharType = eNone
        sChar = Mid$(mvarText, iCount, 1)
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
            If iCount < Len(mvarText) Then
              If Mid$(mvarText, iCount + 1, 1) = vbLf Then
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
      Loop While iCount <= Len(mvarText) And eCharType = eNone
' LINE SPLIT: Examine potential line break
      If iCount > Len(mvarText) Then
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
      If bNewLine Then
        iLines = iLines + 1
        sLine = ""
        eEnd = eNone
      End If
    Loop While Not bNewLine
  Loop While iCount <= Len(mvarText)
End With
mvarLines = iLines

End Sub

Public Property Let Alignment(ByVal vData As qePrinterAlign)
mvarAlignment = vData
End Property

Public Property Get Alignment() As qePrinterAlign
Alignment = mvarAlignment
End Property

Public Property Let IndentLeft(ByVal vData As Single)
    
mvarIndentLeft = ConvertToTwip(mvarIndentScale, vData)
End Property

Public Property Get IndentLeft() As Single
IndentLeft = ConvertFromTwip(mvarIndentScale, mvarIndentLeft)
End Property

Public Property Let IndentRight(ByVal vData As Single)
mvarIndentRight = ConvertToTwip(mvarIndentScale, vData)
End Property

Public Property Get IndentRight() As Single
IndentRight = ConvertFromTwip(mvarIndentScale, mvarIndentRight)
End Property

Public Property Let FontName(ByVal vData As String)
mvarFontName = vData
End Property

Public Property Get FontName() As String
FontName = mvarFontName
End Property
Public Property Let FontSize(ByVal vData As Single)
mvarFontSize = vData
End Property

Public Property Get FontSize() As Single
FontSize = mvarFontSize
End Property

Public Property Let FontColor(ByVal vData As Long)
mvarFontColor = vData
End Property

Public Property Get FontColor() As Long
FontColor = mvarFontColor
End Property

Public Property Let FontBold(ByVal vData As Boolean)
mvarFontBold = vData
End Property

Public Property Get FontBold() As Boolean
FontBold = mvarFontBold
End Property

Public Property Let FontItalic(ByVal vData As Boolean)
mvarFontItalic = vData
End Property

Public Property Get FontItalic() As Boolean
FontItalic = mvarFontItalic
End Property
Public Property Let FontUnderline(ByVal vData As Boolean)
mvarFontUnderline = vData
End Property

Public Property Get FontUnderline() As Boolean
FontUnderline = mvarFontUnderline
End Property

Friend Property Get Lines() As Integer
Lines = mvarLines
End Property

Friend Property Get LineHeight() As Single
LineHeight = mvarLineHeight
End Property

Public Property Let Text(ByVal vData As String)
mvarText = vData
End Property

Public Property Get Text() As String
Text = mvarText
End Property
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


