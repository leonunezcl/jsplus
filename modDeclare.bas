Attribute VB_Name = "modDeclare"
Option Explicit
Public Type qtPageCacheType
  Cached As Boolean
  Item As Long
  ItemType As qeDocumentObjectTypeEnum
  KeepNext As Boolean
  Line As Integer
  CharPos As Long
  TempLIndent As Long
  TempRIndent As Long
  TempAlign As qePrinterAlign
  AbsTotal As Long
  AbsItem() As Long
  AbsType() As qeDocumentObjectTypeEnum
  PageHeaderRepeatStart As Long
  PageHeaderRepeatEnd As Long
  PageHeaderHeight As Single
  UsePageHeader As Boolean
End Type
Private Type qtPrinterPageInfo
  Width As Single
  Height As Single
  AvailWidth As Single
  AvailHeight As Single
  LeftM As Single
  RightM As Single
  TopM As Single
  BottomM As Single
  HeaderH As Single
  FooterH As Single
  HFAvailHeight As Single
End Type
Private Type qtCompressRunReplaceType
  CharCode As Byte
  Total As Long
  Position As Long
End Type
Private Type qtCompressSequenceType
  Char1 As Byte
  Char2 As Byte
  Total As Long
End Type
Public qPage As qtPrinterPageInfo
Public mTwipMLeft As Single
Public mTwipMRight As Single
Public mTwipMTop As Single
Public mTwipMBottom As Single
Public bPageChange As Boolean
Public bPrinterChange As Boolean
Public mvarOrientation As qePrintOrientation
Public mvarScaleMode As qePrinterScale
Public gstrTempPath As String
Public gblnCancelDocument As Boolean
Public gblnUpdateProgress As Boolean
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
' LANGUAGE INFORMATION
Public Const LOCALE_IDEFAULTLANGUAGE As Long = &H9
Public Const LOCALE_SYSTEM_DEFAULT As Long = &H400
Public Const LOCALE_USER_DEFAULT As Long = &H800
Public Declare Function GetLocaleInfo Lib "kernel32" _
   Alias "GetLocaleInfoA" _
   (ByVal Locale As Long, ByVal LCType As Long, _
    ByVal lpLCData As String, ByVal cchData As Long) As Long

Public Function ConvertToTwip(ByVal eScale As qePrinterScale, _
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

Public Function ConvertFromTwip(ByVal eScale As qePrinterScale, _
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

Public Function ConvertHTMColor(ByVal HTMColor As String) As Long
  Dim lRed As Long, lGreen As Long, lBlue As Long
  Dim sHexRed As String, sHexBlue As String, sHexGreen As String
  Dim lColor As Long
  HTMColor = UCase$(HTMColor)

  If HTMColor Like "[#][0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F][0-9A-F]" Then
    sHexRed = "&H" & Mid$(HTMColor, 2, 2)
    sHexGreen = "&H" & Mid$(HTMColor, 4, 2)
    sHexBlue = "&H" & Mid$(HTMColor, 6, 2)
    lBlue = CDec(sHexBlue) * &H10000
    lGreen = CDec(sHexGreen) * &H100
    lRed = CDec(sHexRed)
    lColor = lBlue + lGreen + lRed
  Else
    lColor = 0
  End If

  ConvertHTMColor = lColor
End Function

Public Function ConvertColorToHTM(ByVal lColor As Long) As String
  Dim lRed As Long, lGreen As Long, lBlue As Long
  Dim sHexRed As String, sHexBlue As String, sHexGreen As String
  Dim sHTMCol As String
  lRed = lColor Mod 256
  lGreen = ((lColor And &HFF00) / 256&) Mod 256&
  lBlue = (lColor And &HFF0000) / 65536
  sHexRed = VBA.Right$("00" & Hex$(lRed), 2)
  sHexGreen = VBA.Right$("00" & Hex$(lGreen), 2)
  sHexBlue = VBA.Right$("00" & Hex$(lBlue), 2)
  sHTMCol = "#" & sHexRed & sHexGreen & sHexBlue
  ConvertColorToHTM = sHTMCol
End Function

Public Function Property_ExtractTab(ByVal sFormat As String, ByVal sItem As String) As Variant
  Dim lStart As Long
  Dim lEnd As Long
  Dim vData As Variant
  Dim eAlign As qePrinterAlign
  lStart = InStr(1, sFormat, sItem)
  lEnd = InStr(lStart + 1, sFormat, " ")

  If lStart > 0 Then

    If lEnd = 0 Then
      vData = Property_Extract(VBA.Right$(sFormat, Len(sFormat) - lStart + 1))
    Else
      vData = Property_Extract(Mid$(sFormat, lStart, lEnd - lStart))
    End If

  End If

  If sItem = "ALIGN" Then

    Select Case vData
      Case "RIGHT"
        eAlign = eRight
      Case "CENTRE", "CENTER"
        eAlign = eCentre
      Case Else
        eAlign = eLeft
    End Select

    Property_ExtractTab = eAlign
  ElseIf IsNumeric(vData) Then
    Property_ExtractTab = CSng(vData)
  Else
    Property_ExtractTab = CStr(vData)
  End If

End Function

Public Function Property_Extract(ByVal sFormat As String)
  Dim lPos As Long
  Dim sReturn As String
  lPos = InStr(1, sFormat, "=")

  If lPos > 0 Then
    sReturn = VBA.Right$(sFormat, Len(sFormat) - lPos)

    If VBA.Left$(sReturn, 1) = "#" Then
      sReturn = ConvertHTMColor(sReturn)
    End If

  End If

  Property_Extract = sReturn
End Function

Public Function qExpand(ByVal SourceFile As String, ByRef DestinationFile As String) As Integer
  Dim lCount As Long, lSize As Long
  Dim byB() As Byte
  Dim iFile As Integer
  Dim sFile As String
  Dim bLast As Byte
  Dim lRun As Long
  Dim byNew() As Byte
  Dim lNewPos As Long
  Dim lIns As Long
  Dim bReplace As Boolean
  Dim sType As String
  Dim lOrigSize As Long
  Dim qReturn(255) As qtCompressSequenceType
  Dim nSeqChar As Integer
  Dim bSequence As Boolean
  Dim lSpare As Long
  sFile = SourceFile
  lSize = FileLen(sFile) - 1
  ReDim byB(lSize)
  iFile = FreeFile
  Open sFile For Binary As iFile
  Get #iFile, , byB()
  Close iFile
  ' Get the header information
  lCount = 0

  Do While lCount < lSize And lCount < 3
    sType = sType & Chr$(byB(lCount))
    lCount = lCount + 1
  Loop

  If sType <> "qCx" Then
    ' Not compressed - return sourcefile to destination
    GoTo WriteExpanded
  End If

  ' Check if sequence is included

  If byB(4) = 1 Then
    nSeqChar = byB(5)
    bSequence = True
  Else
    nSeqChar = -2
    bSequence = False
  End If

  ' Number of spare characters
  lSpare = byB(6)
  ' Get original file size
  sType = ""
  lCount = 7

  Do While lCount < lSize And lCount < 11
    sType = sType & Chr$(byB(lCount))
    lCount = lCount + 1
  Loop

  lOrigSize = GetBinaryLong(sType)
  'ReDim byNew(lOrigSize)
  lSpare = lSpare * 3 + lCount

  Do While lCount < lSpare
    qReturn(byB(lCount)).Char1 = byB(lCount + 1)
    qReturn(byB(lCount)).Char2 = byB(lCount + 2)
    qReturn(byB(lCount)).Total = 1
    lCount = lCount + 3
  Loop

  ReDim byNew(lSize - lCount)
  lNewPos = 0

  Do While lCount <= lSize
    byNew(lNewPos) = byB(lCount)
    lNewPos = lNewPos + 1
    lCount = lCount + 1
  Loop

  lNewPos = lNewPos - 1
  ReDim byB(lNewPos)
  byB() = byNew()
  ReDim byNew(lOrigSize)
  ' Replace pairs
  lSize = lNewPos
  lNewPos = 0
  lCount = 0

  Do While lCount <= lSize

    If byB(lCount) = nSeqChar Then
      byNew(lNewPos) = byB(lCount)
      byNew(lNewPos + 1) = byB(lCount + 1)
      byNew(lNewPos + 2) = byB(lCount + 2)
      lNewPos = lNewPos + 3
      lCount = lCount + 2
    ElseIf qReturn(byB(lCount)).Total > 0 Then
      byNew(lNewPos) = qReturn(byB(lCount)).Char1
      byNew(lNewPos + 1) = qReturn(byB(lCount)).Char2
      lNewPos = lNewPos + 2
    Else
      byNew(lNewPos) = byB(lCount)
      lNewPos = lNewPos + 1
    End If

    lCount = lCount + 1
  Loop

  ' Replace sequences if required
  ReDim byB(lOrigSize)
  byB() = byNew()

  If bSequence Then
    lSize = lNewPos - 1
    lCount = 0
    lNewPos = 0

    Do While lCount <= lSize

      If byB(lCount) = nSeqChar Then
        lRun = byB(lCount + 1)
        bLast = byB(lCount + 2)
        lCount = lCount + 2
        lIns = 1

        Do While lIns <= lRun
          byNew(lNewPos) = bLast
          lNewPos = lNewPos + 1
          lIns = lIns + 1
        Loop

      Else
        byNew(lNewPos) = byB(lCount)
        lNewPos = lNewPos + 1
      End If

      lCount = lCount + 1
    Loop

    ReDim byB(lNewPos - 1)
    byB() = byNew()
WriteExpanded:
    ' Check if destination exists and delete

    If Len(Dir$(DestinationFile)) > 0 Then
      Kill DestinationFile
    End If

  End If

  ' Save expanded file
  Open DestinationFile For Binary As #iFile
  Put #iFile, , byB()
  Close iFile
End Function

Public Sub qCompress(ByRef SourceFile As String, ByRef DestinationFile As String)
  'Variables: Count/Size
  Dim lCount As Long, lSize As Long, lItem As Long
  Dim byB() As Byte
  Dim lPair(255, 255) As Long
  Dim iFile As Integer
  Dim lChar(255) As Long
  Dim bLast As Byte
  Dim lRun As Long, lMaxRun As Long
  Dim lCurReplace As Long
  Dim qRep() As qtCompressRunReplaceType
  Dim qSeq() As qtCompressSequenceType
  Dim lCurSeq As Long
  Dim lSpare As Long
  Dim byNew() As Byte
  Dim lNewPos As Long
  Dim lIns As Long
  Dim bSpare() As Byte
  Dim bReplace As Boolean
  Dim lOrigSize As Long
  Dim lTotal As Long
  Dim sSize As String
  Dim qTemp As qtCompressSequenceType
  Dim nPrevious As Integer
  Dim bSequence As Boolean
  Dim nPairStart As Integer
  Dim nSeqSpare As Integer
  On Error GoTo CompressError
  ' Get sizes and open source file
  lSize = FileLen(SourceFile) - 1
  lOrigSize = lSize
  ReDim byB(lSize)
  ReDim byNew(lSize)
  iFile = FreeFile
  Open SourceFile For Binary As iFile
  Get #iFile, , byB()
  Close iFile
  ' Analyse characters and sequences
  bLast = byB(0)
  lChar(bLast) = 1

  For lCount = 1 To lSize
    lChar(byB(lCount)) = lChar(byB(lCount)) + 1

    If byB(lCount) = bLast Then
      lRun = lRun + 1

      If lRun = 3 Then
        lCurReplace = lCurReplace + 1
        ReDim Preserve qRep(lCurReplace)
        qRep(lCurReplace).Position = lCount - lRun
        qRep(lCurReplace).CharCode = bLast
        qRep(lCurReplace).Total = lRun
      ElseIf lRun > 3 Then
        qRep(lCurReplace).Total = lRun
      End If

    Else
      lRun = 0
    End If

    bLast = byB(lCount)
  Next

  ' Check spare characters
  lSpare = 0

  For lCount = 0 To 255

    If lChar(lCount) = 0 Then
      lSpare = lSpare + 1
      ReDim Preserve bSpare(lSpare)
      bSpare(lSpare) = lCount
    End If

  Next

  If lSpare = 0 Then
    ' No spare characters
    GoTo WriteCompressed
  End If

  ' If Sequences found, replace with sequence information

  If lCurReplace > 0 Then
    nSeqSpare = bSpare(1)
    nPairStart = 2
    bSequence = True
    lCount = 0
    lRun = 1
    lMaxRun = lCurReplace

    Do While lCount <= lSize

      If lRun < lCurReplace Then

        If lCount = qRep(lRun).Position Then
          lIns = qRep(lRun).Total + 1

          Do While lIns > 255
            byNew(lNewPos) = nSeqSpare
            byNew(lNewPos + 1) = 255
            byNew(lNewPos + 2) = qRep(lRun).CharCode
            lNewPos = lNewPos + 3
            lIns = lIns - 255
          Loop

          If lIns > 0 Then
            byNew(lNewPos) = nSeqSpare
            byNew(lNewPos + 1) = lIns
            byNew(lNewPos + 2) = qRep(lRun).CharCode
            lNewPos = lNewPos + 3
          End If

          lCount = lCount + qRep(lRun).Total
          lRun = lRun + 1
        Else
          byNew(lNewPos) = byB(lCount)
          lNewPos = lNewPos + 1
        End If

      Else
        byNew(lNewPos) = byB(lCount)
        lNewPos = lNewPos + 1
      End If

      lCount = lCount + 1
    Loop

    lNewPos = lNewPos - 1
    ReDim Preserve byNew(lNewPos)
    ReDim byB(lNewPos)
    byB() = byNew()
  Else
    ' No sequences
    lNewPos = lSize
    bSequence = False
    nPairStart = 1
    nSeqSpare = -2
  End If

  ' Analyse character pairs
  bLast = byB(0)
  lCount = 1

  Do While lCount < lNewPos

    If bLast = nSeqSpare Then
      lCount = lCount + 2
    ElseIf byB(lCount) <> nSeqSpare Then
      lPair(bLast, byB(lCount)) = lPair(bLast, byB(lCount)) + 1
    End If

    bLast = byB(lCount)
    lCount = lCount + 1
  Loop

  ' Sort character pairs
  ReDim qSeq(1)
  qSeq(1).Total = 2
  lCurSeq = 1

  For lCount = 0 To 255

    For lRun = 0 To 255

      If lPair(lCount, lRun) > qSeq(lCurSeq).Total Then

        If lCurSeq <= lSpare - nPairStart Then
          lCurSeq = lCurSeq + 1
          ReDim Preserve qSeq(lCurSeq)
        End If

        With qSeq(lCurSeq)
          .Char1 = lCount
          .Char2 = lRun
          .Total = lPair(lCount, lRun)
        End With

        lItem = lCurSeq

        Do While lItem > 1

          If qSeq(lItem).Total > qSeq(lItem - 1).Total Then
            qTemp = qSeq(lItem - 1)
            qSeq(lItem - 1) = qSeq(lItem)
            qSeq(lItem) = qTemp
            lItem = lItem - 1
          Else
            lItem = 0
          End If

        Loop

      End If

      lPair(lCount, lRun) = 0
    Next

  Next

  ' Set up pair information

  For lCount = 1 To lCurSeq

    With qSeq(lCount)
      lPair(.Char1, .Char2) = lCount + nPairStart - 1
    End With

  Next 'lCount

  If lSpare > (lCurSeq + nPairStart - 1) Then
    lSpare = lCurSeq
  End If

  ' Replace prevalent pairs
  lSize = lNewPos
  lNewPos = 0
  bLast = byB(0)
  lCount = 1
  nPrevious = -1

  Do While lCount <= lSize
    lRun = lPair(bLast, byB(lCount))

    If bReplace Then
      bReplace = False
      nPrevious = bLast
      bLast = byB(lCount)
    ElseIf nPrevious = nSeqSpare Then
      byNew(lNewPos) = bLast
      byNew(lNewPos + 1) = byB(lCount)
      nPrevious = bLast
      bLast = byB(lCount)
      lNewPos = lNewPos + 2
      bReplace = True
    ElseIf lRun > 0 And Not bReplace Then
      byNew(lNewPos) = bSpare(lRun)
      lNewPos = lNewPos + 1
      bReplace = True
      nPrevious = bLast
      bLast = bSpare(lRun)
    Else
      byNew(lNewPos) = bLast
      lNewPos = lNewPos + 1
      nPrevious = bLast
      bLast = byB(lCount)
    End If

    lCount = lCount + 1
  Loop

  If Not bReplace Then
    byNew(lNewPos) = bLast
  Else
    lNewPos = lNewPos - 1
  End If

  ReDim Preserve byNew(lNewPos)
  ReDim byB((lCurSeq * 3) + 11 + lNewPos)
  ' Create compress header
  byB(0) = 113       ' q
  byB(1) = 67        ' C
  byB(2) = 120       ' x
  byB(3) = 1         ' Compression version

  If bSequence Then  ' If sequence character was used:
    byB(4) = 1         ' Sequence used
    byB(5) = nSeqSpare ' Sequence character
  Else
    byB(4) = 0         ' No sequence
    byB(5) = 0         ' No sequence
  End If

  byB(6) = lCurSeq   ' Number of Pair replacement characters
  lRun = 7
  ' Create binary original file size
  sSize = SetBinaryLong(lOrigSize)
  ' Add binary original file size

  For lCount = 1 To 4
    byB(lRun) = Asc(Mid$(sSize, lCount, 1))
    lRun = lRun + 1
  Next 'lCount

  ' Add pair replacement characters

  For lCount = 1 To lCurSeq
    byB(lRun) = bSpare(lCount + nPairStart - 1)
    byB(lRun + 1) = qSeq(lCount).Char1
    byB(lRun + 2) = qSeq(lCount).Char2
    lRun = lRun + 3
  Next 'lCount

  ' Add compressed info to header

  For lCount = 0 To lNewPos
    byB(lRun) = byNew(lCount)
    lRun = lRun + 1
  Next

WriteCompressed:

  If lSpare = 0 And DestinationFile = SourceFile Then
    Exit Sub
  End If

  ' Check and remove DestinationFile if already exists

  If Len(Dir$(DestinationFile)) > 0 Then
    Kill DestinationFile
  End If

  ' Save compressed file
  Open DestinationFile For Binary As #iFile
  Put #iFile, , byB()
  Close iFile
  Exit Sub
CompressError:
  Debug.Assert False
End Sub

Public Function SetBinaryLong(ByVal lNumber As Long) As String
  Dim lCount As Long
  Dim sReturn As String
  Dim lVal(4) As Long
  lVal(1) = (lNumber And &HFF000000) \ &H1000000
  lVal(2) = (lNumber And &HFF0000) \ &H10000
  lVal(3) = (lNumber And &HFF00&) \ &H100
  lVal(4) = (lNumber And &HFF&)

  For lCount = 1 To 4
    sReturn = sReturn & Chr$(lVal(lCount))
  Next

  SetBinaryLong = sReturn
End Function

Public Function GetBinaryLong(ByVal sText As String) As Long
  Dim lReturn As Long
  Dim lVal(4) As Long
  Dim lCount As Long

  For lCount = 1 To 4
    lVal(lCount) = Asc(Mid$(sText, lCount, 1))
  Next

  lReturn = lReturn + lVal(1) * &H1000000
  lReturn = lReturn + lVal(2) * &H10000
  lReturn = lReturn + lVal(3) * &H100
  lReturn = lReturn + lVal(4)
  GetBinaryLong = lReturn
End Function

Public Sub GetBlockSize(ByRef sText As String, _
                        ByRef sWidth As Single, sDefaultWidth As Single, _
                        ByRef sHeight As Single, ByRef mvarWidth As Single, _
                        ByRef iLines As Integer, _
                        ByRef mvarLineH() As Single, ByRef mvarPosition() As Single, _
                        ByRef mvarFontName As String, ByRef mvarFontColor As Long, ByRef mvarFontSize As Single)
  '
  '  Dim sSizeX As Single
  '  Dim sLine As String, sChar As String, sWord As String
  '  Dim sLineW As Single, sLineH As Single, sWordW As Single
  '  Dim lCount As Long
  '  Dim eCharType As qePrinterChar, eEnd As qePrinterChar
  '  Dim bNewLine As Boolean
  '  Dim sFormat As String, lFPos As Long
  '  Dim bCheck As Boolean
  '  Dim bForceSameLine As Boolean
  '  Dim sTempLIndent As Single
  '  Dim sTempRIndent As Single
  '
  '    ReDim mvarLineH(0)
  '    ReDim mvarPosition(0)
  '    If sWidth <= 0 Then
  '      Exit Sub
  '
  '    End If
  '    With Printer
  '      sLine = ""
  '      sWord = ""
  '      sLineH = 0: sWordW = 0: sLineW = 0
  '      eEnd = eNone
  '      lCount = 1
  '      Do
  '
  '        bNewLine = False
  '
  '        Do
  '
  '          Do
  '
  '            eCharType = eNone
  '            bCheck = True
  '            sChar = Mid$(sText, lCount, 1)
  '  ' Check for formatting codes
  '            If sChar = "<" Then
  '              lFPos = InStr(lCount, sText, ">")
  '              If lFPos > 0 Then
  '                sFormat = Mid$(sText, lCount + 1, lFPos - lCount - 1)
  '                sFormat = UCase$(sFormat)
  '                bCheck = False
  '                If Len(sFormat) < 3 Or Left$(sFormat, 1) = "/" Then
  '
  '                  Select Case sFormat
  '
  '                    Case "B"
  '                      Printer.Font.Bold = True
  '
  '                    Case "U"
  '                      Printer.Font.Underline = True
  '
  '                    Case "I"
  '                      Printer.Font.Italic = True
  '
  '                    Case "/B"
  '                      Printer.Font.Bold = False
  '
  '                    Case "/I"
  '                      Printer.Font.Italic = False
  '
  '                    Case "/U"
  '                      Printer.Font.Underline = False
  '
  '                    Case "/FONT"
  '                      Printer.Font.Name = mvarFontName
  '
  '                    Case "/COLOR"
  '                      Printer.ForeColor = mvarFontColor
  '
  '                    Case "/SIZE"
  '                      Printer.Font.Size = mvarFontSize
  '
  '                    Case "/ALIGN"
  '  ' Do nothing
  '
  '                    Case "/LINDENT"
  '                      sWidth = sDefaultWidth - sTempRIndent
  '                      sTempLIndent = 0
  '
  '                    Case "/RINDENT"
  '                      sWidth = sDefaultWidth - sTempLIndent
  '                      sTempRIndent = 0
  '
  '                    Case Else
  '                      lFPos = lCount
  '                      bCheck = True
  '
  '                  End Select
  '
  '                ElseIf Left$(sFormat, 5) = "FONT=" Then
  '                  Printer.Font.Name = Property_Extract(sFormat)
  '                ElseIf Left$(sFormat, 6) = "COLOR=" Then
  '                  Printer.ForeColor = CLng(Property_Extract(sFormat))
  '                ElseIf Left$(sFormat, 5) = "SIZE=" Then
  '                  Printer.Font.Size = Val(Property_Extract(sFormat))
  '                ElseIf Left$(sFormat, 6) = "ALIGN=" Then
  '  ' do nothing
  '
  '                ElseIf Left$(sFormat, 8) = "LINDENT=" Then
  '                  sTempLIndent = ConvertToTwip(mvarScaleMode, Val(Property_Extract(sFormat)))
  '                  sWidth = sDefaultWidth - sTempLIndent - sTempRIndent
  '                  bForceSameLine = True
  '                ElseIf Left$(sFormat, 8) = "RINDENT=" Then
  '                  sTempRIndent = ConvertToTwip(mvarScaleMode, Val(Property_Extract(sFormat)))
  '                  sWidth = sDefaultWidth - sTempLIndent - sTempRIndent
  '                  bForceSameLine = True
  '                ElseIf Left$(sFormat, 5) = "FORCE" Then
  '                  bForceSameLine = True
  '                ElseIf Left$(sFormat, 4) = "TAB=" Then
  '                  bForceSameLine = True
  '  'Property_ExtractTab sFormat, sTempLIndent
  '                Else
  '                  bCheck = False
  '
  '                End If
  '
  '                If Not bCheck Then
  '                  lCount = lFPos
  '
  '                End If
  '
  '              End If
  '
  '            End If
  '
  '  ' CHARACTER CHECK: Look for potential line breaks or where text
  '  '                  width is greater than boundary
  '            If bCheck Then
  '
  '              Select Case sChar
  '
  '                Case " "
  '                  eCharType = eSpace
  '
  '                Case "-"
  '                  sSizeX = sLineW + sWordW + .TextWidth(sChar)
  '                  If .TextHeight(sChar) > sLineH Then
  '                    sLineH = .TextHeight(sChar)
  '
  '                  End If
  '
  '                  If sSizeX > sWidth Then
  '                    eCharType = eOops
  '                  Else
  '                    eCharType = eDash
  '
  '                  End If
  '
  '                Case vbLf
  '                  sChar = ""
  '                  eCharType = eLine
  '
  '                Case vbCr
  '                  If lCount < Len(sText) Then
  '                    If Mid$(sText, lCount + 1, 1) = vbLf Then
  '                      lCount = lCount + 1
  '
  '                    End If
  '
  '                  End If
  '
  '                  sChar = ""
  '                  eCharType = eLine
  '
  '                Case Else
  '  ' CHARACTER CHECK: See if addition of character makes line too long
  '                  sSizeX = sLineW + sWordW + .TextWidth(sChar)
  '                  If .TextHeight(sChar) > sLineH Then
  '                    sLineH = .TextHeight(sChar)
  '
  '                  End If
  '
  '                  If sSizeX > sWidth Then
  '                    eCharType = eOops
  '                  Else
  '                    sWord = sWord & sChar
  '                    sWordW = sWordW + .TextWidth(sChar)
  '
  '                  End If
  '
  '              End Select
  '
  '            End If
  '
  '            lCount = lCount + 1
  '
  '          Loop While lCount <= Len(sText) And eCharType = eNone And Not bForceSameLine
  '
  '          If bForceSameLine Then
  '            eCharType = eLine
  '
  '          End If
  '
  '  ' LINE SPLIT: Examine potential line break
  '          If lCount > Len(sText) Then
  '            eCharType = eLine
  '
  '          End If
  '
  '          Select Case eCharType
  '
  '            Case qePrinterChar.eNone
  '              sLine = sLine & sWord
  '              sLineW = sLineW + sWordW
  '              eEnd = eLine
  '
  '            Case qePrinterChar.eOops
  '              If eEnd = eNone Then
  '                sLine = sWord
  '                sLineW = sLineW + sWordW
  '                sWord = sChar
  '                sWordW = .TextWidth(sChar)
  '              Else
  '                sLine = Trim$(sLine)
  '                sWord = sWord & sChar
  '                sWordW = sWordW + .TextWidth(sChar)
  '
  '              End If
  '
  '              bNewLine = True
  '
  '            Case qePrinterChar.eDash, qePrinterChar.eSpace
  '              eEnd = eCharType
  '              sLine = sLine & sWord & sChar
  '              sLineW = sLineW + sWordW + .TextWidth(sChar)
  '              If sWordW > msLargeWord Then
  '                msLargeWord = sWordW
  '
  '              End If
  '
  '              sWord = ""
  '              sWordW = 0
  '
  '            Case qePrinterChar.eLine
  '              If sLineH = 0 Then
  '                sLineH = Printer.TextHeight("H")
  '
  '              End If
  '
  '              sLine = sLine & sWord
  '              sLineW = sLineW + sWordW
  '              If sWordW > msLargeWord Then
  '              msLargeWord = sWordW
  '              End If
  '              eEnd = eLine
  '              sWord = ""
  '              sWordW = 0
  '              bNewLine = True
  '
  '          End Select
  '
  '  ' LINE SPLIT: Add new line if required
  '          If bNewLine Then
  '            If Not bForceSameLine Then
  '
  '              iLines = iLines + 1
  '              ReDim Preserve mvarLineH(iLines)
  '              If sLineW > mvarWidth Then
  '                mvarWidth = sLineW
  '
  '              End If
  '
  '              mlTextWidth = mlTextWidth + sLineW
  '  ' *** Changed 1.6.0
  '              mvarLineH(iLines) = sLineH + mvarLineSpacing
  '              sHeight = sHeight + sLineH + mvarLineSpacing
  '  ' *** End Change 1.6.0
  '              sLineH = 0
  '
  '            End If
  '
  '            sLine = ""
  '            sLineW = 0
  '            eEnd = eNone
  '            bForceSameLine = False
  '
  '          End If
  '
  '        Loop While Not bNewLine
  '
  '      Loop While lCount <= Len(sText)
  '    End With
  '
  '
End Sub

