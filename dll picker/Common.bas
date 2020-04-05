Attribute VB_Name = "Common"
Option Explicit
Option Base 1

Public gCodeColor As String
Public gSelectColor As Long
Public pRows As Integer, pCols As Integer

Type ColorPickerHeader
  sName As String
  sVersion As String
  sCopyright As String
  iCount As Integer
End Type

Type RIFFPaletteHeader
  RIFF As String * 4
  Reserved(1 To 18) As Byte
  Cols As Integer
End Type

Const COLORPICKER = "Color Picker Palette"
Const VERSION = "Version 1.0 2004"
Const COPYRIGHT = "Copyright (c) 2004 by Haidau Alin alin78hai@yahoo.com"

Const JASCPAL = "JASC-PAL"
Const JASCPAL1 = "0100"

Const HOMESITE = "Palette"
Const HOMESITE1 = "Version 3.0"
Const HOMESITE2 = "-----------"

Enum ColorControlsPaletteFormats
  ccColorPicker
  ccJASC
  ccHomesite
End Enum

Public m_oColors() As Long
Public m_oClrNames() As String
Public m_oCustClrs() As Long

Public m_lDefault As Long
Public m_sLastPal As String
Public m_lBoxSize As Integer
Public m_lSpace As Integer

Public m_iRows As Byte
Public m_iCols As Byte
Public m_iPaletteType As Integer

'Public Sub Timer(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
'    Call frmColorPalette.TipTimer(hwnd, uMsg, idEvent, dwTime)
'End Sub

Public Function RGB2Hex(lCdlColor As Long) As String
    Dim lCol As Long
    Dim iRed, iGreen, iBlue As Integer
    Dim vHexR, vHexG, vHexB As Variant
    
    'Break out the R, G, B values from the common dialog color
    lCol = lCdlColor
    iRed = lCol Mod &H100
    lCol = lCol \ &H100
    iGreen = lCol Mod &H100
    lCol = lCol \ &H100
    iBlue = lCol Mod &H100
    
    'Determine Red Hex
    vHexR = Hex(iRed)

    If Len(vHexR) < 2 Then
      vHexR = "0" & vHexR
    End If

    'Determine Green Hex
    vHexG = Hex(iGreen)
    If Len(vHexG) < 2 Then
      vHexG = "0" & iGreen
    End If

    'Determine Blue Hex
    vHexB = Hex(iBlue)
    If Len(vHexB) < 2 Then
      vHexB = "0" & vHexB
    End If
    'Add it up, return the function value
    RGB2Hex = "#" & vHexR & vHexG & vHexB
End Function

Public Function HexToLong(sHexColor As String) As Long
  Dim lCol As Long, i, N
  
  If Left(sHexColor, 1) = "#" Then sHexColor = Mid(sHexColor, 2)
  sHexColor = UCase(sHexColor)
    
  For i = 1 To Len(sHexColor) Step 2
    lCol = lCol + Dec(Mid(sHexColor, i, 2)) * 256 ^ N
    N = N + 1
  Next i
  HexToLong = lCol
End Function

Public Function Dec(ByVal sHex As String) As Long 'Converts Hex to Decimal
    Const HVal = "0123456789ABCDEF"
    Dim iPos As Byte, i As Integer, lDec As Long
    Dim L As Integer, x As Byte
    L = Len(sHex)
    If L > 255 Then Exit Function
    lDec = 0
    For i = L To 1 Step -1
        x = InStr(1, HVal, Mid(sHex, i, 1), vbTextCompare)
        If x = 0 Then Exit Function Else x = x - 1
        lDec = lDec + x * 16 ^ (L - i)
    Next i
    Dec = lDec
End Function

Public Sub DrawRect(hdc As Long, R As RECT, Optional LightColor As Long = vbScrollBars, Optional DarkColor As Long = vbButtonShadow, Optional FillColor As Long = vbButtonFace, Optional bNoFill As Boolean = False)
  Dim hBr As Long
  Dim tJunk As PointAPI
  Dim hPen As Long, hPenOld As Long
    
  If Not bNoFill Then
    hBr = CreateSolidBrush(VBClr(FillColor))
    Call FillRect(hdc, R, hBr)
    Call DeleteObject(hBr)
  End If
  hPen = CreatePen(PS_SOLID, 1, VBClr(LightColor))
  hPenOld = SelectObject(hdc, hPen)
  MoveToEx hdc, R.Left, R.Top, tJunk
  LineTo hdc, R.Right, R.Top
  MoveToEx hdc, R.Left, R.Top, tJunk
  LineTo hdc, R.Left, R.Bottom
  Call DeleteObject(hPen)
  Call DeleteObject(hPenOld)
      
  hPen = CreatePen(PS_SOLID, 1, VBClr(DarkColor))
  hPenOld = SelectObject(hdc, hPen)
  MoveToEx hdc, R.Right, R.Top, tJunk
  LineTo hdc, R.Right, R.Bottom
  LineTo hdc, R.Left, R.Bottom
  Call DeleteObject(hPen)
  Call DeleteObject(hPenOld)
End Sub

' Converts a OLE_COLOR to COLORREF
Public Function VBClr(ByVal clr As Long) As Long
  OleTranslateColor clr, 0, VBClr
End Function

' Load a palette file in HOMESITE, JASC or COLORPICKER format.
Public Function LoadPalette(ByVal Filename As String, ByRef ClrNames() As String) As Long
  Dim lFile As Long, lIdx As Integer
  Dim RIFFHdr As RIFFPaletteHeader
  Dim hdr As ColorPickerHeader, lNameLen As Long
  Dim Char As Byte, lClrCount As Long

  'Get a free file handle
  lFile = FreeFile()
  
  
  'MsgBox Filename
  
  ' Open the palette
  Open Filename For Binary As lFile
  ' Read the RIFF header
  Get lFile, , RIFFHdr
  Close lFile

  Select Case RIFFHdr.RIFF
    Case "RIFF"
      ' If the palette header starts with RIFF then
      ' it'Char a Microsoft palette file
      ' Read the palette
      Open Filename For Binary As lFile
      lClrCount = RIFFHdr.Cols
      ReDim m_oColors(0 To RIFFHdr.Cols - 1)
      ReDim ClrNames(0 To RIFFHdr.Cols - 1)
      ' Skip the header
      Seek lFile, 25
      ' Get the Colors
      For lIdx = 0 To RIFFHdr.Cols - 1
        Get lFile, , m_oColors(lIdx)
      Next
      ' Close the file
      Close lFile
    Case Else
      ' The file is either a JASC palette, a ColorPicker or a Homesite palette
      Dim Lne As String, r1 As Long, G As Long, B As Long
      ' Open the file
      Open Filename For Input As lFile
      ' Get the first line
      Line Input #lFile, Lne
      ' Check if it's a JASC palette
      If UCase$(Left$(Lne, Len(JASCPAL))) = JASCPAL Then
        ' Skip the next line
        Line Input #lFile, Lne
        ' Read the color count
        Line Input #lFile, Lne
        ' Get the color count value
        lClrCount = Val(Lne)
        If lClrCount <= 0 Or lClrCount > 256 Then
          ' Close the file
          Close lFile
          ' Raise an error
          Err.Raise vbObjectError + 2, , "Invalid color count in palette file."
        Else
          ReDim m_oColors(0 To lClrCount - 1)
          ReDim ClrNames(0 To lClrCount)
          ' Read the colors
          For lIdx = 0 To lClrCount - 1
            Line Input #lFile, Lne
            r1 = InStr(Lne, " ")
            G = InStr(r1 + 1, Lne, " ")
            B = Val(Mid$(Lne, G + 1))
            G = Val(Mid$(Lne, r1 + 1, G - r1 - 1))
            r1 = Val(Left$(Lne, r1 - 1))
            m_oColors(lIdx) = RGB(r1, G, B)
          Next
        End If
        ' Close the file
        Close lFile
      ElseIf UCase$(Left$(Lne, Len(HOMESITE))) = UCase$(HOMESITE) Then
        ' The file is a Homesite Palette
        ' Skip next 2 lines
        Line Input #lFile, Lne
        Line Input #lFile, Lne
        lClrCount = 0
        ' There's no color count in this format so read the file
        ' until EOF
        Do Until EOF(lFile)
          ReDim Preserve m_oColors(0 To lClrCount) As Long
          Line Input #lFile, Lne
          r1 = InStr(Lne, " ")
          G = InStr(r1 + 1, Lne, " ")
          B = Val(Mid$(Lne, G + 1))
          G = Val(Mid$(Lne, r1 + 1, G - r1 - 1))
          r1 = Val(Left$(Lne, r1 - 1))
          m_oColors(lClrCount) = RGB(r1, G, B)
          lClrCount = lClrCount + 1
        Loop
      ElseIf UCase$(Left$(Lne, Len(COLORPICKER))) = UCase$(COLORPICKER) Then
        'this is our own propertary format :P
        'skip the next line which containes the version of the file
        Line Input #lFile, Lne
        'skip the next line which containes our copyright info :))
        Line Input #lFile, Lne
        lClrCount = 0
        ' There's no color count in this format so read the file
        ' until EOF
        Do Until EOF(lFile)
          ReDim Preserve m_oColors(0 To lClrCount) As Long
          Line Input #lFile, Lne
          r1 = InStr(Lne, " ")
          G = InStr(r1 + 1, Lne, " ")
          B = Val(Mid$(Lne, G + 1))
          G = Val(Mid$(Lne, r1 + 1, G - r1 - 1))
          r1 = Val(Left$(Lne, r1 - 1))
          m_oColors(lClrCount) = RGB(r1, G, B)
          lClrCount = lClrCount + 1
        Loop
        ' Close the file
        Close lFile
      Else
        Close lFile
        Err.Raise vbObjectError + 1, , "Archivo de paleta no válido."
      End If
  End Select
  Close lFile
  pRows = lClrCount \ 18 + IIf(lClrCount Mod 18 > 0, 1, 0)
  pCols = IIf(lClrCount Mod 18 >= 0, 18, lClrCount Mod 18)
  
  LoadPalette = lClrCount
End Function

'*********************************************************************************************
' SavePalette
'
' Saves a palette file. Microsoft format is not supported.
'*********************************************************************************************
Public Sub SavePalette(ByVal Filename As String, ByVal Format As ColorControlsPaletteFormats, Clrs() As Long, Names() As String)
Dim lFile As Long, lIdx As Long, clr As Long, s As String
Dim Max As Long, hdr As ColorPickerHeader

   ' Get a free handle
   lFile = FreeFile

   Max = UBound(Clrs)

   ' Open the file for output
   Open Filename For Output As lFile

   Select Case Format

      Case ccColorPicker

         ' Close the file
         Close lFile

         ' Open the file in binary mode
         Open Filename For Binary As lFile

         ' Fill the header
         hdr.iCount = Max
         'hdr.Magic = ccMAGIC

         ' Write using the new version
         'hdr.sVersion = ccPALETTEVERSION101

         ' Save the header
         Put lFile, , hdr

         ' Write colors
         For lIdx = 0 To Max

            Put lFile, , Clrs(lIdx)

            Put lFile, , Len(Names(lIdx))
            Put lFile, , Names(lIdx)

         Next

      Case ccJASC

         ' Write the header
         Print #lFile, JASCPAL
         Print #lFile, JASCPAL1

         Select Case Max

            ' This format supports only 16
            ' and 256 colors

            Case Is <= 15

               ' Write color count
               Print #lFile, "16"

               For lIdx = 0 To 15

                  If lIdx > Max Then
                     clr = 0
                  Else
                     clr = VBClr(Clrs(lIdx))
                  End If

                  ' Write color components
                  Print #lFile, (clr And &HFF&) & " " & ((clr \ &H100&) And &HFF&) & " " & (clr \ &H10000)
               Next

            Case Is <= 255

               ' Write color count
               Print #lFile, "256"

               For lIdx = 0 To 255

                  If lIdx > Max Then
                     clr = 0
                  Else
                     clr = VBClr(Clrs(lIdx))
                  End If

                  ' Write color components
                  Print #lFile, (clr And &HFF&) & " " & ((clr \ &H100&) And &HFF&) & " " & (clr \ &H10000)

               Next

         End Select

      Case ccHomesite

         ' Write the header
         Print #lFile, HOMESITE
         Print #lFile, HOMESITE1
         Print #lFile, HOMESITE2

         For lIdx = 0 To Max

            clr = VBClr(Clrs(lIdx))

            ' Write color components
            Print #lFile, (clr And &HFF&) & " " & ((clr \ &H100&) And &HFF&) & " " & (clr \ &H10000)

         Next

   End Select

   ' Close the file
   Close lFile

End Sub

