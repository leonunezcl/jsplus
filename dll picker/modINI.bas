Attribute VB_Name = "modINI"
Option Explicit

Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

' ReadINIValue: reads the specified value from the inifile
' In:
'   Section name, Key name, INI file name
' Returns:
'   the string value read from the ini
Public Function ReadINIValue(sSection As String, sKey As String, sFileName As String) As String
  ' get out if we don't have a filename ....
  If sFileName = "" Then
    ReadINIValue = "error!"
    Exit Function
  End If
  Dim sRet As String
  sRet = String(10000, Chr(0))
  ReadINIValue = Left(sRet, GetPrivateProfileString(sSection, ByVal sKey, "", sRet, Len(sRet), sFileName))
End Function

' ReadSection: reads specified section at once
' In:
'   section name, key array, value array
' returns:
'   NOTE that both arrays must not be fixed dimension, else the redim will fail!!!
'   these 2 arrays will be redimmed as much as needed to contain all the keynames and values!
'-----------------------------------------------------------------------------------
Public Sub ReadINISection(sSection As String, sKeys() As String, sValues() As String, iCount As Long, sFileName As String)
  If sFileName = "" Then Exit Sub
  ' return value
  Dim sReturned As String * 32767 ' max chars allowed in Win95
  Dim lRet As Long
  
  lRet = GetPrivateProfileSection(sSection, sReturned, Len(sReturned), sFileName)
  
  ReDim sKeys(0)
  ReDim sValues(0)
  Dim iNull As Integer, iStart As Integer, i As Integer, s As String
  
  iCount = 0
  iStart = 1
  iNull = InStr(iStart, sReturned, vbNullChar)
  Do While iNull
    ReDim Preserve sKeys(iCount)
    ReDim Preserve sValues(iCount)
    s = Mid(sReturned, iStart, iNull - iStart)
    sKeys(iCount) = Left(s, InStr(1, s, "=") - 1)
    sValues(iCount) = Right(s, Len(s) - InStr(1, s, "="))
    iStart = iNull + 1
    iNull = InStr(iStart, sReturned, vbNullChar)
    ' lRet contains the numbers of chars copied to the buffer, so if iNull > lRet then we have it all...
    If iNull > lRet Then iNull = 0
    iCount = iCount + 1
  Loop
  iCount = iCount - 1
End Sub

'-----------------------------------------------------------------------------------
' WriteINIValue: writes the specified value to the ini file pointed to in the Filename property
' In:
'   Section name, Key name, Value, File name
'-----------------------------------------------------------------------------------
Public Sub WriteINIValue(sSection As String, sKey As String, sValue As String, sFileName As String)
  Call WritePrivateProfileString(sSection, sKey, sValue, sFileName)
End Sub

' WriteINISection: Write a section at once.
' In:
'   Section name, array of keys, array of values, File name
'   Both array must be of the same size, else nothing will be written
Public Sub WriteINISection(sSection As String, sKeys() As String, sValues() As String, sFileName As String)
  If sFileName = "" Then Exit Sub
  If UBound(sKeys) <> UBound(sValues) Then
        Exit Sub
  End If
  ' tempstring which will contain the value to write on this format: key=value+vbNullChar+key=value etc...
  Dim s As String, L As Long
  ' format the string to write
  For L = LBound(sKeys) To UBound(sKeys)
     s = s & sKeys(L) & "=" & sValues(L) & vbNullChar
  Next
  ' write section
  Call WritePrivateProfileSection(sSection, s, sFileName)
End Sub

Public Sub GetINISectionNames(ByRef sSectionNames() As String, ByRef iCount As Long, sFileName As String)
  
  If sFileName = "" Then Exit Sub
  
  Dim sReturned As String * 32767 ' max chars allowed in Win95
  Dim lRet As Long
  Erase sSectionNames
  iCount = 0
  lRet = GetPrivateProfileSectionNames(sReturned, Len(sReturned), sFileName)
  If lRet <> 0 Then
    Dim iNull As Integer, iStart As Integer, i As Integer
    iStart = 1
    iNull = InStr(iStart, sReturned, vbNullChar)
    Do While iNull
      ReDim Preserve sSectionNames(iCount)
      sSectionNames(iCount) = Mid(sReturned, iStart, iNull - iStart)
      iStart = iNull + 1
      iNull = InStr(iStart, sReturned, vbNullChar)
      ' lRet contains the numbers of chars copied to the buffer, so if iNull > lRet then we have it all...
      If iNull > lRet Then iNull = 0
      iCount = iCount + 1
    Loop
    iCount = iCount - 1
  End If
End Sub

Public Sub DeleteINIKey(sSection As String, sKeyName As String, sFileName)
  Call WritePrivateProfileString(sSection, sKeyName, 0&, sFileName)
End Sub

Public Sub DeleteINISection(sSection As String, sFileName As String)
  Call WritePrivateProfileString(sSection, 0&, 0&, sFileName)
End Sub

