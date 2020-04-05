Attribute VB_Name = "CRC32"
'Executable Tamper-Protection: CRC32 Checksum Validation
'by Detonate (detonate@start.com.au)

'CRC32 code by Neo (http://vbcode.8m.com/)
'His original CRC32 code can be found at http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=4822
'The CRC32 routines of his are very high speed and this is an excellent algorithm to use for this program

'How it works:
'Simply add CRC32.BAS module to your VB project, and activate it at the start of your program like this:
' If IntegrityOK <> 1 Then Msgbox "CRC32 checksum failed"
'You must run the APPLYCRC32 program over your exe before you run it. APPLYCRC32 reads the file, calculates
'a checksum, and appends the checksum to the end of the exe file. When youre exe file calls IntegrityOK(),
'it reads the last 8 bytes of its own file, and if the two checksums match, then you know your file hasn't
'been tampered with.

'CRC32 is a very effective (and fast) checksum calculation, and this is a great way of preventing crackers
'from at least writing software patches (those little programs which change a few bytes in order to do
'things the programmers didnt want them to do :-)

Option Explicit
Option Compare Text
'// Then declare this array variable Crc32Table
Private Crc32Table(255) As Long
'// Then all we have to do is writing public functions like these...


Public Function InitCrc32(Optional ByVal Seed As Long = &HEDB88320, Optional ByVal Precondition As Long = &HFFFFFFFF) As Long
    '// Declare counter variable iBytes, counter variable iBits, value variables lCrc32 and lTempCrc32
    Dim iBytes As Integer, iBits As Integer, lCrc32 As Long, lTempCrc32 As Long
    '// Turn on error trapping
    On Error Resume Next
    '// Iterate 256 times

    For iBytes = 0 To 255
        '// Initiate lCrc32 to counter variable
        lCrc32 = iBytes
        '// Now iterate through each bit in counter byte


        For iBits = 0 To 7
            '// Right shift unsigned long 1 bit
            lTempCrc32 = lCrc32 And &HFFFFFFFE
            lTempCrc32 = lTempCrc32 \ &H2
            lTempCrc32 = lTempCrc32 And &H7FFFFFFF
            '// Now check if temporary is less than zero and then mix Crc32 checksum with Seed value


            If (lCrc32 And &H1) <> 0 Then
                lCrc32 = lTempCrc32 Xor Seed
            Else
                lCrc32 = lTempCrc32
            End If
        Next
        '// Put Crc32 checksum value in the holding array
        Crc32Table(iBytes) = lCrc32
    Next
    '// After this is done, set function value to the precondition value
    InitCrc32 = Precondition
End Function
'// The function above is the initializing function, now we have to write the computation function


Public Function AddCrc32(ByVal Item As String, ByVal CRC32 As Long) As Long
    '// Declare following variables
    Dim bCharValue As Byte, iCounter As Integer, lIndex As Long
    Dim lAccValue As Long, lTableValue As Long
    '// Turn on error trapping
    On Error Resume Next
    '// Iterate through the string that is to be checksum-computed


    For iCounter = 1 To Len(Item)
        '// Get ASCII value for the current character
        bCharValue = Asc(Mid$(Item, iCounter, 1))
        '// Right shift an Unsigned Long 8 bits
        lAccValue = CRC32 And &HFFFFFF00
        lAccValue = lAccValue \ &H100
        lAccValue = lAccValue And &HFFFFFF
        '// Now select the right adding value from the holding table
        lIndex = CRC32 And &HFF
        lIndex = lIndex Xor bCharValue
        lTableValue = Crc32Table(lIndex)
        '// Then mix new Crc32 value with previous accumulated Crc32 value
        CRC32 = lAccValue Xor lTableValue
    Next
    '// Set function value the the new Crc32 checksum
    AddCrc32 = CRC32
End Function
'// At last, we have to write a function so that we can get the Crc32 checksum value at any time


Public Function GetCrc32(ByVal CRC32 As Long) As Long
    '// Turn on error trapping
    On Error Resume Next
    '// Set function to the current Crc32 value
    GetCrc32 = CRC32 Xor &HFFFFFFFF
End Function
'// To Test the Routines Above...


Public Function Compute(ToGet As String) As String
    Dim lCrc32Value As Long
    On Error Resume Next
    lCrc32Value = InitCrc32()
    lCrc32Value = AddCrc32(ToGet, lCrc32Value)
    Compute = Hex$(GetCrc32(lCrc32Value))
End Function
Private Function AppExe() As String
On Error Resume Next
Dim AP As String
AP = App.Path
If Right(AP, 1) <> "\" Then AP = AP & "\"
AppExe = AP & App.EXEName & ".exe"
End Function
Public Function IntegrityOK() As Integer
'Returns:
'  -1   =   No CRC found at the end of the file :-/
'  -2   =   File CRC and Real CRC dont match :-/
'  1    =   Both CRCs match - file is ok! :-)
    Dim lCrc32Value As Long
    Dim CRCStr As String * 8
    Dim FL As Long  'file length
    On Error Resume Next
    Dim FileStr$
    FL = FileLen(AppExe)
    MsgBox "FL=" & FL
    FileStr$ = String(FL - 8, 0)
    Open AppExe For Binary As #1
     Get #1, 1, FileStr$
     Get #1, FL - 7, CRCStr
    Close #1
    If Trim(CRCStr) = "" Or Trim(CRCStr) = String(8, 0) Then
       IntegrityOK = -1
       Exit Function
    End If
    lCrc32Value = InitCrc32()
    lCrc32Value = AddCrc32(FileStr$, lCrc32Value)
    Dim RealCRC As String
    RealCRC = CStr(Hex$(GetCrc32(lCrc32Value)))
    MsgBox "Real CRC=" & RealCRC & vbCrLf & "File CRC=" & CRCStr, vbInformation + vbOKOnly, "CRC32 Results for " & AppExe
    If RealCRC <> CRCStr Or Trim(CRCStr) = "" Or Trim(CRCStr) = String(8, 0) Then
       IntegrityOK = -2
       Exit Function
    Else
       IntegrityOK = 1
       Exit Function
    End If
End Function

