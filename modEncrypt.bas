Attribute VB_Name = "Module4"
Option Explicit
Const Base64Chars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Function Base10ToBinary(ByVal Base10 As Long) As String
    Dim PrevResult As Integer, CurResult As Integer
    If Base10 = 0 Then
        Base10ToBinary = "0"
        Exit Function
    End If
    Do
        CurResult = Int(Log(Base10) / Log(2))
        If PrevResult = 0 Then PrevResult = CurResult + 1
        Base10ToBinary = Base10ToBinary & String(PrevResult - CurResult - 1, "0") & "1"
        Base10 = Base10 - 2 ^ CurResult
        PrevResult = CurResult
    Loop Until Base10 = 0
    Base10ToBinary = Base10ToBinary & String(CurResult, "0")
End Function

Function BinaryToBase10(ByVal Binary As String) As Long
    Dim I As Integer
    For I = Len(Binary) To 1 Step -1
        BinaryToBase10 = BinaryToBase10 + val(Mid(Binary, I, 1)) * 2 ^ (Len(Binary) - I)
    Next
End Function

Sub Bin3x8To4x6(ByVal Bin1Len8 As String, ByVal Bin2Len8 As String, ByVal Bin3Len8 As String, ByRef Bin1Len6 As String, ByRef Bin2Len6 As String, ByRef Bin3Len6 As String, ByRef Bin4Len6 As String)
    Bin1Len8 = VBA.Right$("0000000" & Bin1Len8, 8)
    Bin2Len8 = VBA.Right$("0000000" & Bin2Len8, 8)
    Bin3Len8 = VBA.Right$("0000000" & Bin3Len8, 8)
    Bin1Len6 = VBA.Left$(Bin1Len8, 6)
    Bin2Len6 = VBA.Right$(Bin1Len8, 2) & VBA.Left$(Bin2Len8, 4)
    Bin3Len6 = VBA.Right$(Bin2Len8, 4) & VBA.Left$(Bin3Len8, 2)
    Bin4Len6 = VBA.Right$(Bin3Len8, 6)
End Sub

Sub Bin4x6To3x8(ByVal Bin1Len6 As String, ByVal Bin2Len6 As String, ByVal Bin3Len6 As String, ByVal Bin4Len6 As String, ByRef Bin1Len8 As String, ByRef Bin2Len8 As String, ByRef Bin3Len8 As String)
    Bin1Len6 = VBA.Right$("00000" & Bin1Len6, 6)
    Bin2Len6 = VBA.Right$("00000" & Bin2Len6, 6)
    Bin3Len6 = VBA.Right$("00000" & Bin3Len6, 6)
    Bin4Len6 = VBA.Right$("00000" & Bin4Len6, 6)
    Bin1Len8 = Bin1Len6 & VBA.Left$(Bin2Len6, 2)
    Bin2Len8 = VBA.Right$(Bin2Len6, 4) & VBA.Left$(Bin3Len6, 4)
    Bin3Len8 = VBA.Right$(Bin3Len6, 2) & Bin4Len6
End Sub
Function RemoveFromString(ByVal TheString As String, ByVal WhatToRemove As String) As String
    Dim lPos As Long
    If Len(WhatToRemove) = 0 Then Exit Function
    lPos = InStr(TheString, WhatToRemove)
    While lPos > 0
        TheString = VBA.Left$(TheString, lPos - 1) & Mid(TheString, lPos + Len(WhatToRemove))
        lPos = InStr(TheString, WhatToRemove)
    Wend
    RemoveFromString = TheString
End Function
Function Base64Encode(ByVal NormalString As String, Optional ByVal Break As Integer = 0) As String
    Dim I As Integer, Bin1Len8 As String, Bin2Len8 As String, Bin3Len8 As String
    Dim Bin1Len6 As String, Bin2Len6 As String, Bin3Len6 As String, Bin4Len6 As String
    If NormalString = vbNullString Then Exit Function
    For I = 1 To Len(NormalString) - 3 Step 3
        Bin1Len8 = Base10ToBinary(Asc(Mid(NormalString, I, 1)))
        Bin2Len8 = Base10ToBinary(Asc(Mid(NormalString, I + 1, 1)))
        Bin3Len8 = Base10ToBinary(Asc(Mid(NormalString, I + 2, 1)))
        Call Bin3x8To4x6(Bin1Len8, Bin2Len8, Bin3Len8, Bin1Len6, Bin2Len6, Bin3Len6, Bin4Len6)
        Base64Encode = Base64Encode & Mid(Base64Chars, BinaryToBase10(Bin1Len6) + 1, 1)
        Base64Encode = Base64Encode & Mid(Base64Chars, BinaryToBase10(Bin2Len6) + 1, 1)
        Base64Encode = Base64Encode & Mid(Base64Chars, BinaryToBase10(Bin3Len6) + 1, 1)
        Base64Encode = Base64Encode & Mid(Base64Chars, BinaryToBase10(Bin4Len6) + 1, 1)
    Next
    NormalString = VBA.Right$(NormalString, Len(NormalString) - IIf(Len(NormalString) / 3 = Int(Len(NormalString) / 3), Len(NormalString) - 3, Int(Len(NormalString) / 3) * 3))
    Bin1Len8 = Base10ToBinary(Asc(VBA.Left$(NormalString, 1)))
    If Len(NormalString) >= 2 Then Bin2Len8 = Base10ToBinary(Asc(Mid(NormalString, 2, 1))) Else Bin2Len8 = "0"
    If Len(NormalString) = 3 Then Bin3Len8 = Base10ToBinary(Asc(VBA.Right$(NormalString, 1))) Else Bin3Len8 = "0"
    Call Bin3x8To4x6(Bin1Len8, Bin2Len8, Bin3Len8, Bin1Len6, Bin2Len6, Bin3Len6, Bin4Len6)
    Base64Encode = Base64Encode & Mid(Base64Chars, BinaryToBase10(Bin1Len6) + 1, 1)
    Base64Encode = Base64Encode & Mid(Base64Chars, BinaryToBase10(Bin2Len6) + 1, 1)
    Base64Encode = Base64Encode & IIf(Len(NormalString) >= 2, Mid(Base64Chars, BinaryToBase10(Bin3Len6) + 1, 1), "=")
    Base64Encode = Base64Encode & IIf(Len(NormalString) = 3, Mid(Base64Chars, BinaryToBase10(Bin4Len6) + 1, 1), "=")
    If Break > 0 Then
        I = Break + 1
        While I < Len(Base64Encode)
            Base64Encode = VBA.Left$(Base64Encode, I - 1) & vbCrLf & Mid(Base64Encode, I)
            I = I + Break + 2
        Wend
    End If
End Function
Function Base64Decode(ByVal Base64String As String) As String
    Dim I As Integer, Bin1Len8 As String, Bin2Len8 As String, Bin3Len8 As String
    Dim Bin1Len6 As String, Bin2Len6 As String, Bin3Len6 As String, Bin4Len6 As String
    Base64String = RemoveFromString(Base64String, " ")
    Base64String = RemoveFromString(Base64String, vbCr)
    Base64String = RemoveFromString(Base64String, vbLf)
    If Base64String = vbNullString Then Exit Function
    For I = 0 To 255
        If InStr(Base64String, Chr(I)) > 0 And Not _
            ((InStr(Base64Chars, Chr(I)) > 0) Or (I = Asc("="))) Then Exit Function
    Next
    If Not Len(Base64String) / 4 = Len(Base64String) \ 4 Then Exit Function
    For I = 1 To Len(Base64String) Step 4
        Bin1Len6 = Base10ToBinary(InStr(Base64Chars, Mid(Base64String, I, 1)) - 1)
        Bin2Len6 = Base10ToBinary(InStr(Base64Chars, Mid(Base64String, I + 1, 1)) - 1)
        If Mid(Base64String, I + 2, 1) = "=" Then Bin3Len6 = "0" Else Bin3Len6 = Base10ToBinary(InStr(Base64Chars, Mid(Base64String, I + 2, 1)) - 1)
        If Mid(Base64String, I + 3, 1) = "=" Then Bin4Len6 = "0" Else Bin4Len6 = Base10ToBinary(InStr(Base64Chars, Mid(Base64String, I + 3, 1)) - 1)
        Call Bin4x6To3x8(Bin1Len6, Bin2Len6, Bin3Len6, Bin4Len6, Bin1Len8, Bin2Len8, Bin3Len8)
        Base64Decode = Base64Decode & Chr(BinaryToBase10(Bin1Len8))
        If Not Mid(Base64String, I + 2, 1) = "=" Then Base64Decode = Base64Decode & Chr(BinaryToBase10(Bin2Len8))
        If Not Mid(Base64String, I + 3, 1) = "=" Then Base64Decode = Base64Decode & Chr(BinaryToBase10(Bin3Len8))
    Next
End Function




