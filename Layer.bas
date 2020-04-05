Attribute VB_Name = "Module8"
Option Explicit

'**************************************
' Name: Windows 2000 Layered Windows, ak
'     a Translucent Windows or Alpha Blended W
'     indows (fully movable etc)
' Description:This code utilizes Windows
'     2000's layered window effect, (commonly
'     referred to as alpha blended windows or
'     translucent windows) described at http:/
'     /msdn.microsoft.com/library/techart/laye
'     rwin.htm. You must have Windows 2000 for
'     this code to work.
' By: Rom
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=8901&lngWId=1'for details.'**************************************

'Win2k layered windows module
'
'This information was found at
'http://msdn.microsoft.com/library/techa
'     rt/layerwin.htm
'and other parts of msdn.
'
'If you want to check if a window is alr
'     eady layered,
'CheckLayered(hwnd) will return true or
'     false
'
'To make a window layered, just use SetL
'     ayered,
'where hwnd is the handle of window, and
'     bAlpha
'is the amount of transparency (e.g. 0 =
'     invisible,
'255 = opaque), and if True is passed to
'     SetAs
'it will make the window layered, if Fal
'     se is
'passed then it will get rid of the laye
'     red property.

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type SIZE
    cx As Long
    cy As Long
End Type

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Const WS_EX_LAYERED = &H80000
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const AC_SRC_OVER = &H0
Private Const AC_SRC_ALPHA = &H1
Private Const AC_SRC_NO_PREMULT_ALPHA = &H1
Private Const AC_SRC_NO_ALPHA = &H2
Private Const AC_DST_NO_PREMULT_ALPHA = &H10
Private Const AC_DST_NO_ALPHA = &H20
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private lret As Long

Function CheckLayered(ByVal hWnd As Long) As Boolean

    lret = GetWindowLong(hWnd, GWL_EXSTYLE)


    If (lret And WS_EX_LAYERED) = WS_EX_LAYERED Then
        CheckLayered = True
    Else
        CheckLayered = False
    End If

End Function


Sub SetLayered(ByVal hWnd As Long, SetAs As Boolean)

    Dim cOS As New clsOS
    Dim bAlpha As Byte
    Dim ret As String
    
    If InStr(windows, "95") Then
        Exit Sub
    ElseIf InStr(windows, "98") Then
        Exit Sub
    ElseIf InStr(windows, "Me") Then
        Exit Sub
    End If
    
    lret = GetWindowLong(hWnd, GWL_EXSTYLE)

    If SetAs = True Then
        lret = lret Or WS_EX_LAYERED
    Else
        lret = lret And Not WS_EX_LAYERED
    End If

    ret = util.LeeIni(IniPath, "misc", "alpha")
    
    If Len(ret) > 0 Then
        bAlpha = CByte(ret)
    Else
        bAlpha = 150
    End If
    
    SetWindowLong hWnd, GWL_EXSTYLE, lret
    SetLayeredWindowAttributes hWnd, 0, bAlpha, LWA_ALPHA
    
    Set cOS = Nothing
    
End Sub

