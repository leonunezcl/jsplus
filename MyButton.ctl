VERSION 5.00
Begin VB.UserControl MyButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1860
   ClipBehavior    =   0  'None
   FillStyle       =   0  'Solid
   PropertyPages   =   "MyButton.ctx":0000
   ScaleHeight     =   118
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   124
   ToolboxBitmap   =   "MyButton.ctx":0035
End
Attribute VB_Name = "MyButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'developed by edin omeragic
'from: Bosnia and Hercegovina, Srebrenik
'my email: edoo_ba@hotmail.com
'datum: (20.11 - 3.12) 2002 godine
'type: small project
'==================================================================
'this code is totaly free,
'if you dont like this you may copy it on flopy and throw it away ;},
'If you Like it Then please vote on planetsourcecode
'and also search for:
' - iMenu (old project but good) or
' - iList (cool list)
'==================================================================
'DrawButton(State)  - draws button (main function)
'DrawText(...)      - draws text (called from drawbutton)
'DrawPicture(...)   - draws picture
'DrawPictureDisabled - draws picture grayed
'TilePicture()       - tiles picture
'SetRect (left, top, right, bottom) as RECT 'makes rectangle on flay
'ModyfyRect(RECT,left,top,right,bottom) as RECT
'i.e.
'R = SetRect     (0,0,1,1)
'R = ModifyRect(R,1,1,1,1)
'R is            (1,1,2,2)
'==================================================================
'-for default skin, name the picture box "MyButtonDefSkin"
'-for changing skin in design time set property
' "SkinPictureName" same as picture box name
'==================================================================

Option Explicit

'Default Property Values:
Const m_def_TextAlign = vbCenter
Const m_def_PictureTColor = &HFF00FF
Const m_def_PicturePos = 0
Const m_def_TextColorDisabled2 = 0
Const m_def_DrawFocus = 0
Const m_def_DisplaceText = 0
'Const m_def_DownTextDX = 0
'Const m_def_DownTextDY = 0
Const m_def_DisableHover = False
Const m_def_TextColorEnabled = 0
Const m_def_TextColorDisabled = 0
Const m_def_FillWithColor = True
Const m_def_SizeCW = 3
Const m_def_SizeCH = 3
Const m_def_Text = vbNullString
'Property Variables:
Dim m_TextAlign As AlignmentConstants
Dim m_PictureTColor As Ole_Color
Dim m_PicturePos As Integer
Dim m_Picture As StdPicture
Dim m_TextColorDisabled2 As Ole_Color
Dim m_DrawFocus As Integer
Dim m_DisplaceText As Integer
Dim m_DisableHover As Boolean
Dim m_TextColorEnabled As Ole_Color
Dim m_TextColorDisabled As Ole_Color
Dim m_FillWithColor As Boolean
Dim m_SizeCW As Long
Dim m_SizeCH As Long
Dim m_SkinPicture As PictureBox
Dim m_Text As String
Dim m_State As Integer
Dim m_HasFocus As Boolean
Dim m_BtnDown As Boolean
Dim m_SpcDown As Boolean
Dim m_SkinPictureName As String


Public Enum EnumPicturePos
    ppLeft
    ppTop
    ppBottom
    ppRight
    ppCenter
End Enum
Private Const DI_NORMAL As Long = &H3

Const BTN_NORMAL = 1
Const BTN_FOCUS = 2
Const BTN_HOVER = 3
Const BTN_DOWN = 4
Const BTN_DISABLED = 5
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MouseHover()
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event KeyPress(KeyAscii As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Public Enum EnumDrawTextFormat
    DT_BOTTOM = &H8
    DT_CALCRECT = &H400
    DT_CENTER = &H1
    DT_CHARSTREAM = 4
    DT_DISPFILE = 6
    DT_EXPANDTABS = &H40
    DT_EXTERNALLEADING = &H200
    DT_INTERNAL = &H1000
    DT_LEFT = &H0
    DT_METAFILE = 5
    DT_NOCLIP = &H100
    DT_NOPREFIX = &H800
    DT_PLOTTER = 0
    DT_RASCAMERA = 3
    DT_RASDISPLAY = 1
    DT_RASPRINTER = 2
    DT_RIGHT = &H2
    DT_SINGLELINE = &H20
    DT_TABSTOP = &H80
    DT_TOP = &H0
    DT_VCENTER = &H4
    DT_WORDBREAK = &H10
    DT_WORD_ELLIPSIS = &H40000
    DT_END_ELLIPSIS = 32768
    DT_PATH_ELLIPSIS = &H4000
    DT_EDITCONTROL = &H2000
    '===================
    DT_INCENTER = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
End Enum

Private Const SRCCOPY = &HCC0020
Private Const RGN_AND = 1

Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectClipPath Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function apiDrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function apiTranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As Ole_Color, ByVal palet As Long, Col As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
'MY NOTE: TransparentBlt on Win98 leavs some garbage in memory...
'
'Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long

'for picture
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
'Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
'never enough
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long


Private Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(1) As RGBQUAD
End Type

'windows version
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long



'#############################################
'//GDI + SOMETHING ELSE#######################
Private Sub TransBlt(ByVal hdcDest As Long, ByVal XDest As Long, ByVal YDest As Long, _
            ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, _
            ByVal XSrc As Long, ByVal YSrc As Long, ByVal clrMask As Ole_Color)
    
'one check to see if GdiTransparentblt is supported
'better way to check if function is suported is using LoadLibrary and GetProcAdress
'than using GetVersion or GetVersionEx
'=====================================================
    Dim Lib As Long
    Dim ProcAdress As Long
    Dim lMaskColor As Long
    lMaskColor = TranslateColor(clrMask)
    Lib = LoadLibrary("gdi32.dll")
    '-------------------------------->make sure to specify corect name for function
    ProcAdress = GetProcAddress(Lib, "GdiTransparentBlt")
    FreeLibrary Lib
    If ProcAdress <> 0 Then
        'works on XP
        GdiTransparentBlt hdcDest, XDest, YDest, nWidth, nHeight, hdcSrc, XSrc, YSrc, nWidth, nHeight, lMaskColor
        'Debug.Print "Gdi transparent blt"
        Exit Sub 'make it short
    End If
'=====================================================
    Const DSna              As Long = &H220326
    Dim hdcMask             As Long
    Dim hdcColor            As Long
    Dim hbmMask             As Long
    Dim hbmColor            As Long
    Dim hbmColorOld         As Long
    Dim hbmMaskOld          As Long
    Dim hdcScreen           As Long
    Dim hdcScnBuffer        As Long
    Dim hbmScnBuffer        As Long
    Dim hbmScnBufferOld     As Long
    

   hdcScreen = UserControl.hdc
   
   lMaskColor = TranslateColor(clrMask)
   hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
   hdcScnBuffer = CreateCompatibleDC(hdcScreen)
   hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)

   BitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcDest, XDest, YDest, vbSrcCopy

   hbmColor = CreateCompatibleBitmap(hdcScreen, nWidth, nHeight)
   hbmMask = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)

   hdcColor = CreateCompatibleDC(hdcScreen)
   hbmColorOld = SelectObject(hdcColor, hbmColor)
    
   Call SetBkColor(hdcColor, GetBkColor(hdcSrc))
   Call SetTextColor(hdcColor, GetTextColor(hdcSrc))
   Call BitBlt(hdcColor, 0, 0, nWidth, nHeight, hdcSrc, XSrc, YSrc, vbSrcCopy)

   hdcMask = CreateCompatibleDC(hdcScreen)
   hbmMaskOld = SelectObject(hdcMask, hbmMask)

   SetBkColor hdcColor, lMaskColor
   SetTextColor hdcColor, vbWhite
   BitBlt hdcMask, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcCopy
 
   SetTextColor hdcColor, vbBlack
   SetBkColor hdcColor, vbWhite
   BitBlt hdcColor, 0, 0, nWidth, nHeight, hdcMask, 0, 0, DSna
   BitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcMask, 0, 0, vbSrcAnd
   BitBlt hdcScnBuffer, 0, 0, nWidth, nHeight, hdcColor, 0, 0, vbSrcPaint
   BitBlt hdcDest, XDest, YDest, nWidth, nHeight, hdcScnBuffer, 0, 0, vbSrcCopy
     
     'clear
   DeleteObject SelectObject(hdcColor, hbmColorOld)
   DeleteDC hdcColor
   DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
   DeleteDC hdcScnBuffer
   DeleteObject SelectObject(hdcMask, hbmMaskOld)
   
   DeleteDC hdcMask
   'ReleaseDC 0, hdcScreen
End Sub

Private Function GetRgbQuad(ByVal r As Byte, ByVal G As Byte, ByVal b As Byte) As RGBQUAD
    With GetRgbQuad
        .rgbBlue = b
        .rgbGreen = G
        .rgbRed = r
    End With
End Function
Private Function DrawPictureDisabled(ByVal p As StdPicture, x As Long, y As Long, _
                 w As Long, H As Long, _
                 Optional ColHighlight As Long = vb3DHighlight, _
                 Optional ColShadow As Long = vb3DShadow)
                 
    Dim MemDC As Long
    Dim MyBmp As Long
    Dim cShadow As Long
    Dim cHiglight As Long
    Dim ColPal(0 To 1) As RGBQUAD
    Dim rgbBlack As RGBQUAD
    Dim rgbWhite As RGBQUAD
    Dim BI As BITMAPINFO
    Dim hdc As Long
    Dim hPicDc As Long
    Dim hPicBmp As Long
    hdc = UserControl.hdc
    
    cHiglight = TranslateColor(vb3DHighlight)
    cShadow = TranslateColor(vb3DShadow)
    
    'rgbBlack = GetRgbQuad(0, 0, 0)
    rgbWhite = GetRgbQuad(255, 255, 255)
    
    With BI.bmiHeader
        .biSize = 40 'size of bmiHeader structure
        .biHeight = -H
        .biWidth = w
        .biPlanes = 1
        .biCompression = 0 'BI_RGB
        .biClrImportant = 0
        .biBitCount = 1 'monohrome bitmap
    End With
    
    'color palete
    With BI
        .bmiColors(0) = rgbBlack
        .bmiColors(1) = rgbWhite
    End With
    
    Dim hMonoSec As Long
    Dim pBits As Long
    Dim hdcMono As Long
    
    hMonoSec = CreateDIBSection(hdc, BI, 0, pBits, 0&, 0&)
    'Debug.Print "MonoSec:"; hMonoSec
    hdcMono = CreateCompatibleDC(hdc)
    SelectObject hdcMono, hMonoSec
    
    'create dc for picture
    hPicDc = CreateCompatibleDC(hdc)
    If p.Type = vbPicTypeIcon Then
        hPicBmp = CreateCompatibleBitmap(hdc, w, H)
        SelectObject hPicDc, hPicBmp
        DeleteObject hPicBmp
        ClearRect hPicDc, SetRect(0, 0, w, H), TranslateColor(m_PictureTColor)
        DrawIconEx hPicDc, 0, 0, p.Handle, w, H, 0, 0, DI_NORMAL
        'Debug.Print "DRAW ICON"
    ElseIf p.Type = vbPicTypeBitmap Then
        SelectObject hPicDc, p.Handle
    End If
    
    'copy  hPicDc to hdcMono
    BitBlt hdcMono, 0, 0, w, H, hPicDc, 0, 0, SRCCOPY
    
    DeleteDC hPicDc
    
    Dim r As Integer, G As Integer, b   As Integer
    GetRgb cHiglight, r, G, b
    
    'change black color in palete to highlight(r,g,b) color
    ColPal(0) = GetRgbQuad(r, G, b)
    ColPal(1) = rgbBlack    'change white color in palete to black color
    
    SetDIBColorTable hdcMono, 0, 2, ColPal(0)   'set new palete
    RealizePalette hdcMono                      'update it
    'BitBlt Me.hdc, 1, 1, W, H, hdcMono,  0, 0, SRCCOPY
      
    'transparent blit to dest hDC using black as transparent colour
    'x+1 and y+1 - moves down and left for 1 pixel
    TransBlt hdc, x + 1, y + 1, w, H, hdcMono, 0, 0, 0
    
    'get rgb components of shadow color
    GetRgb cShadow, r, G, b
    'change black color to shadow color in palete
    ColPal(0) = GetRgbQuad(r, G, b)
    ColPal(1) = rgbWhite 'change back to white
    
    'set new palete
    SetDIBColorTable hdcMono, 0, 2, ColPal(0)
    RealizePalette hdcMono ' then update
    
    'transparent blit do dest hdc using white color as transparent
    TransBlt hdc, x, y, w, H, hdcMono, 0, 0, RGB(255, 255, 255)
    
    'BitBlt Me.hDC, 0, 0, W, H, hdcMono, 0, 0, SRCCOPY
    
    'Debug.Print DeleteObject(hMonoSec)
    'Debug.Print DeleteObject(hdcMono)
   
End Function
Sub GetRgb(Color As Long, r As Integer, G As Integer, b As Integer)
       r = Color And 255            'clear bites from 9 to 32
       G = (Color \ 256) And 255    'shift right 8 bits and clear
       b = (Color \ 65536) And 255  'shift 16 bits and clear for any case
End Sub

Private Function GetBmpSize(Bmp As StdPicture, w As Long, H As Long) As Long
'    Dim B As BITMAP
'    GetBmpSize = GetObject(Bmp, Len(B), B)
    
    w = ScaleX(Bmp.Width, vbHimetric, vbPixels)
    H = ScaleY(Bmp.Height, vbHimetric, vbPixels)
        
'    Debug.Print W, H
    
    
'    W = B.bmWidth
'    H = B.bmHeight
'    Debug.Print B.bmType
'    Debug.Print W, H
End Function

Private Sub DrawPicture(hdc As Long, p As StdPicture, x As Long, y As Long, w As Long, H As Long, TOleCol As Long)
    
    'check picture format
    If p.Type = vbPicTypeIcon Then
        DrawIconEx hdc, x, y, p.Handle, w, H, 0, 0, DI_NORMAL
        Exit Sub
    End If
    
    'creting dc with the same format as screen dc
    Dim MemDC As Long
    MemDC = CreateCompatibleDC(0)
    
    'select a picture into memdc
    SelectObject MemDC, p.Handle '
    
    'tranparent blit memdc on usercontrol
    TransBlt UserControl.hdc, x, y, w, H, MemDC, 0, 0, TranslateColor(TOleCol)
    
    DeleteDC MemDC 'its clear, heh
End Sub


Private Function ModifyRect(lpRect As RECT, ByVal Left As Long, ByVal Top As Long, _
               ByVal Right As Long, ByVal Bottom As Long) As RECT
    With ModifyRect
        .Left = lpRect.Left + Left
        .Top = lpRect.Top + Top
        .Right = lpRect.Right + Right
        .Bottom = lpRect.Bottom + Bottom
    End With
End Function
Private Function TranslateColor(ByVal Ole_Color As Long) As Long
        apiTranslateColor Ole_Color, 0, TranslateColor
End Function
Private Function SetRect(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As RECT
  With SetRect
    .Left = Left
    .Top = Top
    .Right = Right
    .Bottom = Bottom
  End With
End Function
Private Sub NormalizeRect(r As RECT)
    Dim C As Long
    If r.Left > r.Right Then
        C = r.Right
        r.Right = r.Left
        r.Left = C
    End If
    If r.Top > r.Bottom Then
        C = r.Top
        r.Top = r.Bottom
        r.Bottom = C
    End If
End Sub
Private Function RoundUp(ByVal Num As Single) As Long
    If Int(Num) < Num Then
        RoundUp = Int(Num) + 1
    Else
        RoundUp = Num
    End If
End Function
Private Function RectHeight(r As RECT) As Long
    RectHeight = r.Bottom - r.Top
End Function
Private Function RectWidth(r As RECT) As Long
    RectWidth = r.Right - r.Left
End Function
Private Sub DrawText(ByVal hdc As Long, ByVal strText As String, r As RECT, ByVal Format As Long)
    apiDrawText UserControl.hdc, strText, Len(strText), r, Format
End Sub
Private Sub TilePicture(DestRect As RECT, SrcRect As RECT, ByVal SrcDC As Long, Optional UseCliper As Boolean = True, Optional ROp As Long = SRCCOPY)
    
    Dim i As Integer
    Dim j As Integer
    Dim rows As Integer
    Dim ColS As Integer
    Dim destW As Long
    Dim destH As Long
    Dim hdc As Long
    hdc = UserControl.hdc
    
    NormalizeRect DestRect
    NormalizeRect SrcRect
       
    'calculates row and cols
    rows = RoundUp(RectHeight(DestRect) / RectHeight(SrcRect))
    ColS = RoundUp(RectWidth(DestRect) / RectWidth(SrcRect))
    
    destW = RectWidth(SrcRect)
    destH = RectHeight(SrcRect)
   
    'prevents drawing out of specified rectangle
    If UseCliper Then
        SelectClipRgn hdc, ByVal 0
        BeginPath hdc
            With DestRect
                 Rectangle hdc, .Left, .Top, .Right + 1, .Bottom + 1
            End With
        EndPath hdc
        SelectClipPath hdc, RGN_AND
    End If
    
    For i = 0 To rows - 1
        For j = 0 To ColS - 1
            BitBlt hdc, j * destW + DestRect.Left, i * destH + DestRect.Top, destW, destH, SrcDC, _
            SrcRect.Left, SrcRect.Top, ROp
        Next
    Next
    
    If UseCliper Then
        SelectClipRgn hdc, ByVal 0
    End If
End Sub

Private Sub ClearRect(ByVal hdc As Long, lRect As RECT, ByVal Color As Long)
    Dim Brush As Long
    Dim PBrush As Long
    Brush = CreateSolidBrush(Color)
    PBrush = SelectObject(hdc, Brush)
    
    FillRect hdc, lRect, Brush
    DeleteObject SelectObject(hdc, PBrush)
End Sub
'//END GDI####################################
'#############################################

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,3
Public Property Get SizeCW() As Long
Attribute SizeCW.VB_Description = "Corner width."
Attribute SizeCW.VB_ProcData.VB_Invoke_Property = ";Position"
    SizeCW = m_SizeCW
End Property

Public Property Let SizeCW(ByVal New_SizeCW As Long)
        m_SizeCW = New_SizeCW
        PropertyChanged "SizeCW"
        Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,3
Public Property Get SizeCH() As Long
Attribute SizeCH.VB_Description = "Corner height."
Attribute SizeCH.VB_ProcData.VB_Invoke_Property = ";Position"
    SizeCH = m_SizeCH
End Property

Public Property Let SizeCH(ByVal New_SizeCH As Long)
        m_SizeCH = New_SizeCH
        PropertyChanged "SizeCH"
        Refresh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9,0,0,0
Public Property Get SkinPicture() As Object
Attribute SkinPicture.VB_Description = "Reference to picture box object."
    Set SkinPicture = m_SkinPicture
End Property

Public Property Set SkinPicture(New_SkinPicture As Object)
    
    
    If (TypeName(New_SkinPicture) <> "PictureBox") And _
       (New_SkinPicture Is Nothing = False) Then
        
        Err.Raise 5, "MyButton::SkinPicture", Err.description
        Exit Property
    End If
               
    Set m_SkinPicture = New_SkinPicture
    
    If m_SkinPicture Is Nothing = False Then
        m_SkinPictureName = m_SkinPicture.Name
    Else
        m_SkinPictureName = vbNullString
    End If
    
    Refresh
    PropertyChanged "SPN"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Text() As String
Attribute Text.VB_Description = "Button text."
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Text"
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    Refresh
    PropertyChanged "Text"
    
    'setting access key (allows alt + accesskey)
    
    Dim i As Long
    Dim C As String
    
    For i = 1 To Len(New_Text) - 1
        If Mid(New_Text, i, 1) = "&" Then
            C = Mid(New_Text, i + 1, 1)
            If C <> "&" Or C <> " " Then
                UserControl.AccessKeys = C
                PropertyChanged "AccessKey"
            End If
        End If
        
    Next
   
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get SkinPictureName() As String
Attribute SkinPictureName.VB_Description = "Allows you to set reference at design time."
Attribute SkinPictureName.VB_ProcData.VB_Invoke_Property = ";Appearance"
    'If m_SkinPicture Is Nothing = False Then
        'SkinPictureName = m_SkinPicture.Name
        SkinPictureName = m_SkinPictureName
    'End If
End Property

Public Property Let SkinPictureName(ByVal New_SkinPictureName As String)
    On Error GoTo NotLegalName
    Dim p As Object
    'Debug.Print New_SkinPictureName
    If New_SkinPictureName <> "" Then
        
        Set p = UserControl.Parent.Controls(New_SkinPictureName)
        
        If p Is Nothing = False Then
            Set SkinPicture = p
            'Debug.Print "Setting p"; P.Name
        End If
    Else
        Set m_SkinPicture = Nothing
        'Debug.Print "P is nothing"
        Refresh
    End If
   
'    m_SkinPictureName = New_SkinPictureName
    PropertyChanged "SPN"
NotLegalName:
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    DrawButton BTN_DOWN
End Sub

Private Sub UserControl_GotFocus()
    m_HasFocus = True
    If m_BtnDown = False Then DrawButton BTN_FOCUS
End Sub

Private Sub UserControl_Initialize()
'    SkinPictureName = m_SkinPictureName
'    MsgBox "Initialize..." + m_SkinPictureName
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_SizeCW = m_def_SizeCW
    m_SizeCH = m_def_SizeCH
    m_Text = Extender.Name
    m_FillWithColor = m_def_FillWithColor
    m_TextColorEnabled = m_def_TextColorEnabled
    m_TextColorDisabled = m_def_TextColorDisabled
    Set UserControl.Font = Ambient.Font
    m_DisableHover = m_def_DisableHover

    m_DisplaceText = m_def_DisplaceText
    m_DrawFocus = m_def_DrawFocus
    m_TextColorDisabled2 = m_def_TextColorDisabled2
    Set m_Picture = LoadPicture("")
    m_PicturePos = m_def_PicturePos
    m_PictureTColor = m_def_PictureTColor
    m_SkinPictureName = "MyButtonDefSkin"
    m_TextAlign = m_def_TextAlign
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    
    If KeyCode = vbKeySpace Then
        m_SpcDown = True
        DrawButton BTN_DOWN
    Else
        m_SpcDown = False
        DrawButton BTN_FOCUS
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    If KeyCode = 32 And m_SpcDown And m_State = BTN_DOWN Then
        m_SpcDown = False
        
        DrawButton BTN_NORMAL
        RaiseEvent Click
        DrawButton BTN_FOCUS
        
    End If
End Sub

Private Sub UserControl_LostFocus()
    m_HasFocus = False
    DrawButton BTN_NORMAL
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseDown(Button, Shift, x, y)
   If Button = 1 Then m_BtnDown = True
   UserControl_MouseMove Button, Shift, x, y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_SpcDown Then Exit Sub
    
    RaiseEvent MouseMove(Button, Shift, x, y)
    SetCapture hwnd
    If PointInControl(x, y) Then
        'if pointer is on control
        If m_BtnDown Then
            If m_State <> BTN_DOWN Then
                DrawButton BTN_DOWN
            End If
        Else
            If m_State <> BTN_HOVER Then
                RaiseEvent MouseHover
                DrawButton BTN_HOVER
            End If
            
        End If
    Else
        'if pointer is out of control
        If m_BtnDown Then
            
            RaiseEvent MouseHover
            DrawButton BTN_HOVER
            
        Else
            RaiseEvent MouseOut
            If m_HasFocus Then
                DrawButton BTN_FOCUS
            Else
                DrawButton BTN_NORMAL
            End If
            ReleaseCapture
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    m_BtnDown = False
'    If m_State <> BTN_NORMAL Then
        DrawButton BTN_NORMAL
'    End If
    
    RaiseEvent MouseUp(Button, Shift, x, y)
    
    If Button = vbLeftButton Then
        'aqui cambie yo el codigo
        'If PointInControl(x, Y) Then RaiseEvent Click
        'lnunez
'        If m_State <> BTN_FOCUS Then
            DrawButton BTN_FOCUS
'        End If
    End If
    
End Sub


Private Sub UserControl_Paint()
    Me.Refresh
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_SizeCW = PropBag.ReadProperty("SizeCW", m_def_SizeCW)
    m_SizeCH = PropBag.ReadProperty("SizeCH", m_def_SizeCH)
    m_SkinPictureName = PropBag.ReadProperty("SPN", "")
   
    'Debug.Print "ReadProp SPN:"; m_SkinPictureName
   
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_FillWithColor = PropBag.ReadProperty("FillWithColor", m_def_FillWithColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.AccessKeys = PropBag.ReadProperty("AccessKey", "")
    m_TextColorEnabled = PropBag.ReadProperty("TextColorEnabled", m_def_TextColorEnabled)
    m_TextColorDisabled = PropBag.ReadProperty("TextColorDisabled", m_def_TextColorDisabled)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    m_DisableHover = PropBag.ReadProperty("DisableHover", m_def_DisableHover)
'    m_DownTextDX = PropBag.ReadProperty("DownTextDX", m_def_DownTextDX)
'    m_DownTextDY = PropBag.ReadProperty("DownTextDY", m_def_DownTextDY)
    m_DisplaceText = PropBag.ReadProperty("DisplaceText", m_def_DisplaceText)
    m_DrawFocus = PropBag.ReadProperty("DrawFocus", m_def_DrawFocus)
    m_TextColorDisabled2 = PropBag.ReadProperty("TextColorDisabled2", m_def_TextColorDisabled2)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    m_PicturePos = PropBag.ReadProperty("PicturePos", m_def_PicturePos)
    m_PictureTColor = PropBag.ReadProperty("PictureTColor", m_def_PictureTColor)
    m_TextAlign = PropBag.ReadProperty("TextAlign", m_def_TextAlign)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Private Sub UserControl_Show()
    
    SkinPictureName = m_SkinPictureName

   ' Refresh
End Sub

Private Sub UserControl_Terminate()
    Set m_SkinPicture = Nothing
    Set m_Picture = Nothing
    
    'Set UserControl = Nothing
    'Set Me = Nothing
    'Debug.Print "TERMINATE"
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SizeCW", m_SizeCW, m_def_SizeCW)
    Call PropBag.WriteProperty("SizeCH", m_SizeCH, m_def_SizeCH)
    
    'If m_SkinPicture Is Nothing = False Then
        Call PropBag.WriteProperty("SPN", m_SkinPictureName, "")
    'End If
    
    'Debug.Print "Write :"; m_SkinPictureName
    
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("FillWithColor", m_FillWithColor, m_def_FillWithColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("AccessKey", UserControl.AccessKeys, "")
    Call PropBag.WriteProperty("TextColorEnabled", m_TextColorEnabled, m_def_TextColorEnabled)
    Call PropBag.WriteProperty("TextColorDisabled", m_TextColorDisabled, m_def_TextColorDisabled)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)

    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("DisableHover", m_DisableHover, m_def_DisableHover)
    Call PropBag.WriteProperty("DisplaceText", m_DisplaceText, m_def_DisplaceText)
    Call PropBag.WriteProperty("DrawFocus", m_DrawFocus, m_def_DrawFocus)
    Call PropBag.WriteProperty("TextColorDisabled2", m_TextColorDisabled2, m_def_TextColorDisabled2)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("PicturePos", m_PicturePos, m_def_PicturePos)
    Call PropBag.WriteProperty("PictureTColor", m_PictureTColor, m_def_PictureTColor)
    Call PropBag.WriteProperty("TextAlign", m_TextAlign, m_def_TextAlign)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
End Sub


Private Sub DrawButton(ByVal State As Integer)
    
    If m_DisableHover Then
        If State = BTN_HOVER Then Exit Sub
        'dont draw hover state if m_DisableHover is true
    End If
'    Debug.Print "State1 "; State

    On Error GoTo UnknownError

    Dim PicW As Long
    Dim PicH As Long 'width and height of picture

    Dim PicX As Long
    Dim PicY As Long 'picture pos

    Dim DH As Long  'button height
    Dim dw As Long  'button width
    Dim Align As Long 'text aligment
    Dim bDrawText As Boolean ' if picture is in center text is not drawn
    bDrawText = True

    Align = DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS

    Select Case m_TextAlign
        Case Is = vbLeftJustify:  Align = Align Or DT_LEFT
        Case Is = vbRightJustify: Align = Align Or DT_RIGHT
        Case Is = vbCenter:       Align = Align Or DT_CENTER
    End Select

    dw = UserControl.ScaleWidth
    DH = UserControl.ScaleHeight

    m_State = State
    'if skin picture is not set then just draw text on control
    If m_SkinPicture Is Nothing Then
        ClearRect hdc, SetRect(0, 0, dw, DH), TranslateColor(UserControl.BackColor)
        DrawText hdc, m_Text, SetRect(0, 0, dw, DH), Align
        If UserControl.AutoRedraw = True Then
            UserControl.Refresh
        End If
        Exit Sub
    End If


    m_SkinPicture.ScaleMode = vbPixels


    Dim SrcLeft As Long     'left cordinate of skin in skinpicture
    Dim SrcRight As Long    'right -II-
    Dim FillColor As Long   'color to fill middle area of button
                            'used if m_FillWithColor is true

    Dim H As Long           'height of skinpicture
    Dim w As Long           'width of button skin

    H = m_SkinPicture.ScaleHeight
    w = m_SkinPicture.ScaleWidth / 5
'Debug.Print H, W
'
    SrcLeft = (State - 1) * w
    SrcRight = State * w

    If m_FillWithColor Then
        'get color to fill with from (SrcLeft+m_SizeCW +1 , m_SizeCH+1) on
        'skin picture
        FillColor = m_SkinPicture.Point(SrcLeft + m_SizeCW + 1, m_SizeCH + 1)
    End If

'Exit Sub
    ClearRect hdc, SetRect(0, 0, dw, DH), TranslateColor(UserControl.BackColor)
    If m_FillWithColor Then
        'paint button with fillcolor
        'NOTE: it would be nice if there is gradient file
        ClearRect hdc, SetRect(m_SizeCW, m_SizeCH, dw - m_SizeCW, DH - m_SizeCH), FillColor
        'ABOUT ADDING GRADIENT FILL
        'read second color from skin at
        'point (srcleft+cw+1, H -m_sizeCH-1)
        'may be implemented in MyButton2
    Else
        'tile skin
         TilePicture SetRect(m_SizeCW, m_SizeCH, dw - m_SizeCW, DH - m_SizeCH), _
           SetRect(SrcLeft + m_SizeCW, m_SizeCH, SrcRight - m_SizeCW, H - m_SizeCH), _
           m_SkinPicture.hdc, False, SRCCOPY
    End If

    'draws borders
    If (m_SizeCH > 0 And m_SizeCW > 0) Then
        TilePicture SetRect(m_SizeCW, 0, dw, m_SizeCH), _
            SetRect(SrcLeft + m_SizeCW, 0, SrcRight - m_SizeCW, m_SizeCH), _
            m_SkinPicture.hdc, False, SRCCOPY

        TilePicture SetRect(m_SizeCW, DH - m_SizeCH, dw, DH), _
            SetRect(SrcLeft + m_SizeCW, H - m_SizeCH, SrcRight - m_SizeCW, H), _
            m_SkinPicture.hdc, False, SRCCOPY

        TilePicture SetRect(0, 0, m_SizeCW, DH), _
            SetRect(SrcLeft, m_SizeCH, SrcLeft + m_SizeCW, H - m_SizeCH), _
            m_SkinPicture.hdc, False, SRCCOPY

        TilePicture SetRect(dw - m_SizeCW, m_SizeCH, dw, DH - m_SizeCH), _
            SetRect(SrcRight - m_SizeCW, m_SizeCH, SrcRight, H - m_SizeCH), _
            m_SkinPicture.hdc, False, SRCCOPY

        'draws corners
        'NOTE: must chage to transparent blit (done)
        TransBlt hdc, 0, 0, m_SizeCW, m_SizeCH, m_SkinPicture.hdc, SrcLeft, 0, &HFF00FF
        TransBlt hdc, 0, DH - m_SizeCH, m_SizeCW, m_SizeCH, m_SkinPicture.hdc, SrcLeft, H - m_SizeCH, &HFF00FF

        TransBlt hdc, dw - m_SizeCW, 0, m_SizeCW, m_SizeCH, m_SkinPicture.hdc, SrcRight - m_SizeCW, 0, &HFF00FF
        TransBlt hdc, dw - m_SizeCW, DH - m_SizeCH, m_SizeCW, m_SizeCH, m_SkinPicture.hdc, SrcRight - m_SizeCW, H - m_SizeCH, &HFF00FF
    End If

    Dim PColor As Long 'previous color

    PColor = UserControl.ForeColor

    Dim TextRect As RECT
    If State = BTN_DOWN Then
        TextRect = SetRect(m_SizeCW + m_DisplaceText, m_SizeCH + m_DisplaceText, dw - m_SizeCW + m_DisplaceText - 3, DH - m_SizeCH + m_DisplaceText)
    Else
        TextRect = SetRect(m_SizeCW, m_SizeCH, dw - m_SizeCW - 3, DH - m_SizeCH)
    End If
        If m_Picture Is Nothing Then
            If m_State = BTN_DISABLED Then
                'draw text only
                'dont draw text2 if colors are the same
                If m_TextColorDisabled <> m_TextColorDisabled2 Then
                    UserControl.ForeColor = m_TextColorDisabled2
                    TextRect = ModifyRect(TextRect, 1, 1, 1, 1)
                    DrawText hdc, m_Text, TextRect, Align
                    TextRect = ModifyRect(TextRect, -1, -1, -1, -1)
                End If
                UserControl.ForeColor = m_TextColorDisabled
                DrawText hdc, m_Text, TextRect, Align
            Else
                'draw text only
                UserControl.ForeColor = m_TextColorEnabled
                DrawText hdc, m_Text, TextRect, Align
            End If
        Else

            GetBmpSize m_Picture, PicW, PicH
            PicY = (DH - PicH) / 2
            If m_State = BTN_DOWN Then
                PicY = PicY + m_DisplaceText
            End If



            Select Case m_PicturePos
                Case Is = ppLeft
                    PicX = TextRect.Left + 3
                    TextRect.Left = PicX + PicW + TextRect.Left
                Case Is = ppRight
                    PicX = TextRect.Right - PicW - 3 + TextRect.Left - m_SizeCW
                    TextRect.Right = PicX - 3
                Case Is = ppTop
                    PicX = (dw - PicW) / 2 + TextRect.Left - SizeCW
                    PicY = (DH - PicH - 3 - UserControl.TextHeight("I")) / 2 + TextRect.Top - SizeCH
                    TextRect.Top = PicY + PicW + 3
                    TextRect.Bottom = TextRect.Top + UserControl.TextHeight("I") * 1.2
                Case Is = ppBottom
                    TextRect.Top = (DH - PicH - 3 - UserControl.TextHeight("I")) / 2 + TextRect.Top - SizeCH
                    PicX = (dw - PicW) / 2 + TextRect.Left - SizeCW
                    TextRect.Bottom = TextRect.Top + UserControl.TextHeight("I") * 1.2
                    PicY = TextRect.Bottom + 3
                Case Is = ppCenter
                    PicX = (dw - PicW) / 2
                    If BTN_DOWN Then PicX = PicX + m_DisplaceText
                    bDrawText = False
            End Select

'            Debug.Print "State2 "; State

            If m_State = BTN_DISABLED Then
                'draw text and picture disabled
                DrawPictureDisabled m_Picture, PicX, PicY, PicW, PicH
                If m_TextColorDisabled <> m_TextColorDisabled2 Then
                    If bDrawText Then
                        UserControl.ForeColor = m_TextColorDisabled2
                        TextRect = ModifyRect(TextRect, 1, 1, 1, 1)

                        DrawText hdc, m_Text, TextRect, Align
                        TextRect = ModifyRect(TextRect, -1, -1, -1, -1)
                    End If
                End If

                UserControl.ForeColor = m_TextColorDisabled
                If bDrawText Then
                    DrawText hdc, m_Text, TextRect, Align
                End If
            Else
                'draw text and picture enabled
                UserControl.ForeColor = m_TextColorEnabled
                If bDrawText Then
                    DrawText hdc, m_Text, TextRect, Align
                End If
                DrawPicture hdc, m_Picture, PicX, PicY, PicW, PicH, m_PictureTColor
            End If
        End If

    Dim F As Long
    If m_DrawFocus > 0 Then
        If State = BTN_DOWN Or State = BTN_FOCUS Then
            F = CLng(m_DrawFocus)
            DrawFocusRect hdc, SetRect(F, F, dw - F, DH - F)
        End If
    End If

    UserControl.ForeColor = PColor
    If UserControl.AutoRedraw = True Then
        UserControl.Refresh
    End If
Exit Sub
UnknownError:

'most important line in this function
'i about 2 hours to find out
Set m_SkinPicture = Nothing
'removing this line form will not unload properly
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get FillWithColor() As Boolean
Attribute FillWithColor.VB_Description = "Middle area of button is filled with color if true or tiled with skin."
Attribute FillWithColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FillWithColor = m_FillWithColor
End Property

Public Property Let FillWithColor(ByVal New_FillWithColor As Boolean)
    m_FillWithColor = New_FillWithColor
    Refresh
    PropertyChanged "FillWithColor"
End Property


Public Sub Refresh()

    If m_State < 1 Or m_State > 5 Then m_State = 1
    If Enabled Then
        DrawButton m_State
    Else
        DrawButton BTN_DISABLED
    End If
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property
 
Private Function PointInControl(x As Single, y As Single) As Boolean
  If x >= 0 And x <= UserControl.ScaleWidth And _
    y >= 0 And y <= UserControl.ScaleHeight Then
    PointInControl = True
  End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled

    If New_Enabled Then
        DrawButton BTN_NORMAL
    Else
        DrawButton BTN_DISABLED
    End If
    
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TextColorEnabled() As Ole_Color
Attribute TextColorEnabled.VB_Description = "Color of text when its enabled."
Attribute TextColorEnabled.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TextColorEnabled = m_TextColorEnabled
End Property

Public Property Let TextColorEnabled(ByVal New_TextColorEnabled As Ole_Color)
    m_TextColorEnabled = New_TextColorEnabled
    Refresh
    PropertyChanged "TextColorEnabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TextColorDisabled() As Ole_Color
Attribute TextColorDisabled.VB_Description = "Color of text when button is disabled"
Attribute TextColorDisabled.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TextColorDisabled = m_TextColorDisabled
End Property

Public Property Let TextColorDisabled(ByVal New_TextColorDisabled As Ole_Color)
    m_TextColorDisabled = New_TextColorDisabled
    Refresh
    PropertyChanged "TextColorDisabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Refresh
    PropertyChanged "Font"
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
Attribute FontUnderline.VB_ProcData.VB_Invoke_Property = ";Font"
    FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    UserControl.FontUnderline() = New_FontUnderline
    Refresh
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
Attribute FontStrikethru.VB_ProcData.VB_Invoke_Property = ";Font"
    FontStrikethru = UserControl.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    UserControl.FontStrikethru() = New_FontStrikethru
    Refresh
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
Attribute FontSize.VB_ProcData.VB_Invoke_Property = ";Font"
    FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    Refresh
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontName
Public Property Get fontname() As String
Attribute fontname.VB_Description = "Specifies the name of the font that appears in each row for the given level."
Attribute fontname.VB_ProcData.VB_Invoke_Property = ";Font"
    fontname = UserControl.fontname
End Property

Public Property Let fontname(ByVal New_FontName As String)
    UserControl.fontname() = New_FontName
    Refresh
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
Attribute FontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
    FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    UserControl.FontItalic() = New_FontItalic
    Refresh
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
Attribute FontBold.VB_ProcData.VB_Invoke_Property = ";Font"
    FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    UserControl.FontBold() = New_FontBold
    Refresh
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get DisableHover() As Boolean
Attribute DisableHover.VB_ProcData.VB_Invoke_Property = ";Behavior"
    DisableHover = m_DisableHover
End Property

Public Property Let DisableHover(ByVal New_DisableHover As Boolean)
    m_DisableHover = New_DisableHover
    PropertyChanged "DisableHover"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get DisplaceText() As Integer
Attribute DisplaceText.VB_Description = "Displaces text when button is down."
Attribute DisplaceText.VB_ProcData.VB_Invoke_Property = ";Behavior"
    DisplaceText = m_DisplaceText
End Property

Public Property Let DisplaceText(ByVal New_DisplaceText As Integer)
    m_DisplaceText = New_DisplaceText
    PropertyChanged "DisplaceText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get DrawFocus() As Integer
Attribute DrawFocus.VB_Description = "Draws focus."
Attribute DrawFocus.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DrawFocus = m_DrawFocus
End Property

Public Property Let DrawFocus(ByVal New_DrawFocus As Integer)
    m_DrawFocus = New_DrawFocus
    PropertyChanged "DrawFocus"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TextColorDisabled2() As Ole_Color
Attribute TextColorDisabled2.VB_Description = "Color of text when button is disabled that make it looks grayed."
Attribute TextColorDisabled2.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TextColorDisabled2 = m_TextColorDisabled2
End Property

Public Property Let TextColorDisabled2(ByVal New_TextColorDisabled2 As Ole_Color)
    m_TextColorDisabled2 = New_TextColorDisabled2
    Refresh
    PropertyChanged "TextColorDisabled2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As StdPicture
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set m_Picture = New_Picture
    Refresh
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get PicturePos() As EnumPicturePos
Attribute PicturePos.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PicturePos = m_PicturePos
End Property

Public Property Let PicturePos(ByVal New_PicturePos As EnumPicturePos)
    m_PicturePos = New_PicturePos
    Refresh
    PropertyChanged "PicturePos"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get PictureTColor() As Ole_Color
Attribute PictureTColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PictureTColor = m_PictureTColor
End Property

Public Property Let PictureTColor(ByVal New_PictureTColor As Ole_Color)
    m_PictureTColor = New_PictureTColor
    Refresh
    PropertyChanged "PictureTColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get TextAlign() As AlignmentConstants
Attribute TextAlign.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TextAlign = m_TextAlign
End Property

Public Property Let TextAlign(ByVal New_TextAlign As AlignmentConstants)
    m_TextAlign = New_TextAlign
    Refresh
    PropertyChanged "TextAlign"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As Ole_Color
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Ole_Color)
    UserControl.BackColor() = New_BackColor
    Refresh
    PropertyChanged "BackColor"
End Property

