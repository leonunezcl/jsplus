VERSION 5.00
Begin VB.UserControl jcForms 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3390
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MouseIcon       =   "jcForms.ctx":0000
   ScaleHeight     =   31
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   226
   ToolboxBitmap   =   "jcForms.ctx":030A
   Begin VB.Menu SysMnu 
      Caption         =   "Menu"
      Begin VB.Menu MnuSyst 
         Caption         =   "Restore       "
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu MnuSyst 
         Caption         =   "Minimize"
         Index           =   1
      End
      Begin VB.Menu MnuSyst 
         Caption         =   "Maximize"
         Index           =   2
      End
      Begin VB.Menu MnuSyst 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuSyst 
         Caption         =   "Close            Alt+F4"
         Index           =   4
      End
      Begin VB.Menu MnuSyst 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu MnuSyst 
         Caption         =   "Always on Top"
         Index           =   6
      End
   End
End
Attribute VB_Name = "jcForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'&                                                                           &
'&   jcForms v 1.0.5                                                         &
'&                                                                           &
'&   Copyright © 2006 Juan Carlos San Román Arias (sanroman2004@yahoo.com)   &
'&                                                                           &
'&   You may use this control in your applications free of charge,           &
'&   provided that you do not redistribute this source code without          &
'&   giving me credit for my work. Of course, credit in your                 &
'&   applications is always welcome.                                         &
'&                                                                           &
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'
'   ===================
'   Credits and Thanks:
'   ===================
'
'   I want specially thanks all PSC friends for their advices and comments,
'   without them It will be impossible to do anything.
'
'   Paul Caton          - For the self-subclassing usercontrol code (Version II).
'   Paul R. Territo     - For his helps and advices to solve UC win98 crashing.
'   Fred.cpp            - For his gradient subs
'   Morgan Haueisen     - For his help in code formatting, error checking and
'                         general suggestions
'   Richard Mewett      - For the Unicode support.
'=================================================================================
'   -----------------------------------
'    Version 1.0.0 Data: 13-March-2006
'   -----------------------------------
'   - It is a simple control that simulate an attractive form
'   - You can resize, minimize and maximize form
'   - It automatically changes backcolor of each contained controls
'
'   ------------
'    How to use
'   ------------
'   - Place usercontrol in the form and it automatically changes form
'     - Borderstyle property to none
'     - ShowInTaskBar property to True
'   - Not applicable for MDI forms or forms with MdiChild property = true
'
'   -----------------------------------
'    Version 1.0.1 Data: 15-March-2006
'   -----------------------------------
'   - Added borderstyle property (fixed and sizable) -> form.borderstyle must be =0.
'   - Fixed resize, now Resize is smoothly as normal.
'   - Minimize and maximize buttons interact with the form.MaxButton.
'     and from.MinButton properties.
'
'   -----------------------------------
'    Version 1.0.2 Data: 16-March-2006
'   -----------------------------------
'   - Added a Paul Caton's modification of his self subclasser to avoid crashing in w98.
'     when unload form.
'   - Added another type of borderstyle: it is a fixed border with minimize button.
'   - Added CustomTheme Theme, ColorFrom and ColorTo properties.
'
'   -----------------------------------
'    Version 1.0.3 Data: 23-March-2006
'   -----------------------------------
'   - Added iconsize property.
'   - Titlebar autosizing considering icon and font size.
'   - Border resize was improved outstandingly.
'   - Now when you place a control in the form it changes automatically its borderstyle
'     and showinTaskBar properties.
'   - Form height autosizing considering icon size and usercontrol style selection.
'
'   -----------------------------------
'    Version 1.0.4 Data: 27-March-2006
'   -----------------------------------
'
'   I want specially thank Morgan Haueisen for his help and suggestions made in this version
'   of jcFomrs. Morgan proposed me to add next properties to UC:
'   - Hide Close Button property
'   - ChangeAllBackgrounds property
'   - Moveable property link to parent form
'   - Error checking
'
'   And to fix next items
'   - to swap customcolors property name for customtheme
'   - caption placement when Close button is not showing
'   - Form drag when Close button is not visible
'   - General clean-up of code and variables
'   - Memory leaks
'
'   Apart from adding and fixing these items I did the next:
'   - Added MinButton and MaxButton properties to usercontrol, we can change their values
'     when usercontrol borderstyle is "Sizable" regardless of form MaxButton and
'     MinButton and borderstyle values
'   - deleted FixedWithMinBtn option in borderstyle property. Now you can choose but fixed
'     or sizable border and to select if you want MinButton, MaxButton and/or Close Button).
'     Now is more flexible than before.
'   - Added in the jcForm Menu "Always on Top" Menu item.
'   - Added 4 public subs to add, remove and modify Menu Items in the form menu by users:
'     (FormMenuAdd, FormMenuRemove, ModifyAddedMenu and GetMenuItemValue).
'   - Added new event: MenuItemSelected Event to manage added menu items
'   - Fixed UserControl_KeyDown in order to prevent form closing using alt+f4 keys when
'     Close button is not visible.
'
'   -----------------------------------
'    Version 1.0.5 Data: 18-April-2006
'   -----------------------------------
'   - Added 2 new styles (style5 and style6)
'   - Added WindowState property (vbNormal, vbMaximized, vbMinimized).
'   - Added TitleBarShadow property (true, false)
'   - Now it is possible to maximized jcForms at start up
'     (set jcForms.WindowState= vbMaximized).
'   - You can now press ALT + F4 keys and it works even if uc hasn't a focus
'     (close button must be visible).
'   - It changes colors when it is activated or when it loses a focus.
'   - Button drawing (minimize, maximize and close button) when cursor is on it has been fixed
'   - Added new sub (SetjcFormsMenuCaption) to change captions of jcForms Menu
'     according to your language (restore, minimize, maximize, close and Always on top).
'   - Paul Caton's self-subclasser (Version 2.1) has been added.
'   - Added Unicode support.
'   - Added system color change detection (it also detects Windows xp theme changes).
'   - Added SetAlwaysOnTop sub to put form always on top
'
'=======================================================================================

Option Explicit

'*************************************************************
'   Required Enum Definitions
'*************************************************************
'gradient types
Public Enum jcGradConst
    VerticalGradient = 0
    horizontalGradient = 1
    VCylinderGradient = 2
    HCylinderGradient = 3
End Enum

'caption types
Public Enum jcCaptionConst
    Style1 = 0
    Style2 = 1
    Style3 = 2
    Style4 = 3
    Style5 = 4
    Style6 = 5
End Enum

'theme colors
Public Enum jcThemeConst
    blue = 0
    silver = 1
    Olive = 2
    Visual2005 = 3
    Norton2004 = 4
    CustomTheme = 5
    Autodetect = 6
End Enum

'backcolor styles
Public Enum jcBackColor
    Default = 0
    Auto = 1
    Custom = 2
End Enum

'caption button states
Public Enum jcBtnState
    STA_NORMAL = 0
    STA_OVER = 1
    STA_PRESSED = 2
End Enum

'border styles
Public Enum jcBorderStyle
    Fixed = 0
    Sizable = 1
End Enum

'jcFroms Menu item properties
Public Enum jcMenuItemProp
    jcCaption = 0
    jcEnabled = 1
    jcChecked = 2
    jcVisible = 3
End Enum

'jcForms Menu items
Public Enum jcSystMenuItem
    jcRestore = 0
    jcMinimize = 1
    jcMaximize = 2
    jcClose = 4
    jcAlwaysOnTop = 6
End Enum

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

'for subclassing
Private Enum eMsgWhen                                     'When to callback
    MSG_BEFORE = 1                                        'Callback before the original WndProc
    MSG_AFTER = 2                                         'Callback after the original WndProc
    MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER            'Callback before and after the original WndProc
End Enum

'*************************************************************
'   Required Type Definitions
'*************************************************************

'Caption button properties
Private Type CaptionBtn
    Left                            As Long
    Top                             As Long
    Width                           As Long
    Height                          As Long
    TooltipText                     As String
    Visible                         As Boolean
End Type

Private Type POINT
    X                               As Long
    Y                               As Long
End Type

Private Type RECT
    Left                            As Long
    Top                             As Long
    Right                           As Long
    Bottom                          As Long
End Type

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                          As Long
    dwFlags                         As TRACKMOUSEEVENT_FLAGS
    hwndTrack                       As Long
    dwHoverTime                     As Long
End Type

'for unicode support
Private Type OSVERSIONINFO
    dwOSVersionInfoSize             As Long
    dwMajorVersion                  As Long
    dwMinorVersion                  As Long
    dwBuildNumber                   As Long
    dwPlatformId                    As Long
    szCSDVersion                    As String * 128        '  Maintenance string for PSS usage
End Type

'*************************************************************
'  api declares for Subclasser
'*************************************************************
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

'*************************************************************
'  api declares for Unicode support.
'*************************************************************
Private Declare Function DrawTextA Lib "user32" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'*************************************************************
'  api declares for xp theme detection
'*************************************************************
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, ByVal dwMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long

'*************************************************************
'  api declares for drawing
'*************************************************************
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

'*************************************************************
'   api declares for general uses
'*************************************************************
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'*************************************************************
'   Defined Constants
'*************************************************************
'for text alignment
Private Const DT_VCENTER            As Long = &H4
Private Const DT_SINGLELINE         As Long = &H20         ' strip cr/lf from string before draw.
Private Const DT_LEFT               As Long = &H0          ' draw from left edge of rectangle.

'for getting the desktop work area
Private Const SPI_GETWORKAREA       As Long = 48

'for windows language detection
Private Const LOCALE_USER_DEFAULT   As Long = &H400
Private Const LOCALE_SENGLANGUAGE   As Long = &H1001       ' English name of language

'for always on top use
Private Const HWND_TOPMOST          As Long = -1
Private Const HWND_NOTOPMOST        As Long = -2
Private Const SWP_NOMOVE            As Long = &H2
Private Const SWP_NOSIZE            As Long = &H1
Private Const SWP_SHOWWINDOW        As Long = &H40

'for ReleaseCapture
Private Const WM_NCLBUTTONDOWN      As Long = &HA1
Private Const HTCAPTION             As Long = 2

'for subclassing
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_MOUSELEAVE         As Long = &H2A3
Private Const WM_SETCURSOR          As Long = &H20
Private Const WM_MOVE               As Long = &H3
Private Const WM_SYSCOMMAND         As Long = &H112
Private Const WM_ACTIVATE           As Long = &H6
Private Const WM_SYSCOLORCHANGE     As Long = &H15&

Private Const ALL_MESSAGES          As Long = -1           'All messages callback
Private Const MSG_ENTRIES           As Long = 32           'Number of msg table entries
Private Const WNDPROC_OFF           As Long = &H38         'Thunk offset to the WndProc execution address
Private Const GWL_WNDPROC           As Long = -4           'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN          As Long = 1            'Thunk data index of the shutdown flag
Private Const IDX_HWND              As Long = 2            'Thunk data index of the subclassed hWnd
Private Const IDX_WNDPROC           As Long = 9            'Thunk data index of the original WndProc
Private Const IDX_BTABLE            As Long = 11           'Thunk data index of the Before table
Private Const IDX_ATABLE            As Long = 12           'Thunk data index of the After table
Private Const IDX_PARM_USER         As Long = 13           'Thunk data index of the User-defined callback parameter data index

'for unicode support
Private Const VER_PLATFORM_WIN32_NT = 2

'*************************************************************
'   Defined variables
'*************************************************************

Private bTrack                      As Boolean
Private bTrackUser32                As Boolean
Private bInCtrl                     As Boolean

'for subclassing
Private z_ScMem                     As Long                'Thunk base address
Private z_Sc(64)                    As Long                'Thunk machine-code initialised here
Private z_Funk                      As Collection          'hWnd/thunk-address collection

'for usercontrol
Private m_lngHeight                 As Long
Private m_lngHeightAux              As Long
Private m_strCaption                As String
Private m_lngVertSpace              As Long
Private m_stdIcon                   As StdPicture
Private m_udtThemeColor             As jcThemeConst
Private m_strCurSysThemeName        As String
Private m_udtCaptionStyle           As jcCaptionConst
Private m_intMouseDown              As Integer
Private m_retPrevSize               As RECT
Private m_lngSpaceForIcon           As Long
Private m_intBtnIndex               As Integer
Private m_intPrevBtnIndex           As Integer
Private m_blnInTitleBar             As Boolean
Private m_lngBackColor              As OLE_COLOR
Private m_udtBackColorStyle         As jcBackColor
Private m_udtBorderStyle            As jcBorderStyle
Private m_blnMinButton              As Boolean
Private m_blnMaxButton              As Boolean
Private m_blnCloseButton            As Boolean
Private m_intWindowState            As FormWindowStateConstants
Private m_strLocalLanguage          As String
Private m_intIconSize               As Integer
Private m_intLeftBtn                As Integer
Private m_lngColorFromPrev          As OLE_COLOR
Private m_lngColorToPrev            As OLE_COLOR
Private m_lngColorFrom              As OLE_COLOR
Private m_lngColorTo                As OLE_COLOR
Private m_lngCustomColorFrom        As OLE_COLOR
Private m_lngCustomColorTo          As OLE_COLOR
Private m_blnChangeTop              As Boolean
Private m_frmPForm                  As Form
Private m_udtTitleBtn(0 To 2)       As CaptionBtn
Private m_intBtnPressed             As Integer
Private WithEvents m_picBottom      As PictureBox
Attribute m_picBottom.VB_VarHelpID = -1
Private WithEvents m_picRight       As PictureBox
Attribute m_picRight.VB_VarHelpID = -1
Private m_blnLoaded                 As Boolean
Private m_blnMoveable               As Boolean
Private m_blnChangeAllBackgrounds   As Boolean
Private m_frmPFormLoaded            As Boolean
Private m_blnFormLoaded             As Boolean
Private m_blnFormActivate           As Boolean
Private m_blnTitleBarShadow         As Boolean
Private m_strMenuCaption(0 To 6)    As String
Private m_intPressed                As Integer
Private m_blnWindowsNT              As Boolean

'=================================================================
'  usercontrol events
'=================================================================
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event jcCloseClick()
Public Event ReSize()
Public Event MenuItemSelected(MenuItem As Integer, MenuCaption As String)

'===================================
'  usercontrol events
'===================================

Private Sub UserControl_Initialize()
   
   Dim OS As OSVERSIONINFO

   OS.dwOSVersionInfoSize = Len(OS)
   Call GetVersionEx(OS)
   m_blnWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)

End Sub

Private Sub UserControl_Click()
    
    RaiseEvent Click

End Sub

Private Sub UserControl_DblClick()
    
    On Error GoTo Err_Proc
    
    If m_udtTitleBtn(1).Visible Then
        If m_intBtnPressed = vbLeftButton Then
            If m_blnInTitleBar = True Then
                CheckWindowState
            End If
        End If
    
        RaiseEvent DblClick
        m_intBtnPressed = 0
    
    End If
    
Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "UserControl_DblClick"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, _
                                  Shift As Integer)
    
    Dim intAltDown As Integer
    
    On Error GoTo Err_Proc
    
    If m_udtTitleBtn(0).Visible = True Then
        
        intAltDown = (Shift And vbAltMask) > 0
        
        If KeyCode = vbKeyF4 Then
            If intAltDown Then
                Unload m_frmPForm
            End If
        End If
    End If
    

Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "UserControl_KeyDown"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub UserControl_InitProperties()
    
    On Error GoTo Err_Proc
    
    m_lngCustomColorFrom = TranslateColor(&H404080)
    m_lngCustomColorTo = TranslateColor(&HC0E0FF)
    m_udtThemeColor = blue
    m_blnTitleBarShadow = True
    Call SetThemeColor
    
    m_lngHeight = 30
    m_udtCaptionStyle = Style1
    m_udtBackColorStyle = Default
    m_udtBorderStyle = Sizable
    m_blnMinButton = True
    m_blnMaxButton = True
    m_blnCloseButton = True
    m_intIconSize = 16
    m_lngBackColor = TranslateColor(&H8000000F)
    m_intWindowState = vbNormal
    m_strCaption = UserControl.Parent.Caption
    m_blnChangeAllBackgrounds = False
    m_blnMoveable = UserControl.Parent.Moveable
    
    SetupLanguageSystemMenu
    
    If m_udtCaptionStyle = Style2 And m_blnChangeTop = False Then
        m_blnChangeTop = True
    ElseIf m_udtCaptionStyle <> Style2 And m_blnChangeTop = True Then
        m_blnChangeTop = False
    End If
    
    With UserControl
        .BackColor = TranslateColor(&H8000000F)
        Set .Font = Ambient.Font
        .Font.Bold = True
        .Parent.BorderStyle = 0
        .Parent.ShowInTaskbar = True
        .Parent.MaxButton = True
        .Parent.MinButton = True
        .Parent.WindowState = vbNormal
        
        .Extender.Move 0, 0, .Parent.ScaleWidth, .Parent.ScaleHeight
        .Extender.Align = vbAlignTop
    End With
    
    BorderStyleSetup

Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "UserControl_InitProperties"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                   Shift As Integer, _
                                       X As Single, _
                                       Y As Single)
    
    On Error GoTo Err_Proc
    
    m_intBtnPressed = Button
    
    If Button = vbLeftButton Then
        If m_intBtnIndex <> -1 Then
            
            DrawCaptionBtns STA_PRESSED, m_lngColorFrom, m_intBtnIndex
            m_intPressed = m_intBtnIndex
        End If
    End If

    If Button = vbLeftButton And (Y < m_lngHeightAux And X > 2 And X < 8 + m_intIconSize) Then
        UserControl.PopupMenu SysMnu, , 3, m_lngHeightAux, MnuSyst(jcClose)
    End If

    RaiseEvent MouseDown(Button, Shift, X, Y)

Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "UserControl_MouseDown"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                       X As Single, _
                                       Y As Single)
    
    Dim lngReturnValue            As Long
    Dim lngFrmHwnd                As Long
        
    On Error GoTo Err_Proc
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    If m_blnLoaded = False Then
        
        Set m_picRight = m_frmPForm.Controls.Add("vb.PictureBox", "m_picRight")
        
        With m_picRight
            .AutoRedraw = True
            .ScaleMode = vbPixels
            .Appearance = 0
            .BorderStyle = 0
            .BackColor = m_lngColorFrom
            .Visible = True
            .ZOrder 0
        End With
        
        Set m_picBottom = m_frmPForm.Controls.Add("vb.PictureBox", "m_picBottom")
        
        With m_picBottom
            .AutoRedraw = True
            .ScaleMode = vbPixels
            .Appearance = 0
            .BorderStyle = 0
            .BackColor = m_lngColorFrom
            .Visible = True
            .ZOrder 0
        End With
        
        m_picRight.Move m_frmPForm.ScaleWidth - 4 * Screen.TwipsPerPixelX, (m_lngHeightAux + 9) * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, m_frmPForm.ScaleHeight - 10 * Screen.TwipsPerPixelY
        m_picBottom.Move 4 * Screen.TwipsPerPixelX, m_frmPForm.ScaleHeight - 4 * Screen.TwipsPerPixelY, m_frmPForm.ScaleWidth - 8 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
        
        m_blnLoaded = True
        
    End If
        
    If Y < m_lngHeightAux Then
        m_blnInTitleBar = True
    Else
        m_blnInTitleBar = False
    End If
    
    m_intBtnIndex = GetWhatButton(X, Y)
    
    If m_intBtnIndex <> -1 Then
        
        UserControl.Extender.TooltipText = m_udtTitleBtn(m_intBtnIndex).TooltipText
        If m_intPrevBtnIndex <> m_intBtnIndex Then
            
            DrawCaptionBtns STA_NORMAL, m_lngColorFrom, m_intPrevBtnIndex
            If m_intPressed <> -1 Then
                If m_intPressed = m_intBtnIndex Then
                    DrawCaptionBtns STA_PRESSED, m_lngColorFrom, m_intBtnIndex
                End If
            Else
                DrawCaptionBtns STA_OVER, m_lngColorFrom, m_intBtnIndex
            End If
            UserControl.MousePointer = 99
            m_intPrevBtnIndex = m_intBtnIndex
        End If
        
        Exit Sub
    
    Else
        
        UserControl.Extender.TooltipText = vbNullString
        If m_intPrevBtnIndex <> m_intBtnIndex Then
            UserControl.MousePointer = 0
            DrawCaptionBtns STA_NORMAL, m_lngColorFrom, m_intPrevBtnIndex
            m_intPrevBtnIndex = m_intBtnIndex
        End If
    
    End If
    
    If m_intWindowState = vbMaximized Then Exit Sub
        
    If Button = vbLeftButton Then
        If m_intMouseDown > 0 Then
            
            If X < 130 Then
                X = 130
                m_picRight.MousePointer = 0
            End If
            
            If Y < m_lngHeightAux + 15 Then
                Y = m_lngHeightAux + 15
                m_picBottom.MousePointer = 0
            End If
            
            Select Case m_intMouseDown
            
                Case 9
                    m_frmPForm.Width = (X + 5) * Screen.TwipsPerPixelX
            
                Case 8
                    m_frmPForm.Width = (X + 5) * Screen.TwipsPerPixelX
                    m_frmPForm.Height = (Y + 5) * Screen.TwipsPerPixelY
                    
                Case 7
                    m_frmPForm.Height = (Y + 5) * Screen.TwipsPerPixelY
                    
            End Select
            
            UserControl.Width = m_frmPForm.Width
            UserControl.Height = m_frmPForm.Height
            m_frmPForm.Refresh
        
        End If
        
        'move form by draging title bar
        If m_intPressed = -1 Then
            If Y < m_lngHeightAux Then
                
                If m_blnMoveable Then
                    If (X < m_udtTitleBtn(m_intLeftBtn).Left) Or Not m_udtTitleBtn(m_intLeftBtn).Visible Then
                        'Release capture
                        Call ReleaseCapture
                        lngFrmHwnd = UserControl.Parent.hwnd
                        lngReturnValue = SendMessage(lngFrmHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
                    End If
                End If
            
            End If
        End If
        
        SetRect m_retPrevSize, m_frmPForm.Left, m_frmPForm.Top, m_frmPForm.Width, m_frmPForm.Height
    Else
        
        m_picRight.MousePointer = 0
        m_picBottom.MousePointer = 0
    
    End If


Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "UserControl_MouseMove"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                 Shift As Integer, _
                                     X As Single, _
                                     Y As Single)
    
    On Error GoTo Err_Proc
    
    m_intMouseDown = 0
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
    If Button = vbLeftButton Then
        If m_intBtnIndex <> -1 Then
            
            If m_intPressed = m_intBtnIndex Or m_intPressed = -1 Then
            
                Select Case m_intBtnIndex
                    Case 0
                        RaiseEvent jcCloseClick
                        m_blnFormLoaded = False
                        Unload m_frmPForm
                    
                    Case 1
                        CheckWindowState
                    
                    Case 2
                        DrawCaptionBtns STA_NORMAL, m_lngColorFrom, m_intPrevBtnIndex
                        m_intPrevBtnIndex = -1
                        m_frmPForm.WindowState = vbMinimized
                
                End Select
            Else
                
                DrawCaptionBtns STA_OVER, m_lngColorFrom, m_intBtnIndex

            End If
        End If
    End If
    
    m_intPressed = -1
    
    If Button = vbRightButton Then
        If Y < m_lngHeightAux Then
            If (X < m_udtTitleBtn(m_intLeftBtn).Left) Or Not m_udtTitleBtn(m_intLeftBtn).Visible Then
                UserControl.PopupMenu SysMnu, , , , MnuSyst(jcClose)
            End If
        End If
    End If


Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "UserControl_MouseUp"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    'Read the properties from the property bag - also, a good place to start the subclassing (if we're running)
    
    On Error GoTo Err_Proc
    
    With PropBag
        m_intWindowState = .ReadProperty("WindowState", vbNormal)
        m_lngCustomColorFrom = .ReadProperty("ColorFrom", TranslateColor(&H404080))
        m_lngCustomColorTo = .ReadProperty("ColorTo", TranslateColor(&HC0E0FF))
        m_udtThemeColor = .ReadProperty("ThemeColor", Autodetect)
        m_udtBackColorStyle = .ReadProperty("BackColorStyle", Default)
        m_udtBorderStyle = .ReadProperty("BorderStyle", Sizable)
        m_blnCloseButton = .ReadProperty("CloseButton", True)
        m_blnMaxButton = .ReadProperty("MaxButton", True)
        m_blnMinButton = .ReadProperty("MinButton", True)
        m_lngBackColor = .ReadProperty("CustomBackColor", TranslateColor(&H8000000F))
        m_udtCaptionStyle = .ReadProperty("Style", Style1)
        m_intIconSize = .ReadProperty("IconSize", 16)
        m_blnChangeAllBackgrounds = .ReadProperty("ChangeAllBackgrounds", False)
        m_blnTitleBarShadow = .ReadProperty("TitleBarShadow", True)
    End With
    
    SetupLanguageSystemMenu
    Call SetThemeColor
    
    If m_blnMaxButton = False Then m_intWindowState = vbNormal
    m_strCaption = UserControl.Parent.Caption
    UserControl.Parent.WindowState = vbNormal
    Set m_frmPForm = UserControl.Parent
    m_blnFormLoaded = True
    m_intPressed = -1
    
    ControlsChangeTop 0
        
    If m_udtCaptionStyle = Style2 Then
        If m_blnChangeTop = False Then
            m_blnChangeTop = True
        End If
    Else
        If m_blnChangeTop = True Then
            m_blnChangeTop = False
        End If
    End If
    
    Select Case m_udtBackColorStyle
        Case Default
            UserControl.BackColor = TranslateColor(&H8000000F)
        Case Auto
            UserControl.BackColor = TranslateColor(m_lngColorTo)
        Case Custom
            UserControl.BackColor = TranslateColor(m_lngBackColor)
    End Select
    
    Set UserControl.Font = Ambient.Font
    UserControl.Font.Bold = True
    
    TitleBarHeightSetup
    
    If m_intMouseDown = 0 Then
        UserControl.Extender.Move 0, 0, UserControl.Parent.Width, UserControl.Parent.Height
        SetRect m_retPrevSize, m_frmPForm.Left, m_frmPForm.Top, m_frmPForm.Width, m_frmPForm.Height
    End If
    
    BorderStyleSetup
    
    If m_intWindowState = vbMinimized Then
        m_intPrevBtnIndex = -1
        m_frmPForm.WindowState = vbMinimized
    End If

    If Ambient.UserMode Then        'If we're not in design mode
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        
        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If
    
        If bTrack Then
            
            'OS supports mouse leave so subclass for it
            With UserControl
                'Start subclassing the UserControl
                sc_Subclass .hwnd
                sc_AddMsg .hwnd, WM_MOUSEMOVE
                sc_AddMsg .hwnd, WM_MOUSELEAVE
          
                'Subclass the parent form
                With .Parent
                    sc_Subclass .hwnd
                    sc_AddMsg .hwnd, WM_SETCURSOR
                    sc_AddMsg .hwnd, WM_SYSCOMMAND
                    sc_AddMsg .hwnd, WM_MOVE
                    sc_AddMsg .hwnd, WM_ACTIVATE
                    sc_AddMsg .hwnd, WM_SYSCOLORCHANGE

                End With
                
                m_blnMoveable = .Parent.Moveable
                .Parent.BorderStyle = 0
            End With
        End If
    
    End If

Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "UserControl_ReadProperties"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    On Error GoTo Err_Proc
    
    With PropBag
        .WriteProperty "ColorFrom", m_lngCustomColorFrom, &H404080
        .WriteProperty "ColorTo", m_lngCustomColorTo, &HC0E0FF
        .WriteProperty "ThemeColor", m_udtThemeColor, Autodetect
        .WriteProperty "BackColorStyle", m_udtBackColorStyle, Default
        .WriteProperty "BorderStyle", m_udtBorderStyle, Sizable
        .WriteProperty "CloseButton", m_blnCloseButton, True
        .WriteProperty "MaxButton", m_blnMaxButton, True
        .WriteProperty "MinButton", m_blnMinButton, True
        .WriteProperty "WindowState", m_intWindowState, vbNormal
        .WriteProperty "CustomBackColor", m_lngBackColor, &H8000000F
        .WriteProperty "Style", m_udtCaptionStyle, Style1
        .WriteProperty "IconSize", m_intIconSize, 16
        .WriteProperty "ChangeAllBackgrounds", m_blnChangeAllBackgrounds, False
        .WriteProperty "TitleBarShadow", m_blnTitleBarShadow, True
    End With


Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "UserControl_WriteProperties"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub UserControl_Resize()
    
    On Error GoTo Err_Proc
    
    UserControl.Cls
    
    If m_blnChangeAllBackgrounds Then
        
        Select Case BackColorStyle
            
            Case Auto
                ControlsChangeBackColor TranslateColor(m_lngColorTo)
            
            Case Custom
                ControlsChangeBackColor TranslateColor(m_lngBackColor)
            
            Case Else
                ControlsChangeBackColor TranslateColor(&H8000000F)
        End Select
    
    End If
    
    DrawTitleBar
    RaiseEvent ReSize
    
    If m_blnLoaded = True Then
        
        With m_picRight
            .Move m_frmPForm.ScaleWidth - 4 * Screen.TwipsPerPixelX, _
                (m_lngHeightAux + 9) * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, _
                m_frmPForm.ScaleHeight - (m_lngHeightAux + 10) * Screen.TwipsPerPixelY
            .BackColor = m_lngColorFrom
            .Refresh
        End With
        
        With m_picBottom
            .Move 4 * Screen.TwipsPerPixelX, m_frmPForm.ScaleHeight - 4 * Screen.TwipsPerPixelY, _
                m_frmPForm.ScaleWidth - 8 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
            .BackColor = m_lngColorFrom
            .Refresh
        End With
    
    End If

Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "UserControl_Resize"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub UserControl_Terminate()
    
    'The control is terminating - a good place to stop the subclasser
    On Error GoTo Catch
    m_blnFormLoaded = False
    
    'Terminate all subclassing
    sc_Terminate
    Set m_frmPForm = Nothing

Catch:

End Sub
    
'===================================
'  jcforms subs
'===================================

Private Sub MnuSyst_Click(Index As Integer)
    
    On Error GoTo Err_Proc
    
    Select Case Index
        
        Case 0, 2 '0- restore, 2- maximize
            CheckWindowState
        
        Case 1  'minimize
            DrawCaptionBtns STA_NORMAL, m_lngColorFrom, m_intPrevBtnIndex
            m_intPrevBtnIndex = -1
            m_frmPForm.WindowState = vbMinimized
        
        Case 4  'close
            m_blnFormLoaded = False
            Unload m_frmPForm
        
        Case 6  'always on top
            SetAlwaysOnTop Not MnuSyst(Index).Checked
        
        Case Else   'added menu items
            RaiseEvent MenuItemSelected(Index, MnuSyst(Index).Caption)
    
    End Select

Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "mnuSyst_Click"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub m_picBottom_MouseDown(Button As Integer, _
                                   Shift As Integer, _
                                       X As Single, _
                                       Y As Single)
    
    If m_intWindowState <> vbMaximized Then
        m_intMouseDown = m_picBottom.MousePointer
    End If
    
End Sub

Private Sub m_picBottom_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                       X As Single, _
                                       Y As Single)
    
    On Error GoTo Err_Proc
    
    If m_intWindowState = vbMaximized Then
        m_picBottom.MousePointer = 0
        Exit Sub
    End If
    
    If m_udtBorderStyle = Sizable Then
        If X > UserControl.ScaleWidth - 15 Then
            m_picBottom.MousePointer = 8
        Else
            m_picBottom.MousePointer = 7
        End If
    End If
    
    If m_intMouseDown > 0 Then
        Call UserControl_MouseMove(Button, Shift, X + m_picBottom.Left / Screen.TwipsPerPixelX, Y + m_picBottom.Top / Screen.TwipsPerPixelY)
    End If


Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "m_picBottom_MouseMove"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub m_picBottom_MouseUp(Button As Integer, _
                                 Shift As Integer, _
                                     X As Single, _
                                     Y As Single)
    
    m_intMouseDown = 0

End Sub

Private Sub m_picRight_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    
    If m_intWindowState <> vbMaximized Then
        m_intMouseDown = m_picRight.MousePointer
    End If
    
End Sub

Private Sub m_picRight_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    
    On Error GoTo Err_Proc
    
    If m_intWindowState = vbMaximized Then
        m_picRight.MousePointer = 0
        Exit Sub
    End If
    
    If m_udtBorderStyle = Sizable Then
        If Y > UserControl.ScaleHeight - m_lngHeightAux - 20 Then
            m_picRight.MousePointer = 8
        Else
            m_picRight.MousePointer = 9
        End If
    End If
    
    If m_intMouseDown > 0 Then
        Call UserControl_MouseMove(Button, Shift, X + m_picRight.Left / Screen.TwipsPerPixelX, Y + m_picRight.Top / Screen.TwipsPerPixelY)
    End If


Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "m_picRight_MouseMove"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub m_picRight_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                    X As Single, _
                                    Y As Single)
    
    m_intMouseDown = 0

End Sub
 

'============================
'   usercontrol properties
'============================
Public Property Get ThemeColor() As jcThemeConst
    
    ThemeColor = m_udtThemeColor

End Property

Public Property Let ThemeColor(ByVal vData As jcThemeConst)
    
    If m_udtThemeColor <> vData Then
        m_udtThemeColor = vData
        Call SetThemeColor
        Refresh
        PropertyChanged "ThemeColor"
    End If

End Property

Public Property Get Style() As jcCaptionConst
    
    Style = m_udtCaptionStyle

End Property

Public Property Let Style(ByVal vData As jcCaptionConst)
    
    On Error GoTo Err_Proc
    
    If m_udtCaptionStyle <> vData Then
        m_udtCaptionStyle = vData
        Refresh
        ControlsChangeTop
        PropertyChanged "Style"
    End If


Exit_Proc:
    On Error Resume Next
    Exit Property

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "Let Style"
    Err.Clear
    Resume Exit_Proc

End Property

Public Property Get BackColorStyle() As jcBackColor
    
    BackColorStyle = m_udtBackColorStyle

End Property

Public Property Let BackColorStyle(ByVal vData As jcBackColor)
    
    m_udtBackColorStyle = vData
    
    Refresh
    PropertyChanged "BackColorStyle"

End Property

Public Property Get MaxButton() As Boolean
    
    MaxButton = m_blnMaxButton

End Property

Public Property Let MaxButton(ByVal vData As Boolean)
    
    m_blnMaxButton = vData
        
    BorderStyleSetup
    Refresh
    PropertyChanged "MaxButton"

End Property

Public Property Get MinButton() As Boolean
    
    MinButton = m_blnMinButton

End Property

Public Property Let MinButton(ByVal vData As Boolean)
    
    m_blnMinButton = vData
        
    BorderStyleSetup
    Refresh
    PropertyChanged "MinButton"

End Property

Public Property Get CloseButton() As Boolean

    CloseButton = m_blnCloseButton

End Property

Public Property Let CloseButton(ByVal vData As Boolean)

    m_blnCloseButton = vData
    
    BorderStyleSetup
    Refresh
    PropertyChanged "CloseButton"

End Property

Public Property Get WindowState() As FormWindowStateConstants

    WindowState = m_intWindowState

End Property

Public Property Let WindowState(ByVal vData As FormWindowStateConstants)

    m_intWindowState = vData
    
    If Ambient.UserMode Then        'If we're not in design mode
        CheckWindowState -1
        BorderStyleSetup
    End If
    
    PropertyChanged "WindowState"

End Property

Public Property Get ChangeAllBackgrounds() As Boolean

   ChangeAllBackgrounds = m_blnChangeAllBackgrounds

End Property

Public Property Let ChangeAllBackgrounds(ByVal vNewValue As Boolean)

    m_blnChangeAllBackgrounds = vNewValue
    
    Refresh
    PropertyChanged "ChangeAllBackgrounds"
    
End Property

Public Property Get BorderStyle() As jcBorderStyle
    
    BorderStyle = m_udtBorderStyle

End Property

Public Property Let BorderStyle(ByVal vData As jcBorderStyle)
    
    m_udtBorderStyle = vData
    
    BorderStyleSetup
    Refresh
    PropertyChanged "BorderStyle"

End Property

Public Property Get CustomBackColor() As OLE_COLOR
    
    CustomBackColor = m_lngBackColor

End Property

Public Property Let CustomBackColor(ByVal vData As OLE_COLOR)
    
    On Error GoTo Err_Proc
    
    m_lngBackColor = vData
    
    If m_udtBackColorStyle = Custom Then
        Refresh
    End If
    
    PropertyChanged "CustomBackColor"


Exit_Proc:
    On Error Resume Next
    Exit Property

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "Let CustomBackColor"
    Err.Clear
    Resume Exit_Proc

End Property

Public Property Get ColorFrom() As OLE_COLOR
    
    ColorFrom = m_lngCustomColorFrom

End Property

Public Property Let ColorFrom(ByRef new_ColorFrom As OLE_COLOR)
    
    m_lngCustomColorFrom = TranslateColor(new_ColorFrom)
    
    If m_udtThemeColor = CustomTheme Then
        m_lngColorFrom = m_lngCustomColorFrom
    End If
    
    Refresh
    PropertyChanged "ColorFrom"

End Property

Public Property Get ColorTo() As OLE_COLOR
    
    ColorTo = m_lngCustomColorTo

End Property

Public Property Let ColorTo(ByRef new_ColorTo As OLE_COLOR)
    
    m_lngCustomColorTo = TranslateColor(new_ColorTo)
    
    If m_udtThemeColor = CustomTheme Then
        m_lngColorTo = m_lngCustomColorTo
    End If
    
    Refresh
    PropertyChanged "ColorTo"

End Property

Public Property Get IconSize() As Integer
    
    IconSize = m_intIconSize

End Property

Public Property Let IconSize(ByRef new_IconSize As Integer)
    
    If m_intIconSize <> new_IconSize Then
        m_intIconSize = new_IconSize
        TitleBarHeightSetup
        ControlsChangeTop 0
        Refresh
        PropertyChanged "IconSize"
    End If

End Property

Public Property Get TitleBarShadow() As Boolean

   TitleBarShadow = m_blnTitleBarShadow

End Property

Public Property Let TitleBarShadow(ByVal vNewValue As Boolean)

    m_blnTitleBarShadow = vNewValue
    
    Refresh
    PropertyChanged "TitleBarShadow"
    
End Property

'===============================
'   Functions and subroutines
'===============================

Private Sub APILineEx(ByRef lngHDC As Long, _
                      ByRef lngX1 As Long, _
                      ByRef lngY1 As Long, _
                      ByRef lngX2 As Long, _
                      ByRef lngY2 As Long, _
                      ByRef lngColor As Long)
    
    '=====================================
    'Use the API LineTo for Fast Drawing
    '=====================================
    
    Dim lngHPen         As Long
    Dim lngHPenOld      As Long
    
    On Error GoTo Err_Proc
    
    lngHPen = CreatePen(0, 1, lngColor)
    lngHPenOld = SelectObject(lngHDC, lngHPen)
    MoveToEx lngHDC, lngX1, lngY1, 0
    LineTo lngHDC, lngX2, lngY2
    SelectObject lngHDC, lngHPenOld
    DeleteObject lngHPen

Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "APILineEx"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Function ApiRectangle(ByVal lngHDC As Long, _
                              ByVal lngX As Long, _
                              ByVal lngY As Long, _
                              ByVal lngW As Long, _
                              ByVal lngH As Long, _
                              Optional ByVal lngColor As OLE_COLOR = -1)
                              
    '===================================
    'Draw a rectangle with api function
    '===================================
   
    Dim lngHPen         As Long
    Dim lngHPenOld      As Long
    
    On Error GoTo Err_Proc
    
    lngHPen = CreatePen(0, 1, lngColor)
    lngHPenOld = SelectObject(lngHDC, lngHPen)
    MoveToEx hDC, lngX, lngY, 0
    LineTo hDC, lngX + lngW, lngY
    LineTo hDC, lngX + lngW, lngY + lngH
    LineTo hDC, lngX, lngY + lngH
    LineTo hDC, lngX, lngY
    SelectObject lngHDC, lngHPenOld
    DeleteObject lngHPen

Exit_Proc:
    On Error Resume Next
    Exit Function

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "ApiRectangle"
    Err.Clear
    Resume Exit_Proc

End Function

Private Sub DrawGradientEx(ByRef lngHDC As Long, _
                           ByVal lngEndColor As Long, _
                           ByVal lngStartcolor As Long, _
                           ByVal lngX As Long, _
                           ByVal lngY As Long, _
                           ByVal lngX2 As Long, _
                           ByVal lngY2 As Long, _
                           Optional ByVal blnVertical As Boolean = True)
    
    '================================================================
    'Draw a Vertical or horizontal Gradient in the current HDC
    '================================================================
    
    Dim sngDR      As Single
    Dim sngDG      As Single
    Dim sngDB      As Single
    Dim sngSR      As Single
    Dim sngSG      As Single
    Dim sngSB      As Single
    Dim sngER      As Single
    Dim sngEG      As Single
    Dim sngEB      As Single
    Dim lngNI      As Long
    
    On Error Resume Next 'GoTo Exit_Proc 'Err_Proc
    
    sngSR = (lngStartcolor And &HFF)
    sngSG = (lngStartcolor \ &H100) And &HFF
    sngSB = (lngStartcolor And &HFF0000) / &H10000
    sngER = (lngEndColor And &HFF)
    sngEG = (lngEndColor \ &H100) And &HFF
    sngEB = (lngEndColor And &HFF0000) / &H10000
    
    If blnVertical Then
        sngDR = (sngSR - sngER) / lngY2
        sngDG = (sngSG - sngEG) / lngY2
        sngDB = (sngSB - sngEB) / lngY2
        
        For lngNI = 1 To lngY2 - 1
            APILineEx lngHDC, lngX, lngY + lngNI, lngX2, lngY + lngNI, RGB(sngER + (lngNI * sngDR), sngEG + (lngNI * sngDG), sngEB + (lngNI * sngDB))
        Next lngNI
    
    Else    'horizontal
        sngDR = (sngSR - sngER) / lngX2
        sngDG = (sngSG - sngEG) / lngX2
        sngDB = (sngSB - sngEB) / lngX2
        
        For lngNI = 1 To lngX2 - 1
            APILineEx lngHDC, lngX + lngNI, lngY, lngX + lngNI, lngY2, RGB(sngER + (lngNI * sngDR), sngEG + (lngNI * sngDG), sngEB + (lngNI * sngDB))
        Next lngNI
    
    End If


'Exit_Proc:
'    On Error Resume Next
'    Exit Sub
'
'Err_Proc:
'    Err_Handler True, Err.Number, Err.Description, "jcForms", "DrawGradientEx"
'    Err.Clear
'    Resume Exit_Proc

End Sub

Private Function BlendColors(ByVal lngColorFrom As Long, _
                             ByVal lngColorTo As Long, _
                             Optional ByVal alpha As Long = 128) As Long

    Dim lngSrcR         As Long
    Dim lngSrcG         As Long
    Dim lngSrcB         As Long
    Dim lngDstR         As Long
    Dim lngDstG         As Long
    Dim lngDstB         As Long
   
    lngSrcR = lngColorFrom And &HFF
    lngSrcG = (lngColorFrom And &HFF00&) \ &H100&
    lngSrcB = (lngColorFrom And &HFF0000) \ &H10000
    lngDstR = lngColorTo And &HFF
    lngDstG = (lngColorTo And &HFF00&) \ &H100&
    lngDstB = (lngColorTo And &HFF0000) \ &H10000
     
   
    BlendColors = RGB( _
                ((lngSrcR * alpha) / 255) + ((lngDstR * (255 - alpha)) / 255), _
                ((lngSrcG * alpha) / 255) + ((lngDstG * (255 - alpha)) / 255), _
                ((lngSrcB * alpha) / 255) + ((lngDstB * (255 - alpha)) / 255))

End Function

Private Sub DrawText(ByVal hDC As Long, _
                     ByVal lpString As String, _
                     ByVal nCount As Long, _
                     ByRef lpRect As RECT, _
                     ByVal wFormat As Long)

    '================================================================
    '* draws the text with Unicode support based on OS version.
    '* Thanks to Richard Mewett.
    '================================================================

    If m_blnWindowsNT Then
        DrawTextW hDC, StrPtr(lpString), nCount, lpRect, wFormat
    Else
        DrawTextA hDC, lpString, nCount, lpRect, wFormat
    End If

End Sub


Private Function TranslateColor(ByVal lngColor As Long) As Long
    
    '================================
    'System color code to long rgb
    '================================
    
    If OleTranslateColor(lngColor, 0, TranslateColor) Then
        TranslateColor = -1
    End If

End Function

Private Sub DrawGradientInRectangle(ByRef lngHDC As Long, _
                                    ByVal lngStartcolor As Long, _
                                    ByVal lngEndColor As Long, _
                                    ByRef RetDefRect As RECT, _
                                    ByVal GradientType As jcGradConst, _
                                    Optional ByVal blnDrawBorder As Boolean = False, _
                                    Optional ByVal lBorderColor As Long = vbBlack, _
                                    Optional ByVal LightCenter As Double = 2.01)

    
    '==============================================================
    'Draws rectangle with vertical, horizontal, vertical cylinder
    'or horizontal cylinder gradients
    '==============================================================
    
    On Error GoTo Err_Proc
    
    Select Case GradientType
        Case VerticalGradient
            DrawGradientEx lngHDC, lngEndColor, lngStartcolor, RetDefRect.Left, _
            RetDefRect.Top, RetDefRect.Right + RetDefRect.Left, RetDefRect.Bottom, True
        
        Case horizontalGradient
            DrawGradientEx lngHDC, lngEndColor, lngStartcolor, RetDefRect.Left, _
            RetDefRect.Top, RetDefRect.Right, RetDefRect.Bottom + RetDefRect.Top, False
        
        Case VCylinderGradient
            DrawGradCilinder lngHDC, lngStartcolor, lngEndColor, RetDefRect, True, LightCenter
        
        Case HCylinderGradient
            DrawGradCilinder lngHDC, lngStartcolor, lngEndColor, RetDefRect, False, LightCenter
    
    End Select
    
    If blnDrawBorder Then
        ApiRectangle lngHDC, RetDefRect.Left, RetDefRect.Top, RetDefRect.Right, _
        RetDefRect.Bottom, lBorderColor
    End If


Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "DrawGradientInRectangle"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub DrawGradCilinder(ByRef lngHDC As Long, _
                             ByVal lngStartcolor As Long, _
                             ByVal lngEndColor As Long, _
                             ByRef retR As RECT, _
                             Optional ByVal blnVertical As Boolean = True, _
                             Optional ByVal dblLightCenter As Double = 2.01)

    '=========================================================================
    'Draws rectangle with vertical cylinder or horizontal cylinder gradients
    '=========================================================================
    
    On Error GoTo Err_Proc

    If dblLightCenter <= 1# Then
        dblLightCenter = 1.01
    End If
    
    If blnVertical Then
        DrawGradientEx lngHDC, lngStartcolor, lngEndColor, retR.Left, retR.Top, _
            retR.Right + retR.Left, retR.Bottom / dblLightCenter, True
        
        DrawGradientEx lngHDC, lngEndColor, lngStartcolor, retR.Left, _
            retR.Top + retR.Bottom / dblLightCenter - 1, retR.Right + retR.Left, _
            (dblLightCenter - 1) * retR.Bottom / dblLightCenter + 1, True
    
    Else
        DrawGradientEx lngHDC, lngStartcolor, lngEndColor, retR.Left, retR.Top, _
            retR.Right / dblLightCenter, retR.Bottom + retR.Top, False
        
        DrawGradientEx lngHDC, lngEndColor, lngStartcolor, _
            retR.Left + retR.Right / dblLightCenter - 1, retR.Top, _
            (dblLightCenter - 1) * retR.Right / dblLightCenter + 1, _
            retR.Bottom + retR.Top, False
    End If


Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "DrawGradCilinder"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub PaintShpInBar(ByVal lngColorA As Long, _
                          ByVal lngColorB As Long, _
                          ByVal m_lngHeight As Long)
    
    Dim intI                As Integer
    Dim intLeft             As Integer
    Dim intTop              As Integer
    Dim retR                As RECT
    
    On Error GoTo Err_Proc
    
    Const intHSpace         As Integer = 2    'space between shapes
    Const intNShp           As Integer = 9    'number of points
    Const lngRHeight        As Long = 2       'shape height
    Const lngRWidth         As Long = 2       'shape width
   
    'x and y shape coordinates
    intLeft = (UserControl.ScaleWidth - intNShp * lngRWidth - (intNShp - 1) * intHSpace) / 2
    intTop = (m_lngHeight - lngRHeight) / 2
    
    For intI = 0 To intNShp - 1
        SetRect retR, intLeft + intI * intHSpace + intI * lngRWidth + 1, intTop + 1, 1, 1
        ApiRectangle UserControl.hDC, retR.Left, retR.Top, retR.Right, retR.Bottom, lngColorA
        SetRect retR, intLeft + intI * intHSpace + intI * lngRWidth, intTop, 1, 1
        ApiRectangle UserControl.hDC, retR.Left, retR.Top, retR.Right, retR.Bottom, lngColorB
    Next intI

Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "PaintShpInBar"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub DrawTitleBar()
    
    Dim retR                    As RECT
    Dim retCptnR                As RECT
    Dim retCptnR_Aux            As String
    Dim lngTextDrawParams       As Long
    Dim intI                    As Integer
    Dim intJ                    As Integer
    Dim intBtnHeight            As Integer
    Dim intLeft                 As Integer
    
    Dim lngBorderColor          As Long
    Const intHSpace             As Integer = 6

    '==================================
    'Draws a title bar of the jcForms
    '==================================
    
    On Error GoTo Err_Proc
    
    If m_udtTitleBtn(0).Visible = False Then
        intBtnHeight = 18
    Else
        intBtnHeight = m_udtTitleBtn(0).Height
    End If
        
    If m_udtTitleBtn(m_intLeftBtn).Visible = False Then
        intLeft = UserControl.ScaleWidth
    Else
        intLeft = m_udtTitleBtn(m_intLeftBtn).Left
    End If
    
'    Paint title bar
    If UserControl.ScaleWidth > m_lngSpaceForIcon + 1 Then
        
        Select Case m_udtCaptionStyle
            
            Case Style1
                m_lngHeightAux = m_lngHeight
                
                With UserControl
                    .ForeColor = vbBlack
                    SetRect retR, m_lngSpaceForIcon, -1, .ScaleWidth + 1, m_lngHeightAux + 1
                    DrawGradientInRectangle .hDC, m_lngColorFrom, m_lngColorTo, retR, VCylinderGradient, False, m_lngColorFrom, 5.01
                End With
                
            Case Style2
                m_lngHeightAux = m_lngHeight + 10
                
                With UserControl
                    .ForeColor = vbBlack
                    SetRect retR, m_lngSpaceForIcon, -2, .ScaleWidth + 1, 10
                    DrawGradientInRectangle .hDC, m_lngColorFrom, m_lngColorTo, retR, VCylinderGradient, False, m_lngColorFrom, 5.01
                    SetRect retR, m_lngSpaceForIcon, 7, .ScaleWidth + 1, m_lngHeightAux - 7
                    DrawGradientInRectangle .hDC, m_lngColorFrom, m_lngColorTo, retR, VCylinderGradient, False, m_lngColorFrom, 5.01
                End With
                
            Case Style3
                'Draw header
                m_lngHeightAux = m_lngHeight
                
                With UserControl
                    .ForeColor = vbWhite
                    SetRect retR, m_lngSpaceForIcon, -2, .ScaleWidth + 1, 10
                    DrawGradientInRectangle .hDC, m_lngColorFrom, m_lngColorTo, retR, VerticalGradient, False, m_lngColorFrom, 5.01
                    SetRect retR, m_lngSpaceForIcon, 7, .ScaleWidth + 1, m_lngHeightAux - 7
                    DrawGradientInRectangle .hDC, m_lngColorFrom, m_lngColorFrom, retR, VCylinderGradient, False, m_lngColorFrom, 5.01
                End With
                
            Case Style4
                m_lngHeightAux = m_lngHeight
                
                With UserControl
                    .ForeColor = vbBlack
                    SetRect retR, m_lngSpaceForIcon, -1, .ScaleWidth + 1, m_lngHeightAux + 1
                    DrawGradientInRectangle .hDC, m_lngColorFrom, m_lngColorTo, retR, VerticalGradient, False, m_lngColorFrom, 5.01
                End With
                
            Case Style5
                'Draw header
                m_lngHeightAux = m_lngHeight
                
                With UserControl
                    .ForeColor = vbWhite
                    SetRect retR, m_lngSpaceForIcon, -2, .ScaleWidth + 1, 7
                    DrawGradientInRectangle .hDC, m_lngColorFrom, m_lngColorTo, retR, VCylinderGradient, False, m_lngColorFrom, 5.01
                    SetRect retR, m_lngSpaceForIcon, 4, .ScaleWidth + 1, m_lngHeightAux - 4
                    DrawGradientInRectangle .hDC, m_lngColorTo, m_lngColorFrom, retR, VCylinderGradient, False, m_lngColorFrom, 1.01
                End With
                
            Case Style6
                'Draw header
                m_lngHeightAux = m_lngHeight
                
                With UserControl
                    .ForeColor = vbWhite
                    SetRect retR, m_lngSpaceForIcon, -2, .ScaleWidth + 1, 12
                    DrawGradientInRectangle .hDC, m_lngColorTo, BlendColors(m_lngColorTo, BlendColors(m_lngColorFrom, vbBlack)), retR, VCylinderGradient, False, m_lngColorFrom, 1.01
                    SetRect retR, m_lngSpaceForIcon, 8, .ScaleWidth + 1, m_lngHeightAux - 8
                    DrawGradientInRectangle .hDC, m_lngColorFrom, m_lngColorFrom, retR, VCylinderGradient, False, m_lngColorFrom, 7.01
                    APILineEx .hDC, m_lngSpaceForIcon, 8, .ScaleWidth + 1, 8, BlendColors(m_lngColorFrom, m_lngColorTo)
                    APILineEx .hDC, m_lngSpaceForIcon, 0, .ScaleWidth + 1, 0, BlendColors(vbWhite, m_lngColorTo)
                End With
                
        End Select
    
    End If
    
    'set caption rect
    If m_udtCaptionStyle <> Style2 Then
        SetRect retCptnR, intHSpace * 1.5, 0 + 4, intLeft - 2 * intHSpace, m_lngHeightAux
        m_lngVertSpace = (m_lngHeightAux - intBtnHeight) / 2
    Else
        SetRect retCptnR, intHSpace * 1.5, 0 + 11, intLeft - 2 * intHSpace, m_lngHeightAux + 3
        PaintShpInBar m_lngColorTo, BlendColors(m_lngColorFrom, vbBlack), 7
        m_lngVertSpace = (m_lngHeightAux - intBtnHeight) / 2 + 3
    End If

    'Draw borders
    With UserControl
        lngBorderColor = m_lngColorFrom
        SetRect retR, 0, m_lngHeightAux + 1, .ScaleWidth - 1, .ScaleHeight - m_lngHeightAux - 2
        ApiRectangle .hDC, retR.Left, retR.Top, retR.Right, retR.Bottom, lngBorderColor
        SetRect retR, 1, m_lngHeightAux + 1, .ScaleWidth - 3, .ScaleHeight - m_lngHeightAux - 3
        ApiRectangle .hDC, retR.Left, retR.Top, retR.Right, retR.Bottom, lngBorderColor
        SetRect retR, 2, m_lngHeightAux + 1, .ScaleWidth - 5, .ScaleHeight - m_lngHeightAux - 4
        ApiRectangle .hDC, retR.Left, retR.Top, retR.Right, retR.Bottom, BlendColors(lngBorderColor, vbBlack)
    End With
    
    'set caption buttons coordinates
    intJ = 0
    
    For intI = 0 To 2
        
        If m_udtTitleBtn(intI).Visible = True Then
            m_udtTitleBtn(intI).Top = 2 + m_lngVertSpace
            m_udtTitleBtn(intI).Left = UserControl.ScaleWidth - 3 - 21 * (intJ + 1)
            intJ = intJ + 1
        End If
    
    Next intI
    
    With UserControl
        'Titlebar Shadow
        If m_blnTitleBarShadow Then
            SetRect retR, 3, m_lngHeightAux, UserControl.ScaleWidth - 6, 8
            DrawGradientInRectangle .hDC, TranslateColor(.BackColor), BlendColors(TranslateColor(.BackColor), vbBlack), retR, VerticalGradient, False, m_lngColorFrom, 2.01
            APILineEx UserControl.hDC, -1, m_lngHeightAux, .ScaleWidth + 1, m_lngHeightAux, m_lngColorFrom
            APILineEx UserControl.hDC, 5, m_lngHeightAux, .ScaleWidth - 5, m_lngHeightAux, BlendColors(m_lngColorFrom, vbBlack)
        Else
            If m_udtCaptionStyle <> Style1 Then
                APILineEx UserControl.hDC, -1, m_lngHeightAux, .ScaleWidth + 1, m_lngHeightAux, m_lngColorFrom
            End If
        End If
        
        'rounded borders
        SetPixel .hDC, 2, m_lngHeightAux + 1, lngBorderColor
        SetPixel .hDC, 2, m_lngHeightAux + 2, lngBorderColor
        SetPixel .hDC, 3, m_lngHeightAux + 1, lngBorderColor
        SetPixel .hDC, 4, m_lngHeightAux + 1, BlendColors(lngBorderColor, vbBlack)
        SetPixel .hDC, 3, m_lngHeightAux + 2, BlendColors(lngBorderColor, vbBlack)
    
        SetPixel .hDC, UserControl.ScaleWidth - 3, m_lngHeightAux + 1, lngBorderColor
        SetPixel .hDC, UserControl.ScaleWidth - 3, m_lngHeightAux + 2, lngBorderColor
        SetPixel .hDC, UserControl.ScaleWidth - 4, m_lngHeightAux + 1, lngBorderColor
        SetPixel .hDC, UserControl.ScaleWidth - 5, m_lngHeightAux + 1, BlendColors(lngBorderColor, vbBlack)
        SetPixel .hDC, UserControl.ScaleWidth - 4, m_lngHeightAux + 2, BlendColors(lngBorderColor, vbBlack)
    End With
    
    If Not (m_stdIcon Is Nothing) Then
        OffsetRect retCptnR, intHSpace + m_intIconSize, 0
        retCptnR.Right = retCptnR.Right - m_intIconSize
    End If

    On Error Resume Next
    
    'Draw caption
    If Ambient.UserMode = True Then
        If UserControl.Parent.Caption <> "" Then
            If m_strCaption <> UserControl.Parent.Caption Then m_strCaption = UserControl.Parent.Caption
        End If
    End If
    
    If LenB(m_strCaption) <> 0 Then
        
        retCptnR_Aux = TrimWord(m_strCaption, retCptnR.Right - retCptnR.Left)

        'Draw text
        lngTextDrawParams = DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
        UserControl.ForeColor = m_lngColorTo
        DrawText UserControl.hDC, retCptnR_Aux, Len(retCptnR_Aux), retCptnR, lngTextDrawParams

        If m_udtCaptionStyle <> Style3 And m_udtCaptionStyle <> Style6 Then
            OffsetRect retCptnR, -1, -1
            UserControl.ForeColor = BlendColors(m_lngColorFrom, vbBlack)
            DrawText UserControl.hDC, retCptnR_Aux, Len(retCptnR_Aux), retCptnR, lngTextDrawParams
        End If
    
    End If
    
    PaintCaptionBtns -1


Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    If Err.Number = 398 Then
    Else
        Err_Handler True, Err.Number, Err.Description, "jcForms", "DrawTitleBar"
    End If
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub DrawTitleBarInitial()
    
    Dim retR                    As RECT
    Const intHSpace             As Integer = 6
    
    '============================================
    'Draws a title bar of the jcForms first time
    '============================================
    
    On Error GoTo Err_Proc
    
    If UserControl.Parent.Icon = 0 Then
        Set m_stdIcon = Nothing
        m_lngSpaceForIcon = 0
    Else
        Set m_stdIcon = UserControl.Parent.Icon
        m_lngSpaceForIcon = UserControl.ScaleWidth - 1
    End If
    
    Set UserControl.Picture = Nothing
    
    Select Case m_udtCaptionStyle
        
        Case Style1
            m_lngHeightAux = m_lngHeight
            UserControl.ForeColor = vbBlack
            SetRect retR, -1, -1, m_lngSpaceForIcon + 2, m_lngHeightAux + 1
            DrawGradientInRectangle UserControl.hDC, m_lngColorFrom, m_lngColorTo, retR, VCylinderGradient, True, m_lngColorFrom, 5.01
        
        Case Style2
            m_lngHeightAux = m_lngHeight + 10
            UserControl.ForeColor = vbBlack
            SetRect retR, -1, -2, m_lngSpaceForIcon + 2, 10
            DrawGradientInRectangle UserControl.hDC, m_lngColorFrom, m_lngColorTo, retR, VCylinderGradient, False, m_lngColorFrom, 5.01
            SetRect retR, -1, 7, m_lngSpaceForIcon + 2, m_lngHeightAux - 7
            DrawGradientInRectangle UserControl.hDC, m_lngColorFrom, m_lngColorTo, retR, VCylinderGradient, False, m_lngColorFrom, 5.01
        
        Case Style3
            'Draw header
            m_lngHeightAux = m_lngHeight
            UserControl.ForeColor = vbWhite
            SetRect retR, -1, -2, m_lngSpaceForIcon + 2, 10
            DrawGradientInRectangle UserControl.hDC, m_lngColorFrom, m_lngColorTo, retR, VerticalGradient, False, m_lngColorFrom, 5.01
            SetRect retR, -1, 7, m_lngSpaceForIcon + 2, m_lngHeightAux - 7
            DrawGradientInRectangle UserControl.hDC, m_lngColorFrom, m_lngColorFrom, retR, VCylinderGradient, False, m_lngColorFrom, 5.01
        
        Case Style4
            m_lngHeightAux = m_lngHeight
            UserControl.ForeColor = vbBlack
            SetRect retR, -1, -1, m_lngSpaceForIcon + 2, m_lngHeightAux + 1
            DrawGradientInRectangle UserControl.hDC, m_lngColorFrom, m_lngColorTo, retR, VerticalGradient, False, m_lngColorFrom, 5.01
    
        Case Style5
            'Draw header
            m_lngHeightAux = m_lngHeight
            UserControl.ForeColor = vbBlack
            SetRect retR, -1, -2, m_lngSpaceForIcon + 2, 7
            DrawGradientInRectangle UserControl.hDC, m_lngColorFrom, m_lngColorTo, retR, VCylinderGradient, False, m_lngColorFrom, 5.01
            SetRect retR, -1, 4, m_lngSpaceForIcon + 2, m_lngHeightAux - 4
            DrawGradientInRectangle UserControl.hDC, m_lngColorTo, m_lngColorFrom, retR, VCylinderGradient, False, m_lngColorFrom, 1.01
        
        Case Style6
            'Draw header
            m_lngHeightAux = m_lngHeight
            UserControl.ForeColor = vbBlack
            SetRect retR, -1, -2, m_lngSpaceForIcon + 2, 12
            DrawGradientInRectangle UserControl.hDC, m_lngColorTo, BlendColors(m_lngColorTo, BlendColors(m_lngColorFrom, vbBlack)), retR, VCylinderGradient, False, m_lngColorFrom, 1.01
            SetRect retR, -1, 8, m_lngSpaceForIcon + 2, m_lngHeightAux - 8
            DrawGradientInRectangle UserControl.hDC, m_lngColorFrom, m_lngColorFrom, retR, VCylinderGradient, False, m_lngColorFrom, 7.01
            APILineEx UserControl.hDC, -1, 0, m_lngSpaceForIcon + 2, 0, BlendColors(vbWhite, m_lngColorTo)
            APILineEx UserControl.hDC, -1, 8, m_lngSpaceForIcon + 2, 8, BlendColors(m_lngColorFrom, m_lngColorTo)
    End Select
    
    'draw picture
    If Not (m_stdIcon Is Nothing) Then
        If m_udtCaptionStyle = Style2 Then
            UserControl.PaintPicture m_stdIcon, intHSpace * 1.5, (m_lngHeightAux - 8 - m_intIconSize) / 2 + 10, m_intIconSize, m_intIconSize
        Else
            UserControl.PaintPicture m_stdIcon, intHSpace * 1.5, (m_lngHeightAux - m_intIconSize) / 2 + 2, m_intIconSize, m_intIconSize
        End If
    End If
    
    Set UserControl.Picture = UserControl.Image

Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    If Err.Number = 398 Then
    Else
        Err_Handler True, Err.Number, Err.Description, "jcForms", "DrawTitleBarInitial"
    End If
    Err.Clear
    Resume Exit_Proc

End Sub

Private Function TrimWord(ByVal strCaption As String, _
                          ByVal lngWidth As Long) As String

    Dim lngLenOfText            As Long

    On Error GoTo Err_Proc
    
    TrimWord = strCaption
    
    If TextWidth(strCaption) > lngWidth Then
        lngLenOfText = Len(strCaption)
        
        Do Until TextWidth(TrimWord & "...") <= lngWidth Or lngLenOfText = 0
            lngLenOfText = lngLenOfText - 1
            TrimWord = Left(TrimWord, lngLenOfText)
        Loop
        
        If lngLenOfText = 0 Then
            TrimWord = Empty
        Else
            TrimWord = TrimWord & "..."
        End If
        
    End If

Exit_Proc:
    On Error Resume Next
    Exit Function

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "TrimWord"
    Err.Clear
    Resume Exit_Proc

End Function

Private Sub SetThemeColor()

    On Error GoTo Err_Proc
    
    '==================================
    'sets a jcForms theme
    '==================================
    
    If m_udtThemeColor = Autodetect Then
       GetGradientColor UserControl.hwnd
    Else
       SetDefaultThemeColor m_udtThemeColor
    End If
                
    If Ambient.UserMode Then        'If we're not in design mode
        If m_blnFormActivate Then
            m_lngColorFrom = m_lngColorFromPrev
            m_lngColorTo = m_lngColorToPrev
        Else
            m_lngColorFrom = BlendColors(m_lngColorFromPrev, vbWhite, 190)
            m_lngColorTo = BlendColors(m_lngColorToPrev, vbWhite, 190)
        End If
    Else
        m_lngColorFrom = m_lngColorFromPrev
        m_lngColorTo = m_lngColorToPrev
    End If
    
Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "SetThemeColor"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub SetDefaultThemeColor(lngThemeType As Long)

    On Error GoTo Err_Proc

    Select Case lngThemeType
        
        Case 0 'NormalColor
            m_lngColorFromPrev = RGB(19, 97, 156)
            m_lngColorToPrev = RGB(221, 236, 254)
        
        Case 1 'Metallic
            m_lngColorFromPrev = RGB(95, 95, 111)
            m_lngColorToPrev = RGB(244, 244, 251)
        
        Case 2 'HomeStead
            m_lngColorFromPrev = RGB(63, 93, 56)
            m_lngColorToPrev = RGB(247, 249, 225)
        
        Case 3 'Visual2005
            m_lngColorFromPrev = RGB(98, 107, 72)
            m_lngColorToPrev = RGB(248, 248, 242)
        
        Case 4 'Norton2004
            m_lngColorFromPrev = RGB(117, 91, 30)
            m_lngColorToPrev = RGB(255, 239, 165)
        
        Case 5 'CustomTheme
            m_lngColorFromPrev = m_lngCustomColorFrom
            m_lngColorToPrev = m_lngCustomColorTo
    
    End Select


Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "SetDefaultThemeColor"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub GetGradientColor(lhWnd As Long)

    On Error GoTo Err_Proc
    
    GetThemeName lhWnd
    
    If AppThemed Then   '/Check if themed.
        
        Select Case m_strCurSysThemeName
            
            Case "NormalColor"
                m_lngColorFromPrev = RGB(19, 97, 156)
                m_lngColorToPrev = RGB(221, 236, 254)
            
            Case "Metallic"
                m_lngColorFromPrev = RGB(95, 95, 111)
                m_lngColorToPrev = RGB(244, 244, 251)
            
            Case "HomeStead"
                m_lngColorFromPrev = RGB(63, 93, 56)
                m_lngColorToPrev = RGB(247, 249, 225)
        
        End Select
    
    Else    'APPTHEMED = FALSE/0
        m_lngColorFromPrev = RGB(19, 97, 156)
        m_lngColorToPrev = RGB(221, 236, 254)
    End If


Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "GetGradientColor"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Function AppThemed() As Boolean
    
    '===========================================
    'Determines If The Current Window is Themed
    '===========================================

    On Error Resume Next
    AppThemed = IsAppThemed()
    On Error GoTo 0

End Function

Private Sub GetThemeName(lngHwnd As Long)
    
    '========================================
    'Returns The current Windows Theme Name
    '========================================

    Dim lngTheme            As Long
    Dim stringShellStyle    As String
    Dim stringThemeFile     As String
    Dim lngPtrThemeFile     As Long
    Dim lngPtrColorName     As Long
    Dim lngPos                As Long
    
    On Error Resume Next
    
    lngTheme = OpenThemeData(lngHwnd, StrPtr("ExplorerBar"))
    
    If Not lngTheme = 0 Then
        
        ReDim bThemeFile(0 To 260 * 2) As Byte
        lngPtrThemeFile = VarPtr(bThemeFile(0))
        
        ReDim bColorName(0 To 260 * 2) As Byte
        lngPtrColorName = VarPtr(bColorName(0))
        
        GetCurrentThemeName lngPtrThemeFile, 260, lngPtrColorName, 260, 0, 0
        stringThemeFile = bThemeFile
        lngPos = InStr(stringThemeFile, vbNullChar)
        
        If lngPos > 1 Then
            stringThemeFile = Left$(stringThemeFile, lngPos - 1)
        End If
        
        m_strCurSysThemeName = bColorName
        lngPos = InStr(m_strCurSysThemeName, vbNullChar)
        
        If lngPos > 1 Then
            m_strCurSysThemeName = Left$(m_strCurSysThemeName, lngPos - 1)
        End If
        
        stringShellStyle = stringThemeFile
        
        For lngPos = Len(stringThemeFile) To 1 Step -1
            If (Mid$(stringThemeFile, lngPos, 1) = "\") Then
                stringShellStyle = Left$(stringThemeFile, lngPos)
                Exit For
            End If
        Next lngPos
        
        stringShellStyle = stringShellStyle & "Shell\" & m_strCurSysThemeName & "\ShellStyle.dll"
        CloseThemeData lngTheme
    
    Else
        m_strCurSysThemeName = "Classic"
    
    End If
    
    On Error GoTo 0

End Sub

Private Sub DrawCaptionBtns(ByVal udtBtnState As jcBtnState, _
                            ByVal lngColor As OLE_COLOR, _
                            ByVal Index As Integer)
    
    Dim retBtn              As RECT
    Dim lngBtnLeft          As Long
    Dim lngBtnTop           As Long
    Dim lngBtnWidth         As Long
    Dim lngBtnHeight        As Long
    Dim lngBorderColor      As OLE_COLOR
    
    '======================================================
    'Draws Minimize, Maximize (Restore) and Close buttons
    'in the Title bar
    '======================================================
    
    On Error GoTo Err_Proc
    
    If Index = -1 Then
        
        With m_udtTitleBtn(m_intLeftBtn)
            lngBtnLeft = .Left
            lngBtnTop = .Top
            lngBtnWidth = .Width
            lngBtnHeight = .Height
        End With
    
    Else
        
        With m_udtTitleBtn(Index)
            lngBtnLeft = .Left
            lngBtnTop = .Top
            lngBtnWidth = .Width
            lngBtnHeight = .Height
        End With
        
    End If
    
    Select Case m_udtCaptionStyle

        Case Style3, Style6
            lngBorderColor = vbBlack

        Case Else
            lngBorderColor = BlendColors(m_lngColorFrom, vbBlack)

    End Select
    
    Select Case udtBtnState
    
        Case STA_NORMAL, STA_PRESSED
        
            Select Case m_udtCaptionStyle
                
                Case Style1
                    SetRect retBtn, lngBtnLeft - 1, -1, lngBtnWidth + 4, m_lngHeightAux + 1
                    DrawGradientInRectangle UserControl.hDC, m_lngColorFrom, m_lngColorTo, retBtn, VCylinderGradient, False, lngColor, 5.01
                
                Case Style2
                    SetRect retBtn, lngBtnLeft - 1, 7, lngBtnWidth + 4, m_lngHeightAux - 7
                    DrawGradientInRectangle UserControl.hDC, m_lngColorFrom, m_lngColorTo, retBtn, VCylinderGradient, False, lngColor, 5.01
                
                Case Style3
                    SetRect retBtn, lngBtnLeft - 1, -2, lngBtnWidth + 4, 10
                    DrawGradientInRectangle UserControl.hDC, m_lngColorFrom, m_lngColorTo, retBtn, VerticalGradient, False, m_lngColorFrom, 5.01
                    SetRect retBtn, lngBtnLeft - 1, 7, lngBtnWidth + 4, m_lngHeightAux - 7
                    DrawGradientInRectangle UserControl.hDC, m_lngColorFrom, m_lngColorFrom, retBtn, VCylinderGradient, False, m_lngColorFrom, 5.01
                    lngColor = m_lngColorTo
                
                Case Style4
                    SetRect retBtn, lngBtnLeft - 1, -1, lngBtnWidth + 4, m_lngHeightAux + 1
                    DrawGradientInRectangle UserControl.hDC, m_lngColorFrom, m_lngColorTo, retBtn, VerticalGradient, False, lngColor, 5.01
            
                Case Style5
                    SetRect retBtn, lngBtnLeft - 1, -2, lngBtnWidth + 4, 7
                    DrawGradientInRectangle UserControl.hDC, m_lngColorFrom, m_lngColorTo, retBtn, VCylinderGradient, False, m_lngColorFrom, 5.01
                    SetRect retBtn, lngBtnLeft - 1, 4, lngBtnWidth + 4, m_lngHeightAux - 4
                    DrawGradientInRectangle UserControl.hDC, m_lngColorTo, m_lngColorFrom, retBtn, VCylinderGradient, False, m_lngColorFrom, 1.01
            
                Case Style6
                    SetRect retBtn, lngBtnLeft - 1, -2, lngBtnWidth + 4, 12
                    DrawGradientInRectangle UserControl.hDC, m_lngColorTo, BlendColors(m_lngColorTo, BlendColors(m_lngColorFrom, vbBlack)), retBtn, VCylinderGradient, False, m_lngColorFrom, 1.01
                    SetRect retBtn, lngBtnLeft - 1, 8, lngBtnWidth + 4, m_lngHeightAux - 9
                    DrawGradientInRectangle UserControl.hDC, m_lngColorFrom, m_lngColorFrom, retBtn, VCylinderGradient, False, m_lngColorFrom, 7.01
                    APILineEx UserControl.hDC, lngBtnLeft - 1, 8, lngBtnLeft - 1 + lngBtnWidth + 2, 8, BlendColors(m_lngColorFrom, m_lngColorTo)
                    APILineEx UserControl.hDC, lngBtnLeft - 1, 0, lngBtnLeft - 1 + lngBtnWidth + 2, 0, BlendColors(vbWhite, m_lngColorTo)
                    lngColor = m_lngColorTo
            
            End Select
            
            If udtBtnState = STA_PRESSED Then
                
                lngBtnLeft = lngBtnLeft + 1
                lngBtnTop = lngBtnTop + 1
                SetRect retBtn, lngBtnLeft, lngBtnTop, lngBtnWidth, lngBtnHeight
                
                lngColor = m_lngColorFrom
                DrawGradientInRectangle UserControl.hDC, BlendColors(lngColor, vbWhite), vbWhite, retBtn, VCylinderGradient, False, lngBorderColor, 2.01
                
                APILineEx UserControl.hDC, lngBtnLeft + 1, lngBtnTop, lngBtnLeft + lngBtnWidth, lngBtnTop, lngBorderColor
                APILineEx UserControl.hDC, lngBtnLeft, lngBtnTop + 1, lngBtnLeft, lngBtnTop + lngBtnHeight, lngBorderColor
                APILineEx UserControl.hDC, lngBtnLeft + 1, lngBtnTop + lngBtnHeight, lngBtnLeft + lngBtnWidth, lngBtnTop + lngBtnHeight, lngBorderColor
                APILineEx UserControl.hDC, lngBtnLeft + lngBtnWidth, lngBtnTop + 1, lngBtnLeft + lngBtnWidth, lngBtnTop + lngBtnHeight, lngBorderColor
                
            End If
            
        Case STA_OVER
            
            SetRect retBtn, lngBtnLeft + 1, lngBtnTop, lngBtnWidth - 1, lngBtnHeight
            DrawGradientInRectangle UserControl.hDC, BlendColors(lngColor, vbWhite), vbWhite, retBtn, VCylinderGradient, False, lngBorderColor, 2.01
            
            APILineEx UserControl.hDC, lngBtnLeft + 1, lngBtnTop, lngBtnLeft + lngBtnWidth, lngBtnTop, lngBorderColor
            APILineEx UserControl.hDC, lngBtnLeft, lngBtnTop + 1, lngBtnLeft, lngBtnTop + lngBtnHeight, lngBorderColor
            APILineEx UserControl.hDC, lngBtnLeft + 1, lngBtnTop + lngBtnHeight, lngBtnLeft + lngBtnWidth, lngBtnTop + lngBtnHeight, lngBorderColor
            APILineEx UserControl.hDC, lngBtnLeft + lngBtnWidth, lngBtnTop + 1, lngBtnLeft + lngBtnWidth, lngBtnTop + lngBtnHeight, lngBorderColor
               
            If Index = 1 Then
                
                If m_intWindowState = vbMaximized Then
                    m_udtTitleBtn(Index).TooltipText = MnuSyst(jcRestore).Caption
                Else
                    m_udtTitleBtn(Index).TooltipText = MnuSyst(jcMaximize).Caption
                End If
                    
            End If
            
    End Select
    
    Select Case Index
        
        Case -1 'all buttons
            
            With m_udtTitleBtn(0)
                If .Visible = True Then
                    lngBtnLeft = m_udtTitleBtn(0).Left
                    APILineEx UserControl.hDC, lngBtnLeft + 6, lngBtnTop + 5, lngBtnLeft + 14, lngBtnTop + 13, lngColor
                    APILineEx UserControl.hDC, lngBtnLeft + 5, lngBtnTop + 5, lngBtnLeft + 14, lngBtnTop + 14, lngColor
                    APILineEx UserControl.hDC, lngBtnLeft + 5, lngBtnTop + 6, lngBtnLeft + 13, lngBtnTop + 14, lngColor
                    APILineEx UserControl.hDC, lngBtnLeft + 5, lngBtnTop + 12, lngBtnLeft + 13, lngBtnTop + 4, lngColor
                    APILineEx UserControl.hDC, lngBtnLeft + 5, lngBtnTop + 13, lngBtnLeft + 14, lngBtnTop + 4, lngColor
                    APILineEx UserControl.hDC, lngBtnLeft + 6, lngBtnTop + 13, lngBtnLeft + 14, lngBtnTop + 5, lngColor
                End If
            End With
            
            With m_udtTitleBtn(1)
                If .Visible = True Then
                    lngBtnLeft = .Left
                    If m_intWindowState = vbMaximized Then
                        If Ambient.UserMode Then        'If we're not in design mode
                            APILineEx UserControl.hDC, lngBtnLeft + 7, lngBtnTop + 5, lngBtnLeft + 7, lngBtnTop + 8, lngColor
                            APILineEx UserControl.hDC, lngBtnLeft + 7, lngBtnTop + 5, lngBtnLeft + 16, lngBtnTop + 5, lngColor
                            APILineEx UserControl.hDC, lngBtnLeft + 7, lngBtnTop + 4, lngBtnLeft + 16, lngBtnTop + 4, lngColor
                            APILineEx UserControl.hDC, lngBtnLeft + 15, lngBtnTop + 4, lngBtnLeft + 15, lngBtnTop + 11, lngColor
                            APILineEx UserControl.hDC, lngBtnLeft + 15, lngBtnTop + 11, lngBtnLeft + 11, lngBtnTop + 11, lngColor
                            APILineEx UserControl.hDC, lngBtnLeft + 3, lngBtnTop + 9, lngBtnLeft + 11, lngBtnTop + 9, lngColor
                            ApiRectangle UserControl.hDC, lngBtnLeft + 3, lngBtnTop + 8, 8, 7, lngColor
                        Else
                            APILineEx UserControl.hDC, lngBtnLeft + 4, lngBtnTop + 5, lngBtnLeft + 15, lngBtnTop + 5, lngColor
                            ApiRectangle UserControl.hDC, lngBtnLeft + 4, lngBtnTop + 6, 10, 8, lngColor
                        End If
                    Else
                        APILineEx UserControl.hDC, lngBtnLeft + 4, lngBtnTop + 5, lngBtnLeft + 15, lngBtnTop + 5, lngColor
                        ApiRectangle UserControl.hDC, lngBtnLeft + 4, lngBtnTop + 6, 10, 8, lngColor
                    End If
                End If
            End With
            
            With m_udtTitleBtn(2)
                If .Visible = True Then
                    lngBtnLeft = .Left
                    APILineEx UserControl.hDC, lngBtnLeft + 5, lngBtnTop + 11, lngBtnLeft + 14, lngBtnTop + 11, lngColor
                    APILineEx UserControl.hDC, lngBtnLeft + 5, lngBtnTop + 12, lngBtnLeft + 14, lngBtnTop + 12, lngColor
                    APILineEx UserControl.hDC, lngBtnLeft + 5, lngBtnTop + 13, lngBtnLeft + 14, lngBtnTop + 13, lngColor
                End If
            End With
            
        Case 0  'close button
    
            If m_udtTitleBtn(0).Visible = True Then
                With UserControl
                    APILineEx .hDC, lngBtnLeft + 6, lngBtnTop + 5, lngBtnLeft + 14, lngBtnTop + 13, lngColor
                    APILineEx .hDC, lngBtnLeft + 5, lngBtnTop + 5, lngBtnLeft + 14, lngBtnTop + 14, lngColor
                    APILineEx .hDC, lngBtnLeft + 5, lngBtnTop + 6, lngBtnLeft + 13, lngBtnTop + 14, lngColor
                    APILineEx .hDC, lngBtnLeft + 5, lngBtnTop + 12, lngBtnLeft + 13, lngBtnTop + 4, lngColor
                    APILineEx .hDC, lngBtnLeft + 5, lngBtnTop + 13, lngBtnLeft + 14, lngBtnTop + 4, lngColor
                    APILineEx .hDC, lngBtnLeft + 6, lngBtnTop + 13, lngBtnLeft + 14, lngBtnTop + 5, lngColor
                End With
            End If
            
        Case 1 'maximized button

            If m_intWindowState = vbMaximized Then
                If Ambient.UserMode Then        'If we're not in design mode
                    With UserControl
                        APILineEx .hDC, lngBtnLeft + 7, lngBtnTop + 5, lngBtnLeft + 7, lngBtnTop + 8, lngColor
                        APILineEx .hDC, lngBtnLeft + 7, lngBtnTop + 5, lngBtnLeft + 16, lngBtnTop + 5, lngColor
                        APILineEx .hDC, lngBtnLeft + 7, lngBtnTop + 4, lngBtnLeft + 16, lngBtnTop + 4, lngColor
                        APILineEx .hDC, lngBtnLeft + 15, lngBtnTop + 4, lngBtnLeft + 15, lngBtnTop + 11, lngColor
                        APILineEx .hDC, lngBtnLeft + 15, lngBtnTop + 11, lngBtnLeft + 11, lngBtnTop + 11, lngColor
                        APILineEx .hDC, lngBtnLeft + 3, lngBtnTop + 9, lngBtnLeft + 11, lngBtnTop + 9, lngColor
                        ApiRectangle .hDC, lngBtnLeft + 3, lngBtnTop + 8, 8, 7, lngColor
                    End With
                Else
                    APILineEx UserControl.hDC, lngBtnLeft + 4, lngBtnTop + 5, lngBtnLeft + 15, lngBtnTop + 5, lngColor
                    ApiRectangle UserControl.hDC, lngBtnLeft + 4, lngBtnTop + 6, 10, 8, lngColor
                End If
            Else
                APILineEx UserControl.hDC, lngBtnLeft + 4, lngBtnTop + 5, lngBtnLeft + 15, lngBtnTop + 5, lngColor
                ApiRectangle UserControl.hDC, lngBtnLeft + 4, lngBtnTop + 6, 10, 8, lngColor
            End If
    
        Case 2 'minimized button
            
            With UserControl
                APILineEx .hDC, lngBtnLeft + 5, lngBtnTop + 11, lngBtnLeft + 14, lngBtnTop + 11, lngColor
                APILineEx .hDC, lngBtnLeft + 5, lngBtnTop + 12, lngBtnLeft + 14, lngBtnTop + 12, lngColor
                APILineEx .hDC, lngBtnLeft + 5, lngBtnTop + 13, lngBtnLeft + 14, lngBtnTop + 13, lngColor
            End With
    
    End Select
    
    UserControl.Refresh


Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "DrawCaptionBtns"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Function TB_GetDesktopWorkArea(ByRef lngScreenLeft As Long, _
                                       ByRef lngScreenTop As Long, _
                                       ByRef lngScreenWidth As Long, _
                                       ByRef lngScreenHeight As Long) As Boolean
    
    '===========================================================
    ' Get the desktop work area using API SystemParametersInfo
    '===========================================================
    
    Dim retDesktopAreaRect      As RECT
    
    Call SystemParametersInfo(SPI_GETWORKAREA, 0, retDesktopAreaRect, 0)      'issue the API
    
    With retDesktopAreaRect
        lngScreenLeft = .Left * Screen.TwipsPerPixelX
        lngScreenTop = .Top * Screen.TwipsPerPixelY
        lngScreenWidth = (.Right - .Left) * Screen.TwipsPerPixelX
        lngScreenHeight = (.Bottom - .Top) * Screen.TwipsPerPixelY
    End With
    

Exit_Proc:
    On Error Resume Next
    Exit Function

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "TB_GetDesktopWorkArea"
    Err.Clear
    Resume Exit_Proc

End Function

Private Sub TB_Dock(frmMyForm As Form)

    Dim lngScreenLeft           As Long
    Dim lngScreenTop            As Long
    Dim lngScreenWidth          As Long
    Dim lngScreenHeight         As Long
    
    If Not frmMyForm.WindowState = vbMaximized Then
        TB_GetDesktopWorkArea lngScreenLeft, lngScreenTop, lngScreenWidth, lngScreenHeight
        frmMyForm.Move lngScreenLeft, lngScreenTop, lngScreenWidth, lngScreenHeight
        If m_blnLoaded Then
            m_picRight.Visible = True
            m_picBottom.Visible = True
        End If
        frmMyForm.Refresh
    End If

End Sub

Private Sub CaptioBtnSetup()
    
    '=======================================================
    ' Left, Width and Hegith values of each Caption buttons
    ' (minimize, maximize (restore) and close button)
    '
    'Captions for jcForms menu
    '=======================================================
    
    Dim intI                    As Integer
    Dim intJ                    As Integer
    
    On Error GoTo Err_Proc
    
    intJ = 0
    
    For intI = 0 To 2
        
        If m_udtTitleBtn(intI).Visible = True Then
            
            m_udtTitleBtn(intI).Left = UserControl.ScaleWidth - 3 - 21 * (intJ + 1)
            intJ = intJ + 1
            
            With m_udtTitleBtn(intI)
                .Width = 18
                .Height = 18
            End With
            
        End If
    
    Next intI
    
    MnuSyst(jcRestore).Caption = m_strMenuCaption(jcRestore)
    MnuSyst(jcMinimize).Caption = m_strMenuCaption(jcMinimize)
    MnuSyst(jcMaximize).Caption = m_strMenuCaption(jcMaximize)
    MnuSyst(jcClose).Caption = m_strMenuCaption(jcClose) & "            Alt+F4"
    MnuSyst(jcAlwaysOnTop).Caption = m_strMenuCaption(jcAlwaysOnTop)
    
    m_udtTitleBtn(2).TooltipText = MnuSyst(jcMinimize).Caption
    m_udtTitleBtn(0).TooltipText = m_strMenuCaption(jcClose)
    MnuSyst(jcMinimize).Visible = m_udtTitleBtn(2).Visible
    MnuSyst(jcMaximize).Visible = m_udtTitleBtn(1).Visible


Exit_Proc:
   On Error Resume Next
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "jcForms", "CaptionBtnSetup"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Function GetWhatButton(ByVal intX As Integer, _
                               ByVal intY As Integer) As Integer
    
    '===============================================
    'Identifies if we are over any caption button
    '===============================================
    
    Dim intI                As Integer
    
    GetWhatButton = -1
    
    For intI = 0 To 2
        
        If m_udtTitleBtn(intI).Visible = True Then
            If intX > m_udtTitleBtn(intI).Left Then
                If intX < m_udtTitleBtn(intI).Left + m_udtTitleBtn(intI).Width Then
                    If intY > m_udtTitleBtn(intI).Top Then
                        If intY < m_udtTitleBtn(intI).Top + m_udtTitleBtn(intI).Height Then
                            GetWhatButton = intI
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next

End Function

Private Sub PaintCaptionBtns(Index As Integer)
    
    Dim intI                As Integer
    
    On Error GoTo Err_Proc
    
    If Index = -1 Then
        DrawCaptionBtns STA_NORMAL, m_lngColorFrom, Index
        UserControl.MousePointer = 0
        m_intBtnIndex = -1
        m_intPrevBtnIndex = m_intBtnIndex

    Else
        
        For intI = 0 To 2
            
            If m_udtTitleBtn(intI).Visible = True Then
                If intI <> Index Then
                    DrawCaptionBtns STA_NORMAL, m_lngColorFrom, intI
                Else
                    DrawCaptionBtns STA_OVER, m_lngColorFrom, intI
                End If
            End If
        
        Next intI
    
    End If


Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "PaintCaptionBtns"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub CheckWindowState(Optional ByVal intOption As Integer = 0)
    
    '=========================
    'detect screen dimensions
    '=========================
    
    On Error GoTo Err_Proc
    
    If m_intWindowState = vbMinimized Then
        m_intPrevBtnIndex = -1
        m_frmPForm.WindowState = vbMinimized
        Exit Sub
    End If

    If m_intWindowState = vbMaximized Then
        
        If intOption = 0 Then
            m_intWindowState = vbNormal
            
            MnuSyst(jcRestore).Enabled = False
            MnuSyst(jcMaximize).Enabled = True
            
            If m_frmPForm.StartUpPosition = vbStartUpScreen Then
                m_frmPForm.Move (Screen.Width - m_retPrevSize.Right) / 2, (Screen.Height - m_retPrevSize.Bottom) / 2, m_retPrevSize.Right, m_retPrevSize.Bottom
            Else
                m_frmPForm.Move m_retPrevSize.Left, m_retPrevSize.Top, m_retPrevSize.Right, m_retPrevSize.Bottom
            End If
        Else
            MnuSyst(jcRestore).Enabled = True
            MnuSyst(jcMaximize).Enabled = False
    
            If m_blnLoaded Then
                m_picRight.Visible = False
                m_picBottom.Visible = False
            End If
    
            Call TB_Dock(m_frmPForm)
        End If
        
    Else
        
        If intOption = 0 Then
            m_intWindowState = vbMaximized
            
            MnuSyst(jcRestore).Enabled = True
            MnuSyst(jcMaximize).Enabled = False
    
            If m_blnLoaded Then
                m_picRight.Visible = False
                m_picBottom.Visible = False
            End If
    
            Call TB_Dock(m_frmPForm)
        Else
            MnuSyst(jcRestore).Enabled = False
            MnuSyst(jcMaximize).Enabled = True
            
            If m_frmPForm.StartUpPosition = vbStartUpScreen Then
                m_frmPForm.Move (Screen.Width - m_retPrevSize.Right) / 2, (Screen.Height - m_retPrevSize.Bottom) / 2, m_retPrevSize.Right, m_retPrevSize.Bottom
            Else
                m_frmPForm.Move m_retPrevSize.Left, m_retPrevSize.Top, m_retPrevSize.Right, m_retPrevSize.Bottom
            End If
        End If
        
    End If
    
    UserControl.Extender.Move 0, 0, UserControl.Parent.ScaleWidth, UserControl.Parent.ScaleHeight

Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    If Err.Number = 398 Then
    Else
        Err_Handler True, Err.Number, Err.Description, "jcForms", "jcWindowState"
    End If
    Err.Clear
    Resume Exit_Proc
    
End Sub

Private Sub ControlsChangeBackColor(lngBackColor As Long)
    
    Dim ctrlMyControl        As Control
    
    On Error Resume Next
    
    If m_blnChangeAllBackgrounds Then
        For Each ctrlMyControl In m_frmPForm
            
            If ctrlMyControl.BackColor <> &H80000005 Then
                If TypeOf ctrlMyControl Is CommandButton Then ctrlMyControl.Style = 1
                If ctrlMyControl.Name <> "m_picRight" And ctrlMyControl.Name <> "m_picBottom" Then
                    ctrlMyControl.BackColor = lngBackColor
                End If
            End If
        
        Next
    End If
    
End Sub

Private Sub ControlsChangeTop(Optional ByVal intOption As Integer = 1)
    
    Dim ctrlMyControl           As Control
    Dim intSign                 As Integer
    Dim intInitialTop           As Integer
    
    On Error Resume Next
    
    If intOption = 0 Then
        
        intInitialTop = FindMostTopDiffControl(m_lngHeightAux + 20)
        
        If m_intWindowState = vbMaximized Then
            m_retPrevSize.Bottom = m_retPrevSize.Bottom + intInitialTop * Screen.TwipsPerPixelY
        Else
            m_frmPForm.Height = m_frmPForm.Height + intInitialTop * Screen.TwipsPerPixelY
        End If
        
        For Each ctrlMyControl In m_frmPForm
            If TypeOf ctrlMyControl.Container Is jcForms Then
                ctrlMyControl.Top = ctrlMyControl.Top + intInitialTop * Screen.TwipsPerPixelY
            End If
        Next
    
    Else
        
        If m_udtCaptionStyle = Style2 And m_blnChangeTop = False Then
            
            intSign = 1
            
            If m_intWindowState = vbMaximized Then
                m_retPrevSize.Bottom = m_retPrevSize.Bottom + intSign * 10 * Screen.TwipsPerPixelY
            Else
                m_frmPForm.Height = m_frmPForm.Height + intSign * 10 * Screen.TwipsPerPixelY
            End If
            
            m_blnChangeTop = True
        
        ElseIf m_udtCaptionStyle <> Style2 And m_blnChangeTop = True Then
            
            intSign = -1
            
            If m_intWindowState = vbMaximized Then
                m_retPrevSize.Bottom = m_retPrevSize.Bottom + intSign * 10 * Screen.TwipsPerPixelY
            Else
                m_frmPForm.Height = m_frmPForm.Height + intSign * 10 * Screen.TwipsPerPixelY
            End If
            
            m_blnChangeTop = False
        
        Else
            intSign = 0
        End If
        
        For Each ctrlMyControl In m_frmPForm
            If TypeOf ctrlMyControl.Container Is jcForms Then
                If intSign <> 0 Then
                    ctrlMyControl.Top = ctrlMyControl.Top + intSign * 10 * Screen.TwipsPerPixelY
                End If
            End If
        Next
   
    End If

    On Error GoTo 0
    
End Sub

Private Function FindMostTopDiffControl(MinTop As Integer) As Integer
    
    Dim ctrlMyControl       As Control
    
    FindMostTopDiffControl = 0
    
    On Error Resume Next
    
    For Each ctrlMyControl In m_frmPForm
        If TypeOf ctrlMyControl.Container Is jcForms Then
            If ctrlMyControl.Top / Screen.TwipsPerPixelY < MinTop And (MinTop - ctrlMyControl.Top / Screen.TwipsPerPixelY) > FindMostTopDiffControl Then
                FindMostTopDiffControl = (MinTop - ctrlMyControl.Top / Screen.TwipsPerPixelY)
            End If
        End If
    Next
    
    If FindMostTopDiffControl = 0 Then
        FindMostTopDiffControl = 1000
        For Each ctrlMyControl In m_frmPForm
            If TypeOf ctrlMyControl.Container Is jcForms Then
                If (ctrlMyControl.Top / Screen.TwipsPerPixelY - MinTop) < FindMostTopDiffControl Then
                    FindMostTopDiffControl = ctrlMyControl.Top / Screen.TwipsPerPixelY - MinTop
                End If
            End If
        Next
        
        FindMostTopDiffControl = FindMostTopDiffControl * -1
    
    End If
    
    On Error GoTo 0

End Function

Public Sub Refresh(Optional ByVal blnForce As Boolean = False)
    
    On Error GoTo Err_Proc
    
    If m_blnChangeAllBackgrounds Or blnForce Then
        
        Select Case m_udtBackColorStyle
            
            Case Default
                UserControl.BackColor = TranslateColor(&H8000000F)
            
            Case Auto
                UserControl.BackColor = TranslateColor(m_lngColorTo)
            
            Case Custom
                UserControl.BackColor = TranslateColor(m_lngBackColor)
        
        End Select
    
    End If
    
    DrawTitleBarInitial
    UserControl_Resize

Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "Refresh"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Function GetInfo(ByVal lngInfo As Long) As String
    
    Dim strBuffer           As String
    Dim strRet              As String
    
    On Error GoTo Err_Proc
    
    strBuffer = String$(256, 0)
    strRet = GetLocaleInfo(LOCALE_USER_DEFAULT, lngInfo, strBuffer, Len(strBuffer))
    
    If strRet > 0 Then
        GetInfo = Left$(strBuffer, strRet - 1)
    Else
        GetInfo = ""
    End If


Exit_Proc:
    On Error Resume Next
    Exit Function

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "GetInfo"
    Err.Clear
    Resume Exit_Proc

End Function

Private Sub BorderStyleSetup()
    
    Dim intI As Integer
    
    On Error GoTo Err_Proc
    
    If Ambient.UserMode = False Then    'If we're in design mode
        
        With UserControl.Parent
            .BorderStyle = 0
            .ShowInTaskbar = True
            '.StartUpPosition = vbStartUpScreen
        End With
        
    End If
    
    If m_intWindowState = vbMaximized Then
        MnuSyst(jcRestore).Enabled = m_blnMaxButton
    End If
    
    MnuSyst(jcClose).Enabled = m_blnCloseButton
    m_udtTitleBtn(0).Visible = m_blnCloseButton
    m_udtTitleBtn(1).Visible = m_blnMaxButton
    m_udtTitleBtn(2).Visible = m_blnMinButton
    
    
    For intI = 0 To 2
        
        If m_udtTitleBtn(intI).Visible = True Then
            m_intLeftBtn = intI
        End If
    
    Next intI

    CaptioBtnSetup
    DrawTitleBarInitial
    If Ambient.UserMode Then DrawTitleBar
    
Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "BorderStyleSetup"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub TitleBarHeightSetup()
    
    On Error GoTo Err_Proc
    
    m_lngHeight = UserControl.TextHeight("H") + 10
    
    If m_lngHeight < m_intIconSize + 15 Then
        m_lngHeight = m_intIconSize + 15
    End If
    
    If m_udtCaptionStyle = Style2 Then
        m_lngHeightAux = m_lngHeight + 10
    Else
        m_lngHeightAux = m_lngHeight
    End If

    
Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "TitleBarHeightSetup"
    Err.Clear
    Resume Exit_Proc

End Sub

Public Sub FormMenuAdd(ByVal strMenuCaption As String, _
                       Optional ByVal blnAddSeparator As Boolean = False, _
                       Optional ByVal blnEnabled As Boolean = True, _
                       Optional ByVal blnChecked As Boolean = False, _
                       Optional ByVal blnVisible As Boolean = True)
    
    
    Dim intMenuID           As Integer
    
    On Error GoTo Err_Proc
    
Initial:
    
    intMenuID = MnuSyst.Count
    Load MnuSyst(intMenuID)
    MnuSyst(intMenuID).Visible = True
    MnuSyst(intMenuID).Enabled = True
    
    If blnAddSeparator Then           'Add a separator
        MnuSyst(intMenuID).Caption = "-"
        blnAddSeparator = False
        GoTo Initial
    Else
        
        With MnuSyst(intMenuID)
            .Caption = strMenuCaption
            .Enabled = blnEnabled
            .Checked = blnChecked
            .Visible = blnVisible
        End With
    End If

    
Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "FormMenuAdd"
    Err.Clear
    Resume Exit_Proc

End Sub

Public Sub ModifyAddedMenu(intMenuID As Integer, _
                           Optional strMenuCaption As String = "", _
                           Optional blnEnabled As Boolean = True, _
                           Optional blnChecked As Boolean = False, _
                           Optional blnVisible As Boolean = True)
    
    On Error GoTo Err_Proc
    
    If intMenuID > jcAlwaysOnTop Then
    
        With MnuSyst(intMenuID)
            If strMenuCaption <> "" Then
                .Caption = strMenuCaption
            End If
            .Enabled = blnEnabled
            .Checked = blnChecked
            .Visible = blnVisible
        End With
        
   End If
    
Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "ModifyAddedMenu"
    Err.Clear
    Resume Exit_Proc

End Sub

Public Sub FormMenuRemove(ByVal intMenuID As Integer)
    
    On Error GoTo Err_Proc
    
    If intMenuID > jcAlwaysOnTop Then
        If intMenuID <= MnuSyst.Count Then
            Unload MnuSyst(intMenuID)
        End If
    End If
    
    If intMenuID = MnuSyst.Count Then                   'last item was removed
        If MnuSyst(intMenuID - 1).Caption = "-" Then    'it's a seperator
            Unload MnuSyst(intMenuID - 1)
        End If
    End If
    
Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "FormMenuRemove"
    Err.Clear
    Resume Exit_Proc

End Sub

Public Function GetMenuItemValue(ByVal iMenuID As Integer, _
                                 ByVal PropertyID As jcMenuItemProp) As Variant
    
    '==========================================================
    ' Accessing all the properties of Menu items added by user
    '==========================================================
    
    On Error GoTo Err_Proc
    
    Select Case PropertyID
        
        Case jcCaption
            GetMenuItemValue = MnuSyst(iMenuID).Caption
        
        Case jcEnabled
            GetMenuItemValue = MnuSyst(iMenuID).Enabled
        
        Case jcChecked
            GetMenuItemValue = MnuSyst(iMenuID).Checked
        
        Case jcVisible
            GetMenuItemValue = MnuSyst(iMenuID).Visible
    
    End Select
    
    
Exit_Proc:
    On Error Resume Next
    Exit Function

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "GetMenuItemValue"
    Err.Clear
    Resume Exit_Proc

End Function

Public Sub SetjcFormsMenuCaption(ByVal BtnID As jcSystMenuItem, _
                                 ByVal strCaption As String)
    
    '====================================
    ' Change jcForms Menu Caption by user
    '====================================
    
    On Error GoTo Err_Proc
    
    If strCaption <> vbNullString Then
        Select Case BtnID
            
            Case jcRestore, jcMinimize, jcMaximize, jcClose, jcAlwaysOnTop
                
                m_strMenuCaption(BtnID) = strCaption
                MnuSyst(BtnID).Caption = m_strMenuCaption(BtnID)
                
                If BtnID = jcClose Then
                    MnuSyst(BtnID).Caption = m_strMenuCaption(BtnID) & "            Alt+F4"
                    m_udtTitleBtn(0).TooltipText = m_strMenuCaption(BtnID)
                ElseIf BtnID = jcMinimize Then
                    m_udtTitleBtn(2).TooltipText = m_strMenuCaption(BtnID)
                End If
        End Select
    End If
    
Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "SetjcFormsMenuCaption"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub SetupLanguageSystemMenu()
    
    '==============================================================
    ' sets captions for jcForms menu according to windows language
    ' Here you can add your language and captions for jcForms menu
    '
    ' Now in this version if language is different from spanish
    ' uc select english
    '==============================================================
        
    m_strLocalLanguage = GetInfo(LOCALE_SENGLANGUAGE)
    
    Select Case m_strLocalLanguage
        
        Case "Spanish"
            
            m_strMenuCaption(0) = "Restaurar"
            m_strMenuCaption(1) = "Minimizar"
            m_strMenuCaption(2) = "Maximizar"
            m_strMenuCaption(4) = "Cerrar"
            m_strMenuCaption(6) = "Siempre Visible"
            
        Case Else '"English"
            
            m_strMenuCaption(0) = "Restore"
            m_strMenuCaption(1) = "Minimize"
            m_strMenuCaption(2) = "Maximize"
            m_strMenuCaption(4) = "Close"
            m_strMenuCaption(6) = "Always on Top"
    
    End Select

End Sub

Public Sub SetAlwaysOnTop(blnValue As Boolean)
    
   '====================================
   ' Set or not Always on Top style
   '====================================
    
    If blnValue Then
        SetWindowPos m_frmPForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
    Else
        SetWindowPos m_frmPForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
    End If
    
    MnuSyst(jcAlwaysOnTop).Checked = blnValue

End Sub

Private Sub Err_Handler(Optional ByVal vblnDisplayError As Boolean = True, _
                        Optional ByVal vstrErrNumber As String = vbNullString, _
                        Optional ByVal vstrErrDescription As String = vbNullString, _
                        Optional ByVal vstrModuleName As String = vbNullString, _
                        Optional ByVal vstrProcName As String = vbNullString)
   
  Dim strTemp       As String
  Dim lngFN         As Long
   
   '====================================
   ' Purpose: Error handling - On Error
   '====================================
   
   'Show Error Message
   If vblnDisplayError Then
      strTemp = "Error occured: "
      If Len(vstrErrNumber) > 0 Then
         strTemp = strTemp & vstrErrNumber & vbNewLine
      Else
         strTemp = strTemp & vbNewLine
      End If
      If Len(vstrErrDescription) > 0 Then strTemp = strTemp & "Description: " & vstrErrDescription & vbNewLine
      If Len(vstrModuleName) > 0 Then strTemp = strTemp & "Module: " & vstrModuleName & vbNewLine
      If Len(vstrProcName) > 0 Then strTemp = strTemp & "Function: " & vstrProcName
      MsgBox strTemp, vbCritical, App.Title & " - ERROR"
   End If
   
   'Write error log
   lngFN = FreeFile
   Open App.Path & "\ErrorLog.txt" For Append As #lngFN
   Write #lngFN, Now, vstrErrNumber, vstrErrDescription, vstrModuleName, vstrProcName, _
         App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision, _
         Environ("username"), Environ("computername")
   Close #lngFN
   
End Sub

'=================================================
'  Paul Caton's subclassing subs and functions
'=================================================

Private Function IsFunctionExported(ByVal sFunction As String, _
                                    ByVal sModule As String) As Boolean
  
    '==============================================
    'Determine if the passed function is supported
    '==============================================
    
    Dim lnghMod               As Long
    Dim blnLibLoaded          As Boolean

    On Error GoTo Err_Proc
    
    lnghMod = GetModuleHandleA(sModule)

    If lnghMod = 0 Then
        lnghMod = LoadLibraryA(sModule)
        If lnghMod Then
            blnLibLoaded = True
        End If
    End If

    If lnghMod Then
        If GetProcAddress(lnghMod, sFunction) Then
            IsFunctionExported = True
        End If
    End If

    If blnLibLoaded Then
        FreeLibrary lnghMod
    End If


Exit_Proc:
    On Error Resume Next
    Exit Function

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "IsFunctionExported"
    Err.Clear
    Resume Exit_Proc

End Function

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
    
    '===============================================
    'Track the mouse leaving the indicated window
    '===============================================
    
    Dim tme         As TRACKMOUSEEVENT_STRUCT
  
    On Error GoTo Err_Proc
    
    If bTrack Then
        
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With

        If bTrackUser32 Then
            TrackMouseEvent tme
        Else
            TrackMouseEventComCtl tme
        End If
    End If


Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "TrackMouseLeave"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True) As Boolean
                    
                    
    '====================================================
    'uSelfSub code: Subclass the specified window handle
    '====================================================

    '*************************************************************************************************
    '* lng_hWnd   - Handle of the window to subclass
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
    '*************************************************************************************************
    
    Const CODE_LEN      As Long = 260                                        'Thunk length in bytes
    Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))         'Bytes to allocate per thunk, data + code + msg tables
    Const PAGE_RWX      As Long = &H40&                                      'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&                                    'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&                                    'Release allocated memory flag
    Const IDX_EBMODE    As Long = 3                                          'Thunk data index of the EbMode function address
    Const IDX_CWP       As Long = 4                                          'Thunk data index of the CallWindowProc function address
    Const IDX_SWL       As Long = 5                                          'Thunk data index of the SetWindowsLong function address
    Const IDX_FREE      As Long = 6                                          'Thunk data index of the VirtualFree function address
    Const IDX_BADPTR    As Long = 7                                          'Thunk data index of the IsBadCodePtr function address
    Const IDX_OWNER     As Long = 8                                          'Thunk data index of the Owner object's vTable address
    Const IDX_CALLBACK  As Long = 10                                         'Thunk data index of the callback method address
    Const IDX_EBX       As Long = 16                                         'Thunk code patch index of the thunk data
    Const SUB_NAME      As String = "sc_Subclass"                            'This routine's name
    
    Dim nAddr           As Long
    Dim nID             As Long
    Dim nMyID           As Long
  
    If IsWindow(lng_hWnd) = 0 Then                                           'Ensure the window handle is valid
        zError SUB_NAME, "Invalid window handle"
        Exit Function
    End If

    nMyID = GetCurrentProcessId                                              'Get this process's ID
    GetWindowThreadProcessId lng_hWnd, nID                                   'Get the process ID associated with the window handle
  
    If nID <> nMyID Then                                                     'Ensure that the window handle doesn't belong to another process
        zError SUB_NAME, "Window handle belongs to another process"
        Exit Function
    End If
  
    If oCallback Is Nothing Then                                             'If the user hasn't specified the callback owner
        Set oCallback = Me                                                   'Then it is me
    End If
  
    nAddr = zAddressOf(oCallback, nOrdinal)                                  'Get the address of the specified ordinal method
    If nAddr = 0 Then                                                        'Ensure that we've found the ordinal method
        zError SUB_NAME, "Callback method not found"
        Exit Function
    End If
    
    If z_Funk Is Nothing Then                                                'If this is the first time through, do the one-time initialization
        Set z_Funk = New Collection                                          'Create the hWnd/thunk-address collection
        z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
        z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&
    
        z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                 'Store CallWindowProc function address in the thunk data
        z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                  'Store the SetWindowLong function address in the thunk data
        z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                  'Store the VirtualFree function address in the thunk data
        z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")               'Store the IsBadCodePtr function address in the thunk data
    End If
  
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                 'Allocate executable memory

    If z_ScMem <> 0 Then                                                     'Ensure the allocation succeeded
        On Error GoTo CatchDoubleSub                                         'Catch double subclassing
        z_Funk.Add z_ScMem, "h" & lng_hWnd                                   'Add the hWnd/thunk-address to the collection
        On Error GoTo 0
  
        If bIdeSafety Then                                                   'If the user wants IDE protection
            z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                     'Store the EbMode function address in the thunk data
        End If
    
        z_Sc(IDX_EBX) = z_ScMem                                              'Patch the thunk data address
        z_Sc(IDX_HWND) = lng_hWnd                                            'Store the window handle in the thunk data
        z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                'Store the address of the before table in the thunk data
        z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)      'Store the address of the after table in the thunk data
        z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                  'Store the callback owner's object address in the thunk data
        z_Sc(IDX_CALLBACK) = nAddr                                           'Store the callback address in the thunk data
        z_Sc(IDX_PARM_USER) = lParamUser                                     'Store the lParamUser callback parameter in the thunk data
    
        nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
        
        If nAddr = 0 Then                                                    'Ensure the new WndProc was set correctly
            zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
            GoTo ReleaseMemory
        End If
        
        z_Sc(IDX_WNDPROC) = nAddr                                            'Store the original WndProc address in the thunk data
        RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                     'Copy the thunk code/data to the allocated memory
        sc_Subclass = True                                                   'Indicate success
  
    Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
  
    End If
  
    Exit Function                                                            'Exit sc_Subclass

CatchDoubleSub:
    zError SUB_NAME, "Window handle is already subclassed"
  
ReleaseMemory:
    VirtualFree z_ScMem, 0, MEM_RELEASE                                      'sc_Subclass has failed after memory allocation, so release the memory

End Function

Private Sub sc_Terminate()
  
    '==========================
    'Terminate all subclassing
    '==========================
    
    Dim i           As Long

    On Error GoTo Exit_Proc
    
    If Not (z_Funk Is Nothing) Then                                          'Ensure that subclassing has been started
        
        With z_Funk
            For i = .Count To 1 Step -1                                      'Loop through the collection of window handles in reverse order
                z_ScMem = .Item(i)                                           'Get the thunk address
                If IsBadCodePtr(z_ScMem) = 0 Then                            'Ensure that the thunk hasn't already released its memory
                    sc_UnSubclass zData(IDX_HWND)                            'UnSubclass
                End If
            Next i                                                           'Next member of the collection
        End With
        
        Set z_Funk = Nothing                                                 'Destroy the hWnd/thunk-address collection
    End If

Exit_Proc:
   On Error Resume Next
   
End Sub

Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  
    '================================================
    'UnSubclass the specified window handle
    '================================================
  
    If z_Funk Is Nothing Then                                               'Ensure that subclassing has been started
        zError "sc_UnSubclass", "Window handle isn't subclassed"
    
    Else
        If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                       'Ensure that the thunk hasn't already released its memory
            zData(IDX_SHUTDOWN) = -1                                        'Set the shutdown indicator
            zDelMsg ALL_MESSAGES, IDX_BTABLE                                'Delete all before messages
            zDelMsg ALL_MESSAGES, IDX_ATABLE                                'Delete all after messages
        End If
        
        z_Funk.Remove "h" & lng_hWnd                                        'Remove the specified window handle from the collection
    End If

End Sub

Private Sub sc_AddMsg(ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  
    '=====================================================================
    'Add the message value to the window handle's specified callback table
    '=====================================================================
    
    On Error GoTo Err_Proc
  
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                           'Ensure that the thunk hasn't already released its memory
        
        If When And MSG_BEFORE Then                                         'If the message is to be added to the before original WndProc table...
            zAddMsg uMsg, IDX_BTABLE                                        'Add the message to the before table
        End If
    
        If When And MSG_AFTER Then                                          'If message is to be added to the after original WndProc table...
            zAddMsg uMsg, IDX_ATABLE                                        'Add the message to the after table
        End If
    End If

Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "sc_AddMsg"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub sc_DelMsg(ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
                      
    '===========================================================================
    'Delete the message value from the window handle's specified callback table
    '===========================================================================

    On Error GoTo Err_Proc
    
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                    'Ensure that the thunk hasn't already released its memory
        
        If When And MSG_BEFORE Then                                  'If the message is to be deleted from the before original WndProc table...
            zDelMsg uMsg, IDX_BTABLE                                 'Delete the message from the before table
        End If
    
        If When And MSG_AFTER Then                                   'If the message is to be deleted from the after original WndProc table...
            zDelMsg uMsg, IDX_ATABLE                                 'Delete the message from the after table
        End If
  
    End If

Exit_Proc:
    On Error Resume Next
    Exit Sub

Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "jcForms", "sc_DelMsg"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal wParam As Long, _
                                    ByVal lParam As Long) As Long
                                    
    '==========================
    'Call the original WndProc
    '==========================
  
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
        sc_CallOrigWndProc = CallWindowProcA(zData(IDX_WNDPROC), _
                                             lng_hWnd, uMsg, wParam, lParam)  'Call the original WndProc of the passed window handle parameter
    End If

End Function

Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long
    
    '==================================================
    'Get the subclasser lParamUser callback parameter
    '==================================================
    
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
        sc_lParamUser = zData(IDX_PARM_USER)                                  'Get the lParamUser callback parameter
    End If

End Property

Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, _
                                   ByVal NewValue As Long)
    
    '==================================================
    'Let the subclasser lParamUser callback parameter
    '==================================================
  
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                         'Ensure that the thunk hasn't already released its memory
        zData(IDX_PARM_USER) = NewValue                                   'Set the lParamUser callback parameter
    End If

End Property

'=====================================================================
' The following routines are exclusively for the sc_ subclass routines
'=====================================================================

Private Sub zAddMsg(ByVal uMsg As Long, _
                    ByVal nTable As Long)
  
    '============================================================
    'Add the message to the specified table of the window handle
    '============================================================
  
    Dim nCount      As Long                                           'Table entry count
    Dim nBase       As Long                                           'Remember z_ScMem
    Dim i           As Long                                           'Loop index

    nBase = z_ScMem                                                   'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                           'Map zData() to the specified table

    If uMsg = ALL_MESSAGES Then                                       'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                         'Set the table entry count to ALL_MESSAGES
    Else
    
        nCount = zData(0)                                             'Get the current table entry count
        
        If nCount >= MSG_ENTRIES Then                                 'Check for message table overflow
            zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
            GoTo Bail
        End If

        For i = 1 To nCount                                           'Loop through the table entries
            If zData(i) = 0 Then                                      'If the element is free...
                zData(i) = uMsg                                       'Use this element
                GoTo Bail                                             'Bail
            ElseIf zData(i) = uMsg Then                               'If the message is already in the table...
                GoTo Bail                                             'Bail
            End If
        Next i                                                        'Next message table entry

        nCount = i                                                    'On drop through: i = nCount + 1, the new table entry count
        zData(nCount) = uMsg                                          'Store the message in the appended table entry
     End If

    zData(0) = nCount                                                 'Store the new table entry count

Bail:
    z_ScMem = nBase                                                   'Restore the value of z_ScMem

End Sub

Private Sub zDelMsg(ByVal uMsg As Long, _
                    ByVal nTable As Long)
  
    '=================================================================
    'Delete the message from the specified table of the window handle
    '=================================================================
  
    Dim nCount      As Long                                           'Table entry count
    Dim nBase       As Long                                           'Remember z_ScMem
    Dim i           As Long                                           'Loop index

    nBase = z_ScMem                                                   'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                           'Map zData() to the specified table

    If uMsg = ALL_MESSAGES Then                                       'If ALL_MESSAGES are being deleted from the table...
        zData(0) = 0                                                  'Zero the table entry count
    Else
    
        nCount = zData(0)                                             'Get the table entry count
    
        For i = 1 To nCount                                           'Loop through the table entries
      
            If zData(i) = uMsg Then                                   'If the message is found...
                zData(i) = 0                                          'Null the msg value -- also frees the element for re-use
                GoTo Bail                                             'Bail
            End If
        
        Next i                                                        'Next message table entry
    
        zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
  End If
  
Bail:
    z_ScMem = nBase                                                   'Restore the value of z_ScMem

End Sub

Private Sub zError(ByVal sRoutine As String, _
                   ByVal sMsg As String)
  
    '==============
    'Error handler
    '==============
  
    App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
    
    MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine

End Sub

Private Function zFnAddr(ByVal sDLL As String, _
                         ByVal sProc As String) As Long
    
    '===================================================
    'Return the address of the specified DLL/procedure
    '===================================================
    
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)       'Get the specified procedure address
    Debug.Assert zFnAddr                                          'In the IDE, validate that the procedure address was located
    
End Function

Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long

    '=================================================================
    'Map zData() to the thunk address for the specified window handle
    '=================================================================
    
    If z_Funk Is Nothing Then                                     'Ensure that subclassing has been started
        zError "zMap_hWnd", "Subclassing hasn't been started"
    Else
    
        On Error GoTo Catch                                       'Catch unsubclassed window handles
        
        z_ScMem = z_Funk("h" & lng_hWnd)                          'Get the thunk address
        zMap_hWnd = z_ScMem
    
    End If
  
    Exit Function                                                 'Exit returning the thunk address

Catch:
    zError "zMap_hWnd", "Window handle isn't subclassed"

End Function

Private Function zAddressOf(ByVal oCallback As Object, _
                            ByVal nOrdinal As Long) As Long

    '==========================================================================
    'Return the address of the specified ordinal method on the oCallback object
    '1 = last private method
    '2 = second last private method, etc
    '==========================================================================
    
    Dim bSub        As Byte                                           'Value we expect to find pointed at by a vTable method entry
    Dim bVal        As Byte
    Dim nAddr       As Long                                           'Address of the vTable
    Dim i           As Long                                           'Loop index
    Dim j           As Long                                           'Loop limit
  
    RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                 'Get the address of the callback object's instance
  
    If Not zProbe(nAddr + &H1C, i, bSub) Then                         'Probe for a Class method
        If Not zProbe(nAddr + &H6F8, i, bSub) Then                    'Probe for a Form method
            If Not zProbe(nAddr + &H7A4, i, bSub) Then                'Probe for a UserControl method
                Exit Function                                         'Bail...
            End If
        End If
    End If
  
    i = i + 4                                                         'Bump to the next entry
    j = i + 1024                                                      'Set a reasonable limit, scan 256 vTable entries
  
    Do While i < j
        RtlMoveMemory VarPtr(nAddr), i, 4                             'Get the address stored in this vTable entry
    
        If IsBadCodePtr(nAddr) Then                                   'Is the entry an invalid code address?
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4   'Return the specified vTable entry address
            Exit Do                                                   'Bad method signature, quit loop
        End If

        RtlMoveMemory VarPtr(bVal), nAddr, 1                          'Get the byte pointed to by the vTable entry
        
        If bVal <> bSub Then                                          'If the byte doesn't match the expected value...
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4   'Return the specified vTable entry address
            Exit Do                                                   'Bad method signature, quit loop
        End If
    
        i = i + 4                                                     'Next vTable entry
    Loop

End Function

Private Function zProbe(ByVal nStart As Long, _
                        ByRef nMethod As Long, _
                        ByRef bSub As Byte) As Boolean

    '===========================================================
    'Probe at the specified start address for a method signature
    '===========================================================
    
    Dim bVal            As Byte
    Dim nAddr           As Long
    Dim nLimit          As Long
    Dim nEntry          As Long
    
    nAddr = nStart                                                'Start address
    nLimit = nAddr + 32                                           'Probe eight entries
  
    Do While nAddr < nLimit                                       'While we've not reached our probe depth
        RtlMoveMemory VarPtr(nEntry), nAddr, 4                    'Get the vTable entry
    
        If nEntry <> 0 Then                                       'If not an implemented interface
            RtlMoveMemory VarPtr(bVal), nEntry, 1                 'Get the value pointed at by the vTable entry
            
            If bVal = &H33 Or bVal = &HE9 Then                    'Check for a native or pcode method signature
                nMethod = nAddr                                   'Store the vTable entry
                bSub = bVal                                       'Store the found method signature
                zProbe = True                                     'Indicate success
                Exit Function                                     'Return
            End If
        End If
    
        nAddr = nAddr + 4                                         'Next vTable entry
    Loop

End Function

Private Property Get zData(ByVal nIndex As Long) As Long
    
    RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4

End Property

Private Property Let zData(ByVal nIndex As Long, _
                           ByVal nValue As Long)
    
    RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4

End Property

Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)

    '=============================================================================
    'Subclass callback: must be private and the last method in the source file
    '=============================================================================

    '*************************************************************************************************
    '* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
    '*              you will know unless the callback for the uMsg value is specified as
    '*              MSG_BEFORE_AFTER (both before and after the original WndProc).
    '* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
    '*              message being passed to the original WndProc and (if set to do so) the after
    '*              original WndProc callback.
    '* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
    '*              and/or, in an after the original WndProc callback, act on the return value as set
    '*              by the original WndProc.
    '* lng_hWnd   - Window handle.
    '* uMsg       - Message value.
    '* wParam     - Message related data.
    '* lParam     - Message related data.
    '* lParamUser - User-defined callback parameter
    '*************************************************************************************************
    
    Select Case uMsg
        
        Case WM_MOUSEMOVE
            
            If Not bInCtrl Then
                bInCtrl = True
                Call TrackMouseLeave(lng_hWnd)
            End If
            
        Case WM_MOUSELEAVE
            
            bInCtrl = False
            UserControl.Cls
            DrawTitleBar
            
        Case WM_SETCURSOR
            
            UserControl.Extender.Move 0, 0, UserControl.Parent.ScaleWidth, UserControl.Parent.ScaleHeight
            If m_intWindowState = 0 Then SetRect m_retPrevSize, m_frmPForm.Left, m_frmPForm.Top, m_frmPForm.Width, m_frmPForm.Height
        
        Case WM_SYSCOMMAND
            
            If wParam = 61536 Then Call UserControl_KeyDown(vbKeyF4, vbAltMask)

        Case WM_MOVE
            
            If m_intWindowState = vbMaximized Then
                
                If m_blnLoaded = False Then
                    
                    If m_blnMaxButton Then
                        MnuSyst(jcRestore).Enabled = True
                        MnuSyst(jcMaximize).Enabled = False
                        
                        Call TB_Dock(m_frmPForm)
                    
                        UserControl.Extender.Move 0, 0, UserControl.Parent.ScaleWidth, UserControl.Parent.ScaleHeight
                    End If
                
                End If
            
            End If
       
       Case WM_ACTIVATE
        
            If m_blnFormLoaded = True Then
                
                Select Case wParam
                    Case 0
                        m_blnFormActivate = False
                    Case 1, 2
                        m_blnFormActivate = True
                End Select
                
                Call SetThemeColor

                Refresh
            
            End If
            
        Case WM_SYSCOLORCHANGE
                    
            Call SetThemeColor

            Refresh True

    End Select
    
End Sub

