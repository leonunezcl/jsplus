VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'ToolTip Class - allow advance tool tip functions like Multiline, and "Cartoon-Bubble" style.


'Windows API Functions
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'Windows API Constants
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1

Private Const FORE_COLOUR As Long = "&H80000017"
Private Const BACK_COLOUR As Long = "&H80000018"

'Windows API Types
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'Tooltip Window Constants
Private Const TTS_NOPREFIX = &H2
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

'Tooltip Window Types
Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    lHwnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type

'ToolTip Style
Public Enum ttStyleEnum
    TTStandard
    TTBalloon
End Enum

'Local Member Varibales
Private mstrTitle As String
Private mlngForeColor As OLE_COLOR
Private mlngBackColor As OLE_COLOR
Private mlngHwndParentControl As Long
Private mblnCentered As Boolean
Private mlngStyle As ttStyleEnum
Private mstrText As String
Private mblnMultiLine As Boolean    'Multiline Tool Tips??

'private data
Private lngHwnd As Long
Private mtypToolInfo As TOOLINFO

'Init the class
Private Sub Class_Initialize()
    'Set some defaults
    Me.MultiLine = True
    Me.Style = TTStandard
    Me.Centered = False
    Me.HwndParentControl = 0
    Me.BackColor = BACK_COLOUR  'Default Back colour
    Me.ForeColor = FORE_COLOUR  'Default Fore colour
End Sub

'Terminate the class
Private Sub Class_Terminate()
    If lngHwnd <> 0 Then
        DestroyWindow lngHwnd
    End If
End Sub


'Create the tool tip
Public Function Create() As Boolean

    On Error GoTo CreateError

    Dim lpRect As RECT
    Dim lWinStyle As Long
    
    If lngHwnd <> 0 Then
        DestroyWindow lngHwnd
    End If
    
    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    
    'create baloon style if desired
    If Me.Style = TTBalloon Then
        lWinStyle = lWinStyle Or TTS_BALLOON
    End If
    
    'the parent control has to have been set first
    If Me.HwndParentControl <> 0 Then
        lngHwnd = CreateWindowEx(0&, _
                    TOOLTIPS_CLASSA, _
                    vbNullString, _
                    lWinStyle, _
                    CW_USEDEFAULT, _
                    CW_USEDEFAULT, _
                    CW_USEDEFAULT, _
                    CW_USEDEFAULT, _
                    Me.HwndParentControl, _
                    0&, _
                    App.hInstance, _
                    0&)
                    
        'make our tooltip window a "topmost" window
        SetWindowPos lngHwnd, _
            HWND_TOPMOST, _
            0&, _
            0&, _
            0&, _
            0&, _
            SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
                    
        'get the rect of the parent control
        GetClientRect Me.HwndParentControl, lpRect
        
        'now set our tooltip info structure
        With mtypToolInfo
            'if we want it centered, then set that flag
            If Me.Centered = True Then
                .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
            Else
                .lFlags = TTF_SUBCLASS
            End If
            
            'set the hwnd prop to our parent control's hwnd
            .lHwnd = Me.HwndParentControl
            .lId = 0
            .hInstance = App.hInstance
            .lpStr = Me.Text ' ALREADY SET
            .lpRect = lpRect
        End With
        
        'add the tooltip structure
        SendMessage lngHwnd, TTM_ADDTOOLA, 0&, mtypToolInfo
        
        'if we want a title or we want an icon
        If Title <> vbNullString Then
            SendMessage lngHwnd, TTM_SETTITLE, 0, ByVal Title
        End If
        
        'Goes all black if you set it to the standard colours
        If ForeColor <> FORE_COLOUR Then
            SendMessage lngHwnd, TTM_SETTIPTEXTCOLOR, ForeColor, 0&    'Set the ForeColor
        End If
        If BackColor <> BACK_COLOUR Then
            SendMessage lngHwnd, TTM_SETTIPBKCOLOR, BackColor, 0&      'Set the BackColor
        End If
        If MultiLine = True Then
            SendMessage lngHwnd, TTM_SETMAXTIPWIDTH, 0&, 0  'Set to multiline
        End If
    End If
    
    Create = True   'All went well!

CreateExit:
    On Error Resume Next
    Exit Function

CreateError:
    Create = False 'Failed!
    Resume CreateExit
End Function


'Set the control you want the tool tip to apply to
Public Property Let HwndParentControl(ByVal lHwnd As Long)
    mlngHwndParentControl = lHwnd
End Property
Public Property Get HwndParentControl() As Long
    HwndParentControl = mlngHwndParentControl
End Property

'Style of the ToolTip
Public Property Let Style(ByVal lngToolTipStyle As ttStyleEnum)
    mlngStyle = lngToolTipStyle
End Property
Public Property Get Style() As ttStyleEnum
    Style = mlngStyle
End Property

'Want the tool tip Centered? (works well with Baloon type tips!)
Public Property Let Centered(ByVal blnCentered As Boolean)
    mblnCentered = blnCentered
End Property
Public Property Get Centered() As Boolean
    Centered = mblnCentered
End Property

'ToolTip ForeColour
Public Property Let ForeColor(ByVal lngForeColor As OLE_COLOR)
    mlngForeColor = lngForeColor
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mlngForeColor
End Property

'Tool Tip Background Colour
Public Property Let BackColor(ByVal lngBackColor As OLE_COLOR)
    mlngBackColor = lngBackColor
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = mlngBackColor
End Property

'Tool Tip Title
Public Property Let Title(ByVal vData As String)
    mstrTitle = vData
End Property
Public Property Get Title() As String
    Title = mstrTitle
End Property

'The Actual Tool Tip Text
Public Property Let Text(ByVal strText As String)
    mstrText = strText
End Property
Public Property Get Text() As String
    Text = mstrText
End Property

'Want the Tool tip to be able to show multi line text
Public Property Get MultiLine() As Boolean
    MultiLine = mblnMultiLine
End Property
Public Property Let MultiLine(blnMultiLine As Boolean)
    mblnMultiLine = blnMultiLine
End Property

Public Property Get SystemToolTipForeColor() As OLE_COLOR
    SystemToolTipForeColor = FORE_COLOUR
End Property
Public Property Get SystemToolTipBackColor() As OLE_COLOR
    SystemToolTipBackColor = BACK_COLOUR
End Property






