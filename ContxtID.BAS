Attribute VB_Name = "Module1"
Option Explicit
'=====================================================================
'=====================================================================
'
'This source code contains the following routines:
'  o SetAppHelp() 'Called in the main Form_Load event to register your
'                 'program with WINHELP.EXE
'  o QuitHelp()    'Deregisters your program with WINHELP.EXE. Should
'                  'be called in your main Form_Unload event
'  o ShowHelpTopic(Topicnum) 'Brings up context sensitive help based on
'                  'any of the following CONTEXT IDs
'  o ShowContents  'Displays the startup topic
'  o HelpWindowSize(x,y,dx,dy) ' Position help window in a screen
'                              ' independent manner
'  o SearchHelp()  'Brings up the windows help KEYWORD SEARCH dialog box
'***********************************************************************
'
'=====================================================================
'List of Context IDs for <jspad>
'=====================================================================
'
'
'  Help engine section.

' Commands to pass WinHelp()
'Global Const HELP_CONTEXT = &H1 '  Display topic in ulTopic
Private Const HELP_QUIT = &H2    '  Terminate help
'Global Const HELP_FINDER = &HB  '  Display Contents tab
'Global Const HELP_INDEX = &H3   '  Display index
'Global Const HELP_HELPONHELP = &H4      '  Display help on using help
'Global Const HELP_SETINDEX = &H5        '  Set the current Index for multi index help
'Global Const HELP_KEY = &H101           '  Display topic for keyword in offabData
'Global Const HELP_MULTIKEY = &H201
Private Const HELP_CONTENTS = &H3     ' Display Help for a particular topic
'Global Const HELP_SETCONTENTS = &H5  ' Display Help contents topic
'Global Const HELP_CONTEXTPOPUP = &H8 ' Display Help topic in popup window
'Global Const HELP_FORCEFILE = &H9    ' Ensure correct Help file is displayed
'Global Const HELP_COMMAND = &H102    ' Execute Help macro
Private Const HELP_PARTIALKEY = &H105 ' Display topic found in keyword list
'Global Const HELP_SETWINPOS = &H203  ' Display and position Help window

    Type HELPWININFO
      wStructSize As Long
      x As Long
      y As Long
      dx As Long
      dy As Long
      wMax As Long
      rgChMember As String * 2
    End Type
    Declare Function WinHelp Lib "User32.dll" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
    'Declare Function WinHelpByInfo Lib "User32.dll" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, dwData As HELPWININFO) As Long
    'Declare Function WinHelpByStr Lib "User32.dll" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData$) As Long
    Private Declare Function WinHelpByNum Lib "User32.dll" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData&) As Long
    Dim m_hWndMainWindow As Long ' hWnd to tell WINHELP the helpfile owner

'Dim MainWindowInfo As HELPWININFO
Public Sub SetAppHelp(ByVal hWndMainWindow)
'=====================================================================
'To use these subroutines to access WINHELP, you need to add
'at least this one subroutine call to your code
'     o  In the Form_Load event of your main Form enter:
'        Call SetAppHelp(Me.hWnd) 'To setup helpfile variables
'         (If you are not interested in keyword searching or context
'         sensitive help, this is the only call you need to make!)
'=====================================================================
    m_hWndMainWindow = hWndMainWindow
    If VBA.Right$(Trim$(App.path), 1) = "\" Then
        App.HelpFile = App.path + "help\jsplus.hlp"
    Else
        App.HelpFile = App.path + "\help\jsplus.hlp"
    End If
   ' MainWindowInfo.wStructSize = 26
   ' MainWindowInfo.x = 256
   ' MainWindowInfo.y = 256
   ' MainWindowInfo.dx = 512
   ' MainWindowInfo.dy = 512
   ' MainWindowInfo.rgChMember = Chr$(0) + Chr$(0)
End Sub
Public Sub QuitHelp()
    Dim Result As Variant
    Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_QUIT, Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0))
End Sub
Public Sub ShowHelpContents()
'=====================================================================
'  DISPLAY STARTUP TOPIC IN RESPONSE TO A COMMAND BUTTON or MENU ...
'=====================================================================
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, App.HelpFile, HELP_CONTENTS, CLng(0))

End Sub
Public Sub SearchHelp()
'=====================================================================
'  TO ADD KEYWORD SEARCH CAPABILITY...
'=====================================================================
'     o   In your Help|Search menu selection, simply enter:
'         Call SearchHelp() 'To invoke helpfile keyword search dialog
'
    Dim Result As Variant

    Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_PARTIALKEY, ByVal "")

End Sub

