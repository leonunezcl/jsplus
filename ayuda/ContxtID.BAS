Attribute VB_Name = "ContextIDs"
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
'List of Context IDs for <jsplus>
'=====================================================================
Global Const Hlp_Index = 10    'Main Help Window
Global Const Hlp_File_Menu = 20    'Main Help Window
Global Const Hlp_Edit_Menu = 30    'Main Help Window
Global Const Hlp_View_Menu = 40    'Main Help Window
Global Const Hlp_Search_Menu = 50    'Main Help Window
Global Const Hlp_Format_Menu = 60    'Main Help Window
Global Const Hlp_Insert_Menu = 70    'Main Help Window
Global Const Hlp_Macro_Menu = 80    'Main Help Window
Global Const Hlp_Functions_Menu = 90    'Main Help Window
Global Const Hlp_Javascript_Menu = 100    'Main Help Window
Global Const Hlp_Plus_Menu = 110    'Main Help Window
Global Const Hlp_Html_Menu = 120    'Main Help Window
Global Const Hlp_What_itxs = 130    'Main Help Window
Global Const Hlp_InstallingxUninstalling_Javascript = 140    'Main Help Window
Global Const Hlp_Feedback_and = 150    'Main Help Window
Global Const Hlp_How_to = 160    'Main Help Window
Global Const Hlp_Javascript_Plus = 170    'Main Help Window
Global Const Hlp_Whatxs_New = 180    'Main Help Window
Global Const Hlp_Utilities_Menu = 190    'Main Help Window
Global Const Hlp_Shell_Menu = 200    'Main Help Window
Global Const Hlp_Syntax_Highlighting = 220    'Main Help Window
Global Const Hlp_Intellisense = 230    'Main Help Window
Global Const Hlp_Javascript_Plusx = 240    'Main Help Window
Global Const Hlp_Colored_Scrollbar = 250    'Main Help Window
Global Const Hlp_Drop_Down = 260    'Main Help Window
Global Const Hlp_Email_Link = 270    'Main Help Window
Global Const Hlp_IFrame_Wizard = 280    'Main Help Window
Global Const Hlp_Image_Rollover = 290    'Main Help Window
Global Const Hlp_Last_Modified = 300    'Main Help Window
Global Const Hlp_Left_Menu = 310    'Main Help Window
Global Const Hlp_Meta_Tag = 320    'Main Help Window
Global Const Hlp_MouseOver_Text = 330    'Main Help Window
Global Const Hlp_Popup_Window = 340    'Main Help Window
Global Const Hlp_Regular_expression = 350    'Main Help Window
Global Const Hlp_Javascript_Plus1 = 360    'Main Help Window
Global Const Hlp_Array_Wizards = 370    'Main Help Window
Global Const Hlp_Object_Browser = 380    'Main Help Window
Global Const Hlp_Variable_wizard = 390    'Main Help Window
Global Const Hlp_Event_wizard = 400    'Main Help Window
Global Const Hlp_Html_Wizards = 410    'Main Help Window
Global Const Hlp_Form_Wizard = 420    'Main Help Window
Global Const Hlp_CheckBox_Wizard = 430    'Main Help Window
Global Const Hlp_Radio_Wizard = 440    'Main Help Window
Global Const Hlp_Listbox_Wizard = 460    'Main Help Window
Global Const Hlp_ComboBox_Wizard = 470    'Main Help Window
Global Const Hlp_TextBox_Wizard = 480    'Main Help Window
Global Const Hlp_TextArea_Wizard = 490    'Main Help Window
Global Const Hlp_Button_Wizard = 500    'Main Help Window
Global Const Hlp_Hidden_Entry = 510    'Main Help Window
Global Const Hlp_File_Attach = 520    'Main Help Window
Global Const Hlp_Image_Wizard = 530    'Main Help Window
Global Const Hlp_Table_Wizard = 540    'Main Help Window
Global Const Hlp_Other_features = 550    'Main Help Window
Global Const Hlp_Settings_Menu = 560    'Main Help Window
Global Const Hlp_About_VBSoftware = 570    'Main Help Window
Global Const Hlp_About_the = 580    'Main Help Window
'=====================================================================
'
'
'  Help engine section.

' Commands to pass WinHelp()
Global Const HELP_CONTEXT = &H1 '  Display topic in ulTopic
Global Const HELP_QUIT = &H2    '  Terminate help
Global Const HELP_FINDER = &HB  '  Display Contents tab
Global Const HELP_INDEX = &H3   '  Display index
Global Const HELP_HELPONHELP = &H4      '  Display help on using help
Global Const HELP_SETINDEX = &H5        '  Set the current Index for multi index help
Global Const HELP_KEY = &H101           '  Display topic for keyword in offabData
Global Const HELP_MULTIKEY = &H201
Global Const HELP_CONTENTS = &H3     ' Display Help for a particular topic
Global Const HELP_SETCONTENTS = &H5  ' Display Help contents topic
Global Const HELP_CONTEXTPOPUP = &H8 ' Display Help topic in popup window
Global Const HELP_FORCEFILE = &H9    ' Ensure correct Help file is displayed
Global Const HELP_COMMAND = &H102    ' Execute Help macro
Global Const HELP_PARTIALKEY = &H105 ' Display topic found in keyword list
Global Const HELP_SETWINPOS = &H203  ' Display and position Help window

    Type HELPWININFO
      wStructSize As Long
      X As Long
      Y As Long
      dX As Long
      dY As Long
      wMax As Long
      rgChMember As String * 2
    End Type
    Declare Function WinHelp Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
    Declare Function WinHelpByInfo Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, dwData As HELPWININFO) As Long
    Declare Function WinHelpByStr Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData$) As Long
    Declare Function WinHelpByNum Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData&) As Long
    Dim m_hWndMainWindow as Long ' hWnd to tell WINHELP the helpfile owner

Dim MainWindowInfo as HELPWININFO
Public Sub SetAppHelp (ByVal hWndMainWindow)
'=====================================================================
'To use these subroutines to access WINHELP, you need to add
'at least this one subroutine call to your code
'     o  In the Form_Load event of your main Form enter:
'        Call SetAppHelp(Me.hWnd) 'To setup helpfile variables
'         (If you are not interested in keyword searching or context
'         sensitive help, this is the only call you need to make!)
'=====================================================================
    m_hWndMainWindow = hWndMainWindow
    If Right$(Trim$(App.Path),1) = "\" then
        App.HelpFile = App.Path + "jsplus.HLP"
    else
        App.HelpFile = App.Path + "\jsplus.HLP"
    end if
    MainWindowInfo.wStructSize = 26
    MainWindowInfo.X=256
    MainWindowInfo.Y=256
    MainWindowInfo.dX=512
    MainWindowInfo.dY=512
    MainWindowInfo.rgChMember=Chr$(0)+Chr$(0)
End Sub
Public Sub QuitHelp ()
    Dim Result as Variant
    Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_QUIT, Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0))
End Sub
Public Sub ShowHelpTopic (ByVal ContextID As Long)
'=====================================================================
'  FOR CONTEXT SENSITIVE HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic(<any Hlpxxx entry above>)
'=====================================================================
'  TO ADD FORM LEVEL CONTEXT SENSITIVE HELP...
'=====================================================================
'     o  For FORM level context sensetive help, you should set each 
'        Me.HelpContext=<any Hlp_xxx entry above>
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXT, Clng(ContextID))

End Sub
Public Sub ShowHelpTopic2 (ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 2 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic2(<any Hlpxxx entry above>)
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile &">HlpWnd02", HELP_CONTEXT, Clng(ContextID))

End Sub
Public Sub ShowHelpTopic3 (ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 3 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic3(<any Hlpxxx entry above>)
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile &">HlpWnd03", HELP_CONTEXT, Clng(ContextID))

End Sub
Public Sub ShowGlossary ()
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXT, Clng(64000))

End Sub
Public Sub ShowPopupHelp (ByVal ContextID As Long)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXTPOPUP, Clng(ContextID))

End Sub
Public Sub DoHelpMacro (ByVal MacroString As String)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result as Variant

    Result = WinHelpByStr(m_hWndMainWindow, APP.HelpFile, HELP_COMMAND, ByVal(Macrostring))

End Sub
Public Sub ShowHelpContents ()
'=====================================================================
'  DISPLAY STARTUP TOPIC IN RESPONSE TO A COMMAND BUTTON or MENU ...
'=====================================================================
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTENTS, Clng(0))

End Sub
Public Sub ShowContentsTab ()
'=====================================================================
'  DISPLAY Contents tab (*.CNT)
'=====================================================================
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_FINDER, Clng(0))

End Sub
Public Sub ShowHelpOnHelp ()
'=====================================================================
'  DISPLAY HELP for WINHELP.EXE  ...
'=====================================================================
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_HELPONHELP, Clng(0))

End Sub

Public Sub SearchHelp ()
'=====================================================================
'  TO ADD KEYWORD SEARCH CAPABILITY...
'=====================================================================
'     o   In your Help|Search menu selection, simply enter:
'         Call SearchHelp() 'To invoke helpfile keyword search dialog
'
    Dim Result as Variant

    Result = WinHelp(m_hWndMainWindow, APP.HelpFile, HELP_PARTIALKEY, ByVal "" )

End Sub

Public Sub SearchHelpKeyWord (Argument as String)
'=====================================================================
'  TO ADD KEYWORD SEARCH CAPABILITY...
'=====================================================================
'     o   In your Help|Search menu selection, simply enter:
'         Call SearchHelp() 'To invoke helpfile keyword search dialog
'
    Dim Result as Variant

    Result = WinHelp(m_hWndMainWindow, APP.HelpFile, HELP_PARTIALKEY, ByVal Trim$(Argument))

End Sub
Public Sub HelpWindowSize (x As Integer, y As Integer, wx As Integer, wy As Integer)
'=====================================================================
'  TO SET THE SIZE AND POSITION OF THE MAIN HELP WINDOW...
'=====================================================================
'     o   Call HelpWindowSize(x, y, dx, dy), where:
'             x = 1-1024 (position from left edge of screen)
'             y = 1-1024 (position from top of screen)
'             dx= 1-1024 (width)
'             dy= 1-1024 (height)
'
    Dim Result as Variant
    MainWindowInfo.x = x
    MainWindowInfo.y = y
    MainWindowInfo.dx = wx
    MainWindowInfo.dy = wy
    Result = WinHelpByInfo(m_hWndMainWindow, App.HelpFile, HELP_SETWINPOS, MainWindowInfo)
End Sub
