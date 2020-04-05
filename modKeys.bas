Attribute VB_Name = "Module3"

'GetKeyPressName ---
' PURPOSE: Get name of KeyPress
' INPUTS:
'  objKeyPress: HotKey object with keypress in it
' RETURNS: Keypress name (eg. CTRL+Z)
' EXAMPLE: mnuUndo.Caption = "&Undo" & vbTab & GetKeyPressName(CSGlobals.GetHotKeyForCmd(cmCmdUndo, 0))
Function GetKeyPressName(objKeyPress As HotKey) As String
'if no keypress, exit function
If objKeyPress Is Nothing Then Exit Function
If objKeyPress.VirtKey1 = "" Then Exit Function
Dim strResult As String
strResult = vbTab
'get key mask (eg. CTRL+ALT)
strResult = strResult & GetKeyMaskName(objKeyPress.Modifiers1)
'append keyname (eg. Z or ESCAPE)
strResult = strResult & GetVirtKeyName(objKeyPress.VirtKey1)

'if second keypress, ...
If objKeyPress.VirtKey2 <> "" Then
    '...append a comma...
    strResult = strResult & ", "
    '...and do the same as with first keypress
    strResult = strResult & GetKeyMaskName(objKeyPress.Modifiers2)
    strResult = strResult & GetVirtKeyName(objKeyPress.VirtKey2)
End If
'return result
GetKeyPressName = strResult
End Function

'GetVirtKeyName ---
' PURPOSE: Get the name of a key
' INPUTS:
'  VirtKey - String / Integer containing KeyCode
' RETURNS: Key name (eg. Z or INSERT)
' EXAMPLE: msgbox GetVirtKeyName(&H70)
Function GetVirtKeyName(VirtKey) As String
Dim intVirtKey As Integer
'if VirtKey is string...
If VarType(VirtKey) = vbString Then
    '...convert to integer
    intVirtKey = Asc(VirtKey)
Else
    intVirtKey = VirtKey
End If

'find out key name
Select Case intVirtKey
    Case vbKeyAdd: GetVirtKeyName = "+ (KEYPAD)"
    Case vbKeyBack: GetVirtKeyName = "BACKSPACE"
    Case vbKeyCancel: GetVirtKeyName = "CANCEL"
    Case vbKeyCapital: GetVirtKeyName = "CAPSLOCK"
    'Clear? The '5' on the numpad when Num Lock
    'is off - try it!
    Case vbKeyClear: GetVirtKeyName = "CLEAR"
    Case vbKeyControl: GetVirtKeyName = "CONTROL"
    Case vbKeyDecimal: GetVirtKeyName = ". (KEYPAD)"
    Case vbKeyDelete: GetVirtKeyName = "DELETE"
    Case vbKeyDivide: GetVirtKeyName = "/ (KEYPAD)"
    Case vbKeyDown: GetVirtKeyName = "DOWN ARROW"
    Case vbKeyEnd: GetVirtKeyName = "END"
    Case vbKeyEscape: GetVirtKeyName = "ESCAPE"
    'What's an 'EXECUTE' key????
    Case vbKeyExecute: GetVirtKeyName = "EXECUTE"
    Case vbKeyF1: GetVirtKeyName = "F1"
    Case vbKeyF2: GetVirtKeyName = "F2"
    Case vbKeyF3: GetVirtKeyName = "F3"
    Case vbKeyF4: GetVirtKeyName = "F4"
    Case vbKeyF5: GetVirtKeyName = "F5"
    Case vbKeyF6: GetVirtKeyName = "F6"
    Case vbKeyF7: GetVirtKeyName = "F7"
    Case vbKeyF8: GetVirtKeyName = "F8"
    Case vbKeyF9: GetVirtKeyName = "F9"
    Case vbKeyF10: GetVirtKeyName = "F10"
    Case vbKeyF11: GetVirtKeyName = "F11"
    Case vbKeyF12: GetVirtKeyName = "F12"
    Case vbKeyF13: GetVirtKeyName = "F13"
    Case vbKeyF14: GetVirtKeyName = "F14"
    Case vbKeyF15: GetVirtKeyName = "F15"
    Case vbKeyF16: GetVirtKeyName = "F16"
    'What the hell's a help key?
    Case vbKeyHelp: GetVirtKeyName = "HELP"
    Case vbKeyHome: GetVirtKeyName = "HOME"
    Case vbKeyInsert: GetVirtKeyName = "INSERT"
    'Mouse button is a key?
    Case vbKeyLButton: GetVirtKeyName = "LEFT MOUSE"
    Case vbKeyLeft: GetVirtKeyName = "LEFT ARROW"
    Case vbKeyMButton: GetVirtKeyName = "MIDDLE MOUSE"
    'Menu key? apparently 'ALT' is menu key
    'Case vbKeyMenu: GetVirtKeyName = "MENU KEY"
    Case vbKeyMenu: GetVirtKeyName = "ALT"
    Case vbKeyMultiply: GetVirtKeyName = "* (KEYPAD)"
    Case vbKeyNumlock: GetVirtKeyName = "NUMLOCK"
    Case vbKeyNumpad0: GetVirtKeyName = "0 (KEYPAD)"
    Case vbKeyNumpad1: GetVirtKeyName = "1 (KEYPAD)"
    Case vbKeyNumpad2: GetVirtKeyName = "2 (KEYPAD)"
    Case vbKeyNumpad3: GetVirtKeyName = "3 (KEYPAD)"
    Case vbKeyNumpad4: GetVirtKeyName = "4 (KEYPAD)"
    Case vbKeyNumpad5: GetVirtKeyName = "5 (KEYPAD)"
    Case vbKeyNumpad6: GetVirtKeyName = "6 (KEYPAD)"
    Case vbKeyNumpad7: GetVirtKeyName = "7 (KEYPAD)"
    Case vbKeyNumpad8: GetVirtKeyName = "8 (KEYPAD)"
    Case vbKeyNumpad9: GetVirtKeyName = "9 (KEYPAD)"
    Case vbKeyPageDown: GetVirtKeyName = "PAGE DOWN"
    Case vbKeyPageUp: GetVirtKeyName = "PAGE UP"
    Case vbKeyPause: GetVirtKeyName = "PAUSE"
    Case vbKeyPrint: GetVirtKeyName = "PRINT SCREEN"
    Case vbKeyRButton: GetVirtKeyName = "RIGHT MOUSE"
    Case vbKeyReturn: GetVirtKeyName = "ENTER"
    Case vbKeyRight: GetVirtKeyName = "RIGHT ARROW"
    Case vbKeyScrollLock: GetVirtKeyName = "SCROLL LOCK"
    '...select? hmmmmmmmmm
    Case vbKeySelect: GetVirtKeyName = "SELECT"
    Case vbKeySeperator: GetVirtKeyName = "ENTER (KEYPAD)"
    Case vbKeyShift: GetVirtKeyName = "SHIFT"
    'snapshot *BANG* oops
    Case vbKeySnapshot: GetVirtKeyName = "SNAPSHOT"
    Case vbKeySpace: GetVirtKeyName = "SPACE"
    Case vbKeySubtract: GetVirtKeyName = "- (KEYPAD)"
    Case vbKeyTab: GetVirtKeyName = "TAB"
    Case vbKeyUp: GetVirtKeyName = "UP ARROW"
    
    Case 186: GetVirtKeyName = "SEMICOLON"
    Case 187: GetVirtKeyName = "="
    Case 188: GetVirtKeyName = "COMMA"
    Case 189: GetVirtKeyName = "-"
    Case 190: GetVirtKeyName = "DOT (.)"
    Case 191: GetVirtKeyName = "/"
    Case 192: GetVirtKeyName = "`"
    Case 219: GetVirtKeyName = "["
    Case 220: GetVirtKeyName = "\"
    Case 221: GetVirtKeyName = "]"
    Case 222: GetVirtKeyName = "'"
    'ahhhh, the any key!
    Case 223: GetVirtKeyName = "ANY"
    
    'Case Else: GetVirtKeyName = "UNKNOWN" & intVirtKey
    Case Else: GetVirtKeyName = Chr(intVirtKey)
End Select
End Function

'GetKeyMaskName ---
' PURPOSE: Get the key mask, eg. CTRL+ALT+
' INPUTS:
'  bytKeyMask - key mask, eg. alt=4 ctrl=2 shift=1
' RETURNS: Key mask, eg. CTRL or CTRL+SHIFT+ALT
' EXAMPLE: MsgBox GetKeyMaskName(6) 'is the same as
'          MsgBox GetKeyMaskName(vbAltMask + vbCtrlMask)
Function GetKeyMaskName(bytKeyMask As Byte) As String
Dim strResult As String
If (bytKeyMask And vbCtrlMask) = vbCtrlMask Then strResult = "CTRL+"
If (bytKeyMask And vbShiftMask) = vbShiftMask Then strResult = strResult & "SHIFT+"
If (bytKeyMask And vbAltMask) = vbAltMask Then strResult = strResult & "ALT+"
GetKeyMaskName = strResult
End Function
