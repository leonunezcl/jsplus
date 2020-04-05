Attribute VB_Name = "Module1"
Option Explicit

Public util As New cLibrary
Public ListaLangs As New cLanguage
Public Cdlg As New cCommonDialog
Public CSGlobals As New CodeSenseCtl.Globals
Private Const C_INI = "jsplus.ini"
Private bufferini As String

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Sub get_info_section(ByVal seccion As String, ByRef sSections() As String, ByVal StrIniFile As String)

    Dim ret As Long
    Dim iPos As Integer
    Dim iNextPos As Integer
    Dim iSize As Long
    Dim icount As Integer
    Dim sCur As String
    
    bufferini = Space$(8092 * 4)
    iSize = Len(bufferini)
        
    ret = GetPrivateProfileSection(seccion, bufferini, Len(bufferini), StrIniFile)
    icount = 0
    
    If (iSize > 0) Then
        bufferini = VBA.Left$(bufferini, ret)
    Else
        bufferini = ""
    End If
    
    If (Len(bufferini) > 0) Then
        iPos = 1
        iNextPos = InStr(iPos, bufferini, Chr$(0))
        Do While iNextPos <> 0
            If (iNextPos <> iPos) Then
                sCur = Mid$(bufferini, iPos, (iNextPos - iPos))
                icount = icount + 1
                ReDim Preserve sSections(icount) As String
                sSections(icount) = Mid$(sCur, InStr(1, sCur, "=") + 1)
            End If
            iPos = iNextPos + 1
            iNextPos = InStr(iPos, bufferini, Chr$(0))
        Loop
    Else
        ReDim sSections(0)
    End If
    
    bufferini = Space$(0)
    
End Sub

Public Function IniPath() As String

    IniPath = util.StripPath(App.Path) & C_INI
    
End Function
Public Sub clear_memory(frm As Form)

    Dim ctl As Control
    
    For Each ctl In frm
        If TypeOf ctl Is PictureBox Then
            Set ctl.Picture = Nothing
        'ElseIf TypeOf ctl Is MyButton Then
        '    Set ctl.Picture = Nothing
        ElseIf TypeOf ctl Is Image Then
            Set ctl.Picture = Nothing
        End If
    Next
    
    Set ctl = Nothing
    
End Sub
Public Sub set_color_form(frm As Form)

    Dim ctrl As Control
    
    frm.BackColor = vbButtonFace
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is Label Then
            If Not ctrl.Font.Bold Then
                ctrl.BackColor = vbButtonFace
                ctrl.Appearance = 0
                ctrl.BackStyle = 0
            End If
        ElseIf TypeOf ctrl Is Frame Then
            ctrl.BackColor = vbButtonFace
            ctrl.ForeColor = &H0&
        ElseIf TypeOf ctrl Is CheckBox Then
            ctrl.BackColor = vbButtonFace
            ctrl.ForeColor = &H0&
        ElseIf TypeOf ctrl Is OptionButton Then
            ctrl.BackColor = vbButtonFace
            ctrl.ForeColor = &H0&
        ElseIf TypeOf ctrl Is TextBox Then
            ctrl.ForeColor = &H404040
        ElseIf TypeOf ctrl Is ComboBox Then
            ctrl.ForeColor = &H404040
        End If
    Next
    
End Sub
