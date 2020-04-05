Attribute VB_Name = "MTimeOutSupport"
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private m_lngWindowHandle   As Long
Private m_intTimeOutObjects As Integer

Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    '
    Dim lngObjPointer   As Long
    Dim objTimeOut      As CTimeout
    '
    On Error Resume Next
    '
    If Err.Number = 0 Then
        '
        Set objTimeOut = TimeOutObjectByPointer(idEvent)
        '
        objTimeOut.PostTimeOutEvent
        '
    End If
    '
End Sub

Public Function RegisterTimer(ByVal lngTimeOutValue As Long, ByVal lngObjPointer As Long) As Long
    '
    Dim lngElapse   As Long
    Dim lngEventID  As Long
    '
    lngElapse = CLng(lngTimeOutValue * 1000&)
    '
    lngEventID = SetTimer(m_lngWindowHandle, lngObjPointer, lngElapse, AddressOf TimerProc)
    '
    'Debug.Print "TIMER: " & "SetTimer"
    '
    If lngEventID <> 0 Then
        '
        RegisterTimer = lngEventID
        m_intTimeOutObjects = m_intTimeOutObjects + 1
        '
    Else
        Err.Raise 17002, "MTimeOutSupport.RegisterTimer", "Cannot create the timer"
    End If
    '
End Function

Public Sub UnRegisterTimer(ByVal lngEventID As Long)
    '
    Dim lngRetValue As Long
    '
    lngRetValue = KillTimer(m_lngWindowHandle, lngEventID)
    m_intTimeOutObjects = m_intTimeOutObjects - 1
    '
    'Debug.Print "TIMER: " & "KillTimer"
    '
End Sub

Public Sub ResetTimer(ByVal lngEventID As Long, ByVal lngTimeOutValue As Long)
    '
    Dim lngObjPointer   As Long
    Dim lngElapse      As Long
    '
    On Error Resume Next
    '
    lngElapse = CLng(lngTimeOutValue * 1000&)
    '
    If Err.Number = 0 Then
        '
        Call SetTimer(m_lngWindowHandle, lngEventID, lngElapse, AddressOf TimerProc)
        'Debug.Print "TIMER: " & "SetTimer"
        '
    End If
    '
End Sub

Public Sub CreateTimer()
    '
    If m_lngWindowHandle = 0 Then
        If CreateTimerWindow = 0 Then
            Err.Raise 17001, "MTimeOutSupport.CreateTimer", "Cannot create the timer window."
        End If
    End If
    '
End Sub

Public Sub DestroyTimer()
    '
    Call DestroyTimerWindow
    '
End Sub

Private Function CreateTimerWindow() As Long
'********************************************************************************
'Author    :Oleg Gdalevich
'Date/Time :17-12-2001
'Purpose   :Creates a window to hold timers
'Returns   :The window handle
'********************************************************************************
    '
    'Create a window. We'll not see this window as the ShowWindow is never called.
    m_lngWindowHandle = CreateWindowEx(0&, "STATIC", "TIMER_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
    '
    If m_lngWindowHandle = 0 Then
        '
        'I really don't know - is it possible? Probably - yes,
        'due the lack of the system resources, for example.
        '
        'In this case the function returns 0.
        '
    Else
        '
        'Just to let the caller know that the function was executed successfully
        CreateTimerWindow = m_lngWindowHandle
        '
        'Debug.Print "The timer window is created: " & m_lngWindowHandle
        '
    End If
    '
End Function


Private Function DestroyTimerWindow() As Boolean
'********************************************************************************
'Author    :Oleg Gdalevich
'Date/Time :17-12-2001
'Purpose   :Destroyes the window
'Returns   :If the window was destroyed successfully - True.
'********************************************************************************
    '
    On Error GoTo ERR_HANDLER
    '
    'Destroy the window
    DestroyWindow m_lngWindowHandle
    '
    'Debug.Print "The timer window " & m_lngWindowHandle & " is destroyed"
    '
    'Reset the window handle variable
    m_lngWindowHandle = 0
    'If no errors occurred, the function returns True
    DestroyTimerWindow = True
    '
ERR_HANDLER:

End Function

Private Function TimeOutObjectByPointer(ByVal lngObjPointer As Long) As CTimeout
    '
    Dim objTimeOut As CTimeout
    '
    CopyMemory objTimeOut, lngObjPointer, 4&
    Set TimeOutObjectByPointer = objTimeOut
    CopyMemory objTimeOut, 0&, 4&
    '
End Function
