VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTimeout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Private m_intTimeOutValue   As Long
Private m_lngTimerEventID As Long
'
Public Event TimeOut()

Public Sub StartTimer()
    '
    If m_intTimeOutValue > 0 Then
        '
        If m_lngTimerEventID = 0 Then
            m_lngTimerEventID = MTimeOutSupport.RegisterTimer(m_intTimeOutValue, ObjPtr(Me))
        Else
            Call MTimeOutSupport.ResetTimer(m_lngTimerEventID, m_intTimeOutValue)
        End If
        '
    End If
    '
End Sub

Public Sub StopTimer()
    '
    If m_lngTimerEventID <> 0 Then
        Call MTimeOutSupport.UnRegisterTimer(m_lngTimerEventID)
    End If
    '
End Sub

Public Property Let TimeOutValue(NewValue As Long)
    m_intTimeOutValue = NewValue
End Property

Public Property Get TimeOutValue() As Long
    TimeOutValue = m_intTimeOutValue
End Property

Public Sub Reset()
    Call MTimeOutSupport.ResetTimer(m_lngTimerEventID, m_intTimeOutValue)
End Sub

Private Sub Class_Initialize()
    Call MTimeOutSupport.CreateTimer
End Sub

Private Sub Class_Terminate()
    Call MTimeOutSupport.DestroyTimer
End Sub

Friend Sub PostTimeOutEvent()
    RaiseEvent TimeOut
End Sub
