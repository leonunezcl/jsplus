VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPicker"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum eTipoPick
    pick_basico = 1
    pick_avanzado = 2
End Enum
Private m_PathPal As String
Public Sub StartPicker(ByVal Tipo As eTipoPick)

    gCodeColor = vbNullString
    If Tipo = pick_basico Then
        'frmCPick.Show vbModal
    Else
        If Len(m_PathPal) > 0 Then
            frmSelCol.PathPaleta = m_PathPal
            frmSelCol.Show vbModal
        Else
            MsgBox "Falta definir path de paletas", vbCritical
        End If
    End If
End Sub



Public Property Get PathPal() As String
    PathPal = m_PathPal
End Property

Public Property Let PathPal(ByVal pPathPal As String)
    m_PathPal = pPathPal
End Property

Public Property Get CodeColor() As String
    CodeColor = gCodeColor
End Property


Public Property Get SelectColor() As Long
    SelectColor = gSelectColor
End Property

