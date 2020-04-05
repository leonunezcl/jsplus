VERSION 5.00
Object = "{BCA5B647-4A34-488D-8923-BAED19344F42}#2.0#0"; "vbsToolboxBar6.ocx"
Begin VB.UserControl vbsClipboard 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vbsClipboard.ctx":0000
   Begin VB.TextBox txtTmp 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin vbalToolboxBar6.vbalToolBoxBarCtl tboClp 
      Height          =   1095
      Left            =   345
      TabIndex        =   0
      Top             =   360
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   1931
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "vbsClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Img As cVBALImageList
Private WithEvents m_cClipView As cClipboardViewer
Attribute m_cClipView.VB_VarHelpID = -1
Private contador As Integer
Public Property Get JScVBALImageList() As cVBALImageList
    Set JScVBALImageList = m_Img
End Property


Public Property Set JScVBALImageList(ByVal pcVBALImageList As cVBALImageList)
    Set m_Img = pcVBALImageList
End Property

Public Sub Load()

    Dim cBar As cToolBoxBar

    tboClp.ImageList = m_Img.hIml
    Set cBar = tboClp.Bars.Add("k1", , "Clipboard", 2)
    
    Set m_cClipView = New cClipboardViewer
    m_cClipView.InitClipboardChangeNotification UserControl.hwnd
    
End Sub

Private Sub m_cClipView_ClipboardChanged()

    On Error Resume Next
    If (Clipboard.GetFormat(vbCFText)) Then
      additems Clipboard.GetText
   End If
    Err = 0
    
End Sub


Private Sub UserControl_Resize()
    
    LockWindowUpdate hwnd
    tboClp.Move 0, 0, UserControl.Width, UserControl.Height
    LockWindowUpdate False
End Sub


Private Sub additems(ByVal ele As String)

    Dim k As Integer
    Dim j As Integer
    Dim found As Boolean
    Dim cBar As cToolBoxBar
    Dim limite
    
    Set cBar = tboClp.Bars(1)
    
    For k = 1 To cBar.Items.Count
        If cBar.Items(k).Caption = ele Then
            found = True
            Exit For
        End If
    Next k
    
    limite = Util.LeeIni(IniPath, "clipboard", "value")
    If limite = "" Then limite = 15
    
    If Not found Then
        contador = contador + 1
        If contador > CInt(limite) Then
            For k = cBar.Items.Count To 1 Step -1
                cBar.Items.Remove k
            Next k
            contador = 1
        End If
        cBar.Items.Add "k" & contador, , ele, 2
    End If
    
End Sub

