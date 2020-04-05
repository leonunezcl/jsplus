VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmString 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Code Format "
   ClientHeight    =   7725
   ClientLeft      =   2910
   ClientTop       =   2175
   ClientWidth     =   9030
   Icon            =   "frmString.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCodeString 
      Height          =   3075
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   4605
      Width           =   8775
   End
   Begin VB.TextBox txtContent 
      Height          =   3075
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   1155
      Width           =   8775
   End
   Begin VB.Timer WBTimeoutTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8280
      Top             =   480
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   615
      Left            =   5640
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   735
      ExtentX         =   1296
      ExtentY         =   1085
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8280
      Top             =   0
   End
   Begin VB.CheckBox chkAddBreaks 
      Caption         =   "Add real line breaks"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.ComboBox cboLanguage 
      Height          =   315
      ItemData        =   "frmString.frx":000C
      Left            =   1800
      List            =   "frmString.frx":002B
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Format to language:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblStringifiedOutput 
      AutoSize        =   -1  'True
      Caption         =   "JavaScript String"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Content"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   555
   End
End
Attribute VB_Name = "frmString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private wbinit As Boolean

Private Sub cboLanguage_Change()
    lblStringifiedOutput.Caption = cboLanguage.Text & " String"
    Timer1_Timer
End Sub

Private Sub cboLanguage_Click()
    cboLanguage_Change
End Sub

Private Sub chkAddBreaks_Click()
    Timer1_Timer
End Sub

Private Sub Form_Load()

    util.Hourglass hwnd, True
    
    cboLanguage.ListIndex = 0
    Dim lang As String
    lang = GetSetting(App.Title, "Settings", "LastLanguage")
    On Error Resume Next
    cboLanguage.Text = lang
    
    If Len(frmMain.ActiveForm.txtCode.SelText) > 0 Then
        txtContent.Text = frmMain.ActiveForm.txtCode.SelText
    Else
        txtContent.Text = frmMain.ActiveForm.txtCode.Text
    End If
    
    util.CenterForm Me
    'DrawXPCtl Me
    
    util.Hourglass hwnd, False
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'txtContent.Width = Form1.ScaleWidth - (txtContent.Left * 2)
    'txtCodeString.Width = Form1.ScaleWidth - (txtCodeString.Left * 2)
    'Form1.Height = 8130 ' .. too lazy to add height resizing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Settings", "LastLanguage", cboLanguage.Text
    Set frmString = Nothing
End Sub

Private Sub txtContent_Change()
    Timer1.Enabled = True
End Sub

Private Sub txtContent_KeyDown(KeyCode As Integer, Shift As Integer)
    Timer1.Enabled = True
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    Timer1.Enabled = True
End Sub

Private Sub txtContent_KeyUp(KeyCode As Integer, Shift As Integer)
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    txtCodeString.Text = JSFix(txtContent.Text)
    Timer1.Enabled = False
End Sub

Private Function JSFix(str As String) As String

    
    Select Case cboLanguage.Text
    Case "PHP"
        str = Replace(str, "\", "\\")
        str = Replace(str, vbCr, "\r")
        str = Replace(str, """", "\""")
        If chkAddBreaks.Value = 0 Then
            str = Replace(str, vbLf, "\n")
        Else
            str = Replace(str, vbLf, "\n""" & vbCrLf & vbTab & "+ """)
        End If
        str = "$str = """ & str & """;" & vbCrLf
    Case "JSP"
        str = "<% String str=" & Chr$(34) & str & Chr$(34) & ";%>"
    Case "C#"
        str = Replace(str, "\", "\\")
        str = Replace(str, vbCr, "\r")
        str = Replace(str, """", "\""")
        If chkAddBreaks.Value = 0 Then
            str = Replace(str, vbLf, "\n")
        Else
            str = Replace(str, vbLf, "\n""" & vbCrLf & vbTab & "+ """)
        End If
        str = "string str = """ & str & """;" & vbCrLf
    Case "JavaScript/JScript"
        str = Replace(str, vbCr, "")
        str = Replace(str, "\", "\\")
        str = Replace(str, """", "\""")
        If chkAddBreaks.Value = 0 Then
            str = Replace(str, vbLf, "\n")
        Else
            str = Replace(str, vbLf, "\n""" & vbCrLf & vbTab & "+ """)
        End If
        str = "var str = """ & str & """;" & vbCrLf
    Case "VBScript"
        str = Replace(str, """", """""")
        str = Replace(str, vbCr, "")
        If chkAddBreaks.Value = 0 Then
            str = Replace(str, vbLf, """ & vbCrLf & """)
        Else
            str = Replace(str, vbLf, """ & vbCrLf & _" & vbCrLf & vbTab & """")
        End If
        str = "Dim str" & vbCrLf & "str = """ & str & """" & vbCrLf
    Case "VB6"
        str = Replace(str, """", """""")
        str = Replace(str, vbCr, "")
        If chkAddBreaks.Value = 0 Then
            str = Replace(str, vbLf, """ & vbCrLf & """)
        Else
            str = Replace(str, vbLf, """ & vbCrLf & _" & vbCrLf & vbTab & """")
        End If
        str = "Dim str As String" & vbCrLf & "str = """ & str & """" & vbCrLf
    Case "VB.Net"
        str = Replace(str, """", """""")
        str = Replace(str, vbCr, "")
        If chkAddBreaks.Value = 0 Then
            str = Replace(str, vbLf, """ & vbCrLf & """)
            str = "Dim str As String = " & _
              """" & str & """" & vbCrLf
        Else
            str = Replace(str, vbLf, """ & vbCrLf _" & vbCrLf & vbTab & "& """)
            str = "Dim str As String" & vbCrLf & "str = """ & str & """" & vbCrLf
        End If
    Case "HTML (mini)"
        str = Replace(str, "&", "&amp;")
        str = Replace(str, """", "&quot;")
        str = Replace(str, vbCr, "")
        str = Replace(str, "<", "&lt;")
        str = Replace(str, ">", "&gt;")
        If chkAddBreaks.Value = 0 Then
            str = Replace(str, vbLf, "<br>")
        Else
            str = Replace(str, vbLf, "<br>" & vbCrLf)
        End If
    Case "HTML (IE)"
        str = TEXT2IEHTML(str)
    End Select
    JSFix = str
End Function

Function TEXT2IEHTML(str As String)
    'On Error Resume Next
    If Not wbinit Then
        WebBrowser1.Navigate "about:<html><head></head><body>.</body></html>"
        WBTimeoutTimer.Enabled = True
        Do Until WebBrowser1.Busy Or Not WBTimeoutTimer.Enabled
            DoEvents
        Loop
        Do Until Not WebBrowser1.Busy
            DoEvents
        Loop
        wbinit = True
    End If
    
    WebBrowser1.Document.Body.innerText = txtContent.Text
    str = WebBrowser1.Document.Body.innerhtml
    If chkAddBreaks.Value > 0 Then
        str = Replace(str, "<BR>", "<BR>" & vbCrLf)
        str = Replace(str, "<P>", vbCrLf & "<P>")
    End If
    TEXT2IEHTML = str
End Function

Private Sub WBTimeoutTimer_Timer()
    WBTimeoutTimer.Enabled = False
End Sub
