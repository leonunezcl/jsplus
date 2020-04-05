VERSION 5.00
Begin VB.Form frmjslitopt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5490
   ClientLeft      =   3915
   ClientTop       =   2475
   ClientWidth     =   5100
   Icon            =   "frmjslitopt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Help"
      Height          =   375
      Index           =   2
      Left            =   3720
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Select All"
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   30
      Width           =   1005
   End
   Begin VB.Frame fra 
      Caption         =   "Options"
      Height          =   5130
      Left            =   30
      TabIndex        =   0
      Top             =   315
      Width           =   3495
      Begin VB.CheckBox chk 
         Caption         =   "Lax line breaking"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   225
         Width           =   3285
      End
   End
End
Attribute VB_Name = "frmjslitopt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private total As Integer
Private Sub config_jslint()

    Dim str As New cStringBuilder
    Dim Archivo As String
    Dim nFreeFile As Long
    Dim inifile As String
    Dim k As Integer
    
    inifile = util.StripPath(App.Path) & "jslint.ini"
        
    util.BorrarArchivo inifile
    
    For k = 0 To total - 1
        Call util.GrabaIni(inifile, "options", chk(k).Tag, IIf(chk(k).Value, "1", "0"))
    Next k
    
    nFreeFile = FreeFile
    Archivo = util.StripPath(App.Path) & "jslint\check.js"
    
    str.Append "function go(){" & vbNewLine
    str.Append "var o = {};" & vbNewLine
    
    For k = 0 To total - 1
        If chk(k).Value Then
            str.Append "o[" & k & "] = true;" & vbNewLine
        End If
    Next k
    
    str.Append "jslint(document.forms.jslint.input.value, o);" & vbNewLine
    str.Append "document.getElementById('output').innerHTML = jslint.report();" & vbNewLine
    str.Append "}" & vbNewLine
    
    Open Archivo For Output As #nFreeFile
        Print #nFreeFile, str.ToString
    Close #nFreeFile
    
    Set str = Nothing
    
End Sub
Private Sub chkAll_Click()

    Dim ret As Integer
    Dim k As Integer
        
    If chkAll.Value Then
        ret = 1
    End If
    
    For k = 0 To chk.count - 1
        chk(k).Value = ret
    Next k
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call config_jslint
        Unload Me
    ElseIf Index = 1 Then
        Unload Me
    ElseIf Index = 2 Then
        util.ShellFunc "http://www.JSLint.com/lint.html", vbNormalFocus
    End If
           
End Sub

Private Sub Form_Load()
    
    Dim inifile As String
    Dim arr_opciones() As String
    Dim k As Integer
    Dim j As Integer
    Dim selec As String
    
    util.CenterForm Me
    
    inifile = util.StripPath(App.Path) & "analizer.ini"
    
    get_info_section "options", arr_opciones(), inifile
    
    j = 0
    For k = 1 To UBound(arr_opciones)
        If j > 0 Then
            Load chk(j)
            chk(j).Left = chk(j - 1).Left
            chk(j).Height = chk(j - 1).Height
            chk(j).Width = chk(j - 1).Width
            chk(j).Top = chk(j - 1).Top + chk(j).Height
            chk(j).Caption = util.Explode(arr_opciones(k), 2, "|")
            chk(j).Tag = util.Explode(arr_opciones(k), 1, "|")
            chk(j).ToolTipText = util.Explode(arr_opciones(k), 3, "|")
            chk(j).Visible = True
        Else
            chk(0).Caption = util.Explode(arr_opciones(k), 2, "|")
            chk(0).Tag = util.Explode(arr_opciones(k), 1, "|")
            chk(0).ToolTipText = util.Explode(arr_opciones(k), 3, "|")
        End If
        j = j + 1
    Next k
    
    total = UBound(arr_opciones)
    
    inifile = util.StripPath(App.Path) & "jslint.ini"
        
    For k = 0 To total - 1
        selec = util.LeeIni(inifile, "options", chk(k).Tag)
        If selec = "1" Then
            chk(k).Value = 1
        End If
    Next k
    
    Debug.Print "load"
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmjslitopt = Nothing
End Sub


