VERSION 5.00
Begin VB.Form frmNewVar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Variable"
   ClientHeight    =   3135
   ClientLeft      =   4500
   ClientTop       =   3600
   ClientWidth     =   5595
   Icon            =   "frmNewVar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      Height          =   2445
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   75
      Width           =   5400
      Begin VB.Frame fra 
         Caption         =   "Select"
         Height          =   1530
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   780
         Width           =   5145
         Begin VB.OptionButton opt 
            Caption         =   "New Variable"
            Height          =   225
            Index           =   0
            Left            =   840
            TabIndex        =   1
            Top             =   240
            Value           =   -1  'True
            Width           =   1260
         End
         Begin VB.OptionButton opt 
            Caption         =   "New Date"
            Height          =   225
            Index           =   2
            Left            =   840
            TabIndex        =   3
            Top             =   705
            Width           =   1110
         End
         Begin VB.OptionButton opt 
            Caption         =   "New String"
            Height          =   225
            Index           =   3
            Left            =   840
            TabIndex        =   4
            Top             =   945
            Width           =   1110
         End
         Begin VB.OptionButton opt 
            Caption         =   "ParseInt"
            Height          =   225
            Index           =   6
            Left            =   2700
            TabIndex        =   7
            Top             =   480
            Width           =   990
         End
         Begin VB.OptionButton opt 
            Caption         =   "ParseLong"
            Height          =   225
            Index           =   7
            Left            =   2700
            TabIndex        =   8
            Top             =   735
            Width           =   1110
         End
         Begin VB.OptionButton opt 
            Caption         =   "Custom"
            Height          =   225
            Index           =   8
            Left            =   2700
            TabIndex        =   9
            Top             =   975
            Width           =   1110
         End
         Begin VB.OptionButton opt 
            Caption         =   "New Boolean"
            Height          =   225
            Index           =   4
            Left            =   840
            TabIndex        =   5
            Top             =   1185
            Width           =   1350
         End
         Begin VB.OptionButton opt 
            Caption         =   "New Number"
            Height          =   225
            Index           =   5
            Left            =   2700
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton opt 
            Caption         =   "New Array"
            Height          =   225
            Index           =   1
            Left            =   840
            TabIndex        =   2
            Top             =   480
            Width           =   1110
         End
      End
      Begin VB.TextBox txtVar 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   450
         Width           =   5160
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Variable Name"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   240
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmNewVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function InsertaVar() As Boolean

    Dim str As New cStringBuilder
    Dim var As String
    Dim custom As String
    
    If txtVar.Text = "" Then
        MsgBox "Input the variable name", vbCritical
        txtVar.SetFocus
        Exit Function
    End If
    
    var = txtVar.Text
    
    If opt(0).Value Then
        str.Append "var " & var & ";" & vbNewLine
    ElseIf opt(1).Value Then
        str.Append "var " & var & " = new Array();" & vbNewLine
    ElseIf opt(2).Value Then
        str.Append "var " & var & " = new Date([theDate]);" & vbNewLine
    ElseIf opt(3).Value Then
        str.Append "var " & var & " = new String([theString]);" & vbNewLine
    ElseIf opt(4).Value Then
        str.Append "var " & var & " = new Boolean([theBoolean]);" & vbNewLine
    ElseIf opt(5).Value Then
        str.Append "var " & var & " = new Number([theNumber]);" & vbNewLine
    ElseIf opt(6).Value Then
        str.Append "var " & var & " = parseInt(varToParse);" & vbNewLine
    ElseIf opt(7).Value Then
        str.Append "var " & var & " = parseFloat(varToParse);" & vbNewLine
    Else
        custom = InputBox("Custom Var", "Text")
        If custom <> "" Then
            str.Append "var " & var & " = " & custom & vbNewLine
        End If
    End If
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(str.ToString)
    End If
    
    Set str = Nothing
    
    InsertaVar = True
    
End Function

Private Sub cmd_Click(Index As Integer)

    If Index = 1 Then
        If InsertaVar() Then
            Unload Me
        End If
    Else
        Unload Me
    End If
    
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
        
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmNewVar = Nothing
End Sub


