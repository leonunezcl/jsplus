VERSION 5.00
Begin VB.Form frmArray 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Array"
   ClientHeight    =   3135
   ClientLeft      =   4680
   ClientTop       =   2520
   ClientWidth     =   3285
   Icon            =   "frmArray.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Settings"
      ForeColor       =   &H00000000&
      Height          =   2385
      Left            =   75
      TabIndex        =   8
      Top             =   120
      Width           =   3135
      Begin VB.OptionButton opt 
         Caption         =   "Empty Array"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   1560
      End
      Begin VB.OptionButton opt 
         Caption         =   "Array of Days"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   495
         Width           =   1260
      End
      Begin VB.OptionButton opt 
         Caption         =   "Short Months Names Array"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   2
         Top             =   735
         Width           =   2415
      End
      Begin VB.OptionButton opt 
         Caption         =   "Array of Numbers"
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   4
         Top             =   1245
         Width           =   1560
      End
      Begin VB.OptionButton opt 
         Caption         =   "Array of Colors"
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   5
         Top             =   1500
         Width           =   1560
      End
      Begin VB.OptionButton opt 
         Caption         =   "Long Months Names Array"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   3
         Top             =   990
         Width           =   2250
      End
      Begin VB.OptionButton opt 
         Caption         =   "Short Countries Names Array"
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   6
         Top             =   1770
         Width           =   2325
      End
      Begin VB.OptionButton opt 
         Caption         =   "Long Countries Names Array"
         Height          =   255
         Index           =   7
         Left            =   180
         TabIndex        =   7
         Top             =   2025
         Width           =   2325
      End
   End
End
Attribute VB_Name = "frmArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub InsertArray()

    Dim linea As New cStringBuilder
    Dim Num
    Dim k As Integer
    Dim ini As String
    
    Dim sSections() As String
    
    ini = util.StripPath(App.Path) & "config\arrays.ini"
    
    Call util.Hourglass(hwnd, True)
    
    If opt(0).Value Then        'custom
        linea.Append "var strNewArray = new Array();" & vbNewLine
    ElseIf opt(1).Value Then    'days
        get_info_section "days_array", sSections, ini
        linea.Append "aDays[" & Num & "]= new Array();" & vbNewLine & vbNewLine
        For k = 2 To UBound(sSections)
            linea.Append "aDays[" & k - 1 & "]=" & Chr$(34) & sSections(k) & Chr$(34) & ";" & vbNewLine
        Next k
    ElseIf opt(2).Value Then    'short months
        get_info_section "short_month_array", sSections, ini
        linea.Append "var strMonthArray[" & Num & "] = new Array();" & vbNewLine
        For k = 2 To UBound(sSections)
            linea.Append "strMonthArray[" & k - 1 & "]=" & Chr$(34) & sSections(k) & Chr$(34) & ";" & vbNewLine
        Next k
    ElseIf opt(3).Value Then    'long months
        get_info_section "long_month_array", sSections, ini
        linea.Append "var strMonthArray[" & Num & "] = new Array();" & vbNewLine
        For k = 2 To UBound(sSections)
            linea.Append "strMonthArray[" & k - 1 & "]=" & Chr$(34) & sSections(k) & Chr$(34) & ";" & vbNewLine
        Next k
    ElseIf opt(4).Value Then    'numbers
        get_info_section "numbers_array", sSections, ini
        linea.Append "var strNumberArray[" & Num & "] = new Array();" & vbNewLine
        For k = 2 To UBound(sSections)
            linea.Append "strNumberArray[" & k - 1 & "]=" & Chr$(34) & sSections(k) & Chr$(34) & ";" & vbNewLine
        Next k
    ElseIf opt(5).Value Then    'colors
        get_info_section "colors_array", sSections, ini
        linea.Append "var strColorArray[" & Num & "] = new Array();" & vbNewLine
        For k = 2 To UBound(sSections)
            linea.Append "strColorArray[" & k - 1 & "]=" & Chr$(34) & sSections(k) & Chr$(34) & ";" & vbNewLine
        Next k
    ElseIf opt(6).Value Then    'short country
        get_info_section "short_country_array", sSections, ini
        linea.Append "var strShortCountryArray[" & Num & "] = new Array();" & vbNewLine
        For k = 2 To UBound(sSections)
            linea.Append "strShortCountryArray[" & k - 1 & "]=" & Chr$(34) & Trim$(sSections(k)) & Chr$(34) & ";" & vbNewLine
        Next k
    ElseIf opt(7).Value Then    'long country
        get_info_section "long_country_array", sSections, ini
        linea.Append "var strLongCountryArray[" & Num & "] = new Array();" & vbNewLine
        For k = 2 To UBound(sSections)
            linea.Append "strLongCountryArray[" & k - 1 & "]=" & Chr$(34) & Trim$(sSections(k)) & Chr$(34) & ";" & vbNewLine
        Next k
    End If
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(linea.ToString)
    End If
    
    Call util.Hourglass(hwnd, False)
    
End Sub
Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call InsertArray
    End If
    Unload Me
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    util.CenterForm Me
        
    Debug.Print "load"
       
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmArray = Nothing
End Sub


