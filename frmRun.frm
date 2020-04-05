VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Begin VB.Form frmRun 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Run Function ..."
   ClientHeight    =   7770
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   6435
   Icon            =   "frmRun.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   6465
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4005
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.ListBox lstFunciones 
      Appearance      =   0  'Flat
      Height          =   2280
      Left            =   30
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1530
      Width           =   6300
   End
   Begin VB.TextBox txtParam 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   30
      TabIndex        =   2
      Top             =   915
      Width           =   6300
   End
   Begin VB.ComboBox cboFunciones 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   285
      Width           =   6300
   End
   Begin CodeSenseCtl.CodeSense txtCode 
      Height          =   3090
      Left            =   30
      OleObjectBlob   =   "frmRun.frx":000C
      TabIndex        =   7
      Top             =   4065
      Width           =   6300
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   0
      Left            =   585
      TabIndex        =   9
      Top             =   7230
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Generate"
      AccessKey       =   "G"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePos      =   3
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   1
      Left            =   2310
      TabIndex        =   10
      Top             =   7230
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Cancel"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePos      =   3
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   2
      Left            =   4050
      TabIndex        =   11
      Top             =   7230
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Test"
      AccessKey       =   "T"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePos      =   3
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Code to Test. If need customize before run."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   2
      Left            =   30
      TabIndex        =   6
      Top             =   3870
      Width           =   3720
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Related functions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   5
      Left            =   30
      TabIndex        =   5
      Top             =   1275
      Width           =   1515
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Parameters"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   30
      TabIndex        =   4
      Top             =   705
      Width           =   795
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select function"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   75
      Width           =   1065
   End
End
Attribute VB_Name = "frmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Indice As Integer
Private Archivo As String
Private Function PreparaTest() As Boolean

    On Error GoTo ErrorPreparaTest
    
    'Dim archivo As String
    Dim nFreeFile As Long
    Dim funcion As String
    Dim buffer As String
    Dim linea As String
    Dim k As Integer
    Dim j As Integer
    Dim i As Integer
    Dim fin As Integer
    Dim inicio As Integer
    Dim Msg As String
    
    Archivo = util.ArchivoTemporal()
    Archivo = Left$(Archivo, InStr(Archivo, ".") - 1)
    Archivo = Archivo & ".htm"
    nFreeFile = FreeFile
    
    funcion = cboFunciones.Text
    
    If txtParam.Text <> "" Then
        funcion = funcion & "(" & txtParam.Text & ")"
    Else
        funcion = funcion & "()"
    End If
    
    Dim sBuffer As New cStringBuilder
    
    sBuffer.Append "<!DOCTYPE HTML PUBLIC " & Chr$(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr$(34) & ">" & vbNewLine
    sBuffer.Append "<HTML>" & vbNewLine
    sBuffer.Append "<HEAD>" & vbNewLine
    sBuffer.Append "<TITLE>Test Javascript Function</TITLE>" & vbNewLine
    sBuffer.Append "</HEAD>" & vbNewLine
    sBuffer.Append "<BODY onLoad=" & Chr$(34) & funcion & Chr$(34) & ">" & vbNewLine
    sBuffer.Append "<SCRIPT LANGUAJE=""JavaScript"">" & vbNewLine
    sBuffer.Append vbNewLine
    
    'verificar si hay funciones relacionadas
    'de haber incorporarlas
    For k = 0 To lstFunciones.ListCount - 1
        If lstFunciones.Selected(k) Then
            'verificar que la funcion no se repita
            If lstFunciones.List(k) = cboFunciones.Text Then
                MsgBox "The function selected match with the function to run. Please select another.", vbCritical
                Exit Function
            End If
            
            'insertar el codigo de la funcion
            buffer = "function " & LCase$(lstFunciones.List(k))
            With frmMain.ActiveForm
                For j = 0 To .txtCode.LineCount
                    linea = Trim$(LCase$(util.SacarBasura(.txtCode.GetLine(j))))
                    If VBA.Left$(linea, Len(buffer)) = buffer Then
                        inicio = j
                        Exit For
                    End If
                Next j
            End With
            
            'buscar la siguiente coincidencia de funcion
            buffer = "function "
            With frmMain.ActiveForm
                For i = inicio + 1 To .txtCode.LineCount
                    linea = Trim$(LCase$(util.SacarBasura(.txtCode.GetLine(i))))
                    If Trim$(linea) = "}" Then fin = i
                    If Left$(linea, Len(buffer)) = buffer Then
                        Exit For
                    End If
                Next i
            End With
            
            If fin = 0 Then
                Msg = "I can't find any end block } before the next function." & vbNewLine
                Msg = Msg & "May be the file have any syntax errorss." & vbNewLine
                MsgBox Msg, vbCritical
                Exit Function
            End If
                        
            'borrar desde la linea inicio a la linea fin
            With frmMain.ActiveForm
                For i = inicio To fin
                    sBuffer.Append .txtCode.GetLine(i) & vbNewLine
                Next i
            End With
        End If
    Next k
    
    'insertar el codigo de la funcion
    buffer = "function " & LCase$(cboFunciones.Text)
    With frmMain.ActiveForm
        For j = 0 To .txtCode.LineCount
            linea = Trim$(LCase$(util.SacarBasura(.txtCode.GetLine(j))))
            If VBA.Left$(linea, Len(buffer)) = buffer Then
                inicio = j
                Exit For
            End If
        Next j
    End With
    
    'buscar la siguiente coincidencia de funcion
    buffer = "function "
    With frmMain.ActiveForm
        For i = inicio + 1 To .txtCode.LineCount
            linea = Trim$(LCase$(util.SacarBasura(.txtCode.GetLine(i))))
            If Trim$(linea) = "}" Then fin = i
            If Left$(linea, Len(buffer)) = buffer Then
                Exit For
            End If
        Next i
    End With
    
    If fin = 0 Then
        Msg = "I can't find any end block } before the next function." & vbNewLine
        Msg = Msg & "May be the file have any syntax errorss." & vbNewLine
        MsgBox Msg, vbCritical
        Exit Function
    End If
                
    With frmMain.ActiveForm
        For i = inicio To fin
            sBuffer.Append .txtCode.GetLine(i) & vbNewLine
        Next i
    End With
    
    sBuffer.Append vbNewLine
    sBuffer.Append "</SCRIPT>" & vbNewLine
    sBuffer.Append "</BODY>" & vbNewLine
    sBuffer.Append "</HTML>"
    
    Open Archivo For Output As #nFreeFile
        Print #nFreeFile, sBuffer.ToString
    Close #nFreeFile
    
    Set sBuffer = Nothing
    
    txtCode.OpenFile Archivo
    
    PreparaTest = True
    
    Exit Function
ErrorPreparaTest:
    MsgBox "PreparaTest : " & Err & " " & Error$, vbCritical
    
End Function


Private Sub cboFunciones_Click()

    If cboFunciones.ListCount > -1 Then
    
    End If
    
End Sub


Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If cboFunciones.ListIndex <> -1 Then
            If Not PreparaTest() Then
                MsgBox "An error at try to test the selected function.", vbCritical
            End If
        Else
            MsgBox "Select a function.", vbCritical
            cboFunciones.SetFocus
        End If
    ElseIf Index = 2 Then
        If Archivo = "" Then
            MsgBox "Nothing to test!", vbCritical
        Else
            txtCode.SaveFile Archivo, False
            util.ShellFunc Archivo, vbNormalFocus
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
    
    Dim k As Integer
    
    set_color_form Me
    
    With frmMain.ActiveForm
        Dim funcol As New Collection
        
        Set funcol = .FunBox1.GetFunctions
        For k = 1 To funcol.Count
            cboFunciones.AddItem funcol.Item(k)
            lstFunciones.AddItem funcol.Item(k)
        Next k
    End With
    
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    cmd(0).Refresh
    cmd(1).Refresh
    cmd(2).Refresh
    Debug.Print "load"
    DrawXPCtl Me
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmRun = Nothing
End Sub


