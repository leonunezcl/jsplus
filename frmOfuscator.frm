VERSION 5.00
Begin VB.Form frmOfuscator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "JavaScript Ofuscator"
   ClientHeight    =   3630
   ClientLeft      =   3825
   ClientTop       =   2805
   ClientWidth     =   7335
   Icon            =   "frmOfuscator.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Select File"
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmd 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   6480
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtOutput 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   720
         Width           =   5415
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Output File:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Input File:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   690
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Options"
      Height          =   1695
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   7095
      Begin VB.CheckBox chk 
         Caption         =   "Inject comment"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox chk 
         Caption         =   "Obfuscate symbols starting with ""$"""
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox chk 
         Caption         =   "Safely remove unneeded whitespace and control characters."
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   4815
      End
      Begin VB.CheckBox chk 
         Caption         =   "Strip debug code starting with "";;;"""
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2895
      End
      Begin VB.CheckBox chk 
         Caption         =   "Strip Comments"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmOfuscator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function Ofuscate() As Boolean

    On Error GoTo ErrorOfuscate
    
    Dim Archivo As String
    
    Archivo = util.StripPath(App.Path) & "ofuscator\js_juicer.exe"
    
    If Not ArchivoExiste2(Archivo) Then
        MsgBox "File doesn't found : " & Archivo, vbAbortRetryIgnore
        Exit Function
    End If
    
    util.Hourglass hwnd, True
    
    Dim s As String
    Dim d As String
    Dim M As String
    Dim o As String
    Dim C As String
    
    If chk(0).Value = 1 Then s = "s"
    If chk(1).Value = 1 Then d = "d"
    If chk(2).Value = 1 Then M = "m"
    If chk(3).Value = 1 Then o = "o"
    If chk(4).Value = 1 Then C = "c"
    
    If M <> "" And s = "" Then s = "s"
    
    Dim linea As String
    Dim Arch1 As String
    Dim Arch2 As String
    
    Arch1 = util.VBArchivoSinPath(txtInput.Text)
    Arch2 = util.VBArchivoSinPath(txtOutput.Text)
    
    'copiar archivos al path del jsjuicer
    Call util.CopiarArchivo(txtInput.Text, util.StripPath(App.Path) & "ofuscator\" & Arch1)
            
    Dim ArchBat As String
    
    Dim nFreeFile As Long
    
    nFreeFile = FreeFile
    
    linea = "js_juicer.exe -" & s & d & M & o & C & " " & Arch2 & " " & Arch1
    
    ArchBat = util.StripPath(App.Path) & "ofuscator\ofuscator.bat"
    
    Call util.BorrarArchivo(ArchBat)
    
    Open ArchBat For Output As #nFreeFile
        Print #nFreeFile, "@echo off"
        Print #nFreeFile, "cd " & util.StripPath(App.Path) & "ofuscator\"
        Print #nFreeFile, linea
    Close #nFreeFile
    
    Shell ArchBat, vbHide
    
    Dim ahora
    
    ahora = Now
    
    Do While DateDiff("s", ahora, Now) <= 3
        DoEvents
    Loop
    
    util.Hourglass hwnd, False
    
    Call util.BorrarArchivo(util.StripPath(App.Path) & "ofuscator\" & Arch1)
        
    If ArchivoExiste2(util.StripPath(App.Path) & "ofuscator\" & Arch2) Then
        Call util.CopiarArchivo(util.StripPath(App.Path) & "ofuscator\" & Arch2, txtOutput.Text)
        Call util.BorrarArchivo(util.StripPath(App.Path) & "ofuscator\" & Arch2)
        Ofuscate = True
    Else
        MsgBox "Failed to ofuscate file : " & txtInput.Text
        Ofuscate = False
    End If
    
    Exit Function
ErrorOfuscate:
    util.Hourglass hwnd, False
    MsgBox "Ofuscate : " & Err.Number & " " & Err.description, vbCritical
    
End Function
Private Function Validar() As Boolean

    If txtInput.Text = "" Then
        MsgBox "Select input file to ofuscate", vbCritical
        cmd(2).SetFocus
        Exit Function
    End If

    If txtOutput.Text = "" Then
        MsgBox "Select output file.", vbCritical
        cmd(3).SetFocus
        Exit Function
    End If
    
    Dim k As Integer
    Dim chkok As Boolean
    
    For k = 1 To chk.count - 1
        If chk(k).Value = 1 Then
            chkok = True
            Exit For
        End If
    Next k
    
    If chkok Then
        Validar = True
    Else
        MsgBox "Select an option to ofuscate file.", vbCritical
    End If
    
End Function
Private Sub cmd_Click(Index As Integer)

    Dim glosa As String
    Dim Archivo As String
    
    If Index = 0 Then
        If Validar() Then
            If Ofuscate() Then
                MsgBox "Ready!", vbInformation
            End If
        End If
    ElseIf Index = 2 Then
        LastPath = App.Path
    
        glosa = strGlosa()
        
        If Not Cdlg.VBGetOpenFileName(Archivo, , , , , , glosa, , , "Input File ...") Then
            Exit Sub
        End If
        
        txtInput.Text = Archivo
    ElseIf Index = 3 Then
    
        glosa = strGlosa()
        
        If Not Cdlg.VBGetSaveFileName(Archivo, , , glosa, , LastPath, "Output File ...", "js", Me.hwnd) Then
            Exit Sub
        End If
        
        txtOutput.Text = Archivo
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    util.CenterForm Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOfuscator = Nothing
End Sub
