VERSION 5.00
Begin VB.Form frmCalendar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar Wizard"
   ClientHeight    =   4590
   ClientLeft      =   3495
   ClientTop       =   1965
   ClientWidth     =   5085
   Icon            =   "frmCalendar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   15
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Preview"
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Include files :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   75
      TabIndex        =   10
      Top             =   2685
      Width           =   4935
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "calendarus.js, calendar.gif, close.gif, drop1.gif, drop2.gif, left1.gif, left2.gif,right1.gif, right2.gif"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   510
         Index           =   6
         Left            =   150
         TabIndex        =   11
         Tag             =   "calendar.gif, close.gif, drop1.gif, drop2.gif, left1.gif, left2.gif,right1.gif, right2.gif"
         Top             =   300
         Width           =   4605
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Language"
      Height          =   825
      Index           =   1
      Left            =   75
      TabIndex        =   9
      Top             =   1800
      Width           =   4935
      Begin VB.OptionButton optLan 
         Caption         =   "Spanish"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   4
         Top             =   330
         Width           =   1005
      End
      Begin VB.OptionButton optLan 
         Caption         =   "English"
         Height          =   255
         Index           =   0
         Left            =   1005
         TabIndex        =   3
         Top             =   330
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Form Settings"
      Height          =   1680
      Index           =   0
      Left            =   75
      TabIndex        =   5
      Top             =   60
      Width           =   4935
      Begin VB.TextBox txtDatePrompt 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   3435
      End
      Begin VB.TextBox txtForm 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   3435
      End
      Begin VB.TextBox txtDateName 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   3435
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Prompt"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   8
         Top             =   630
         Width           =   885
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Form Name"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   285
         Width           =   810
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Name"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   975
         Width           =   810
      End
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Browser Compatibility"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   2145
      TabIndex        =   12
      Top             =   4260
      Width           =   1485
   End
   Begin VB.Image imgIE 
      Height          =   255
      Left            =   3750
      Top             =   4245
      Width           =   300
   End
   Begin VB.Image imgFX 
      Height          =   255
      Left            =   4065
      Top             =   4245
      Width           =   300
   End
   Begin VB.Image imgNE 
      Height          =   255
      Left            =   4380
      Top             =   4245
      Width           =   300
   End
   Begin VB.Image imgOP 
      Height          =   255
      Left            =   4710
      Top             =   4245
      Width           =   300
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ultimo_path As String
Private Sub CrearCalendario(ByVal preview As Boolean)

    Dim buffer As New cStringBuilder
    Dim Archivo As String
    Dim nFreeFile As Long
    Dim glosa As String
    Dim pathapp As String
    Dim pathwrk As String
    Dim FormName As String
    Dim DatePrompt As String
    Dim FechaName As String
    
    FormName = txtForm.Text
    If Trim$(FormName) = "" Then
        FormName = "frmDate"
    End If
    
    DatePrompt = txtDatePrompt.Text
    If Trim$(DatePrompt) = "" Then
        DatePrompt = "Select Date :"
    End If
    
    FechaName = txtDateName.Text
    If Trim$(FechaName) = "" Then
        FechaName = "txtFecha"
    End If
    
    buffer.Append "<html>" & vbNewLine
    buffer.Append "<head>" & vbNewLine
    
    If optLan(0).Value Then
        buffer.Append "<script language=" & Chr$(34) & "Javascript" & Chr$(34) & " type=""text/javascript"" src=""calendarus.js""></script>" & vbNewLine
    Else
        buffer.Append "<script language=" & Chr$(34) & "Javascript" & Chr$(34) & " type=""text/javascript"" src=""calendares.js""></script>" & vbNewLine
    End If
    
    buffer.Append "</head>" & vbNewLine
    buffer.Append "<body>" & vbNewLine
    buffer.Append "<form name=" & Chr$(34) & FormName & Chr$(34) & ">" & vbNewLine
    
    buffer.Append "<input type=text name=" & FechaName & " size=10 maxlength=10 disabled=true>" & vbNewLine
    
    If optLan(0).Value Then
        buffer.Append "<img src=" & Chr$(34) & "img/calendar.gif" & Chr$(34) & " onClick=" & Chr$(34) & "popUpCalendar(this, document." & FormName & "." & FechaName & ",'mm/dd/yyyy');return false;" & Chr$(34) & " alt='Select Date'>" & vbNewLine
    Else
        buffer.Append "<img src=" & Chr$(34) & "img/calendar.gif" & Chr$(34) & " onClick=" & Chr$(34) & "popUpCalendar(this, document." & FormName & "." & FechaName & ",'dd/mm/yyyy');return false;" & Chr$(34) & " alt='Selecciona Fecha'>" & vbNewLine
    End If
    
    buffer.Append "</form>" & vbNewLine
    buffer.Append "</body>" & vbNewLine
    buffer.Append "</html>" & vbNewLine
    
    nFreeFile = FreeFile
        
    pathapp = util.StripPath(App.Path) & "plus\calendar\"
    pathwrk = util.StripPath(App.Path)
    
    If preview Then
        Archivo = pathapp & "calendar.htm"
        
        Open Archivo For Output As #nFreeFile
            Print #nFreeFile, buffer.ToString
        Close #nFreeFile
            
        util.ShellFunc Archivo, vbNormalFocus
    Else
        glosa = "Hypertext files (*.htm)|*.htm|"
        glosa = glosa & "All Files (*.*)|*.*"
    
        If ultimo_path = "" Then
            ultimo_path = App.Path
        End If
            
        If Cdlg.VBGetSaveFileName(Archivo, , , glosa, , ultimo_path, "Save File As ...", "htm") Then
            ultimo_path = util.PathArchivo(Archivo)
                
            If ultimo_path <> pathapp Then
                Open Archivo For Output As #nFreeFile
                    Print #nFreeFile, buffer.ToString
                Close #nFreeFile
                
                'copiar los archivos necesarios para generar esto
                If optLan(0).Value Then
                    util.CopiarArchivo pathapp & "calendarus.js", ultimo_path & "calendarus.js"
                Else
                    util.CopiarArchivo pathapp & "calendares.js", ultimo_path & "calendares.js"
                End If
                
                util.CrearDirectorio (ultimo_path & "img")
                util.CopiarArchivo pathapp & "img\close.gif", ultimo_path & "img\close.gif"
                util.CopiarArchivo pathapp & "img\calendar.gif", ultimo_path & "img\calendar.gif"
                util.CopiarArchivo pathapp & "img\right1.gif", ultimo_path & "img\right1.gif"
                util.CopiarArchivo pathapp & "img\right2.gif", ultimo_path & "img\right2.gif"
                util.CopiarArchivo pathapp & "img\drop1.gif", ultimo_path & "img\drop1.gif"
                util.CopiarArchivo pathapp & "img\drop2.gif", ultimo_path & "img\drop2.gif"
                util.CopiarArchivo pathapp & "img\left1.gif", ultimo_path & "img\left1.gif"
                util.CopiarArchivo pathapp & "img\left2.gif", ultimo_path & "img\left2.gif"
                                
                util.ShellFunc Archivo, vbNormalFocus
            Else
                MsgBox "Invalid path. You must choice another path", vbCritical
            End If
        End If
    End If
    
    Set buffer = Nothing
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call CrearCalendario(False)
    ElseIf Index = 2 Then
        Call CrearCalendario(True)
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
    
    Set imgIE.Picture = LoadResPicture(1007, vbResBitmap)
    Set imgFX.Picture = LoadResPicture(1008, vbResBitmap)
    Set imgNE.Picture = LoadResPicture(1009, vbResBitmap)
    Set imgOP.Picture = LoadResPicture(1010, vbResBitmap)
    
    Debug.Print "load"
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call clear_memory(Me)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmCalendar = Nothing
End Sub


Private Sub optLan_Click(Index As Integer)

    If Index = 0 Then
        lbl(6).Caption = "calendarus.js, " & lbl(6).Tag
    Else
        lbl(6).Caption = "calendares.js, " & lbl(6).Tag
    End If
    
End Sub


