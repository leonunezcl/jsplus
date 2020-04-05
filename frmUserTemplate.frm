VERSION 5.00
Begin VB.Form frmUserTemplate 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert User Template"
   ClientHeight    =   3870
   ClientLeft      =   3555
   ClientTop       =   1575
   ClientWidth     =   5400
   Icon            =   "frmUserTemplate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   3690
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5250
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.ListBox lstjs 
      Appearance      =   0  'Flat
      Height          =   3540
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   3780
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   0
      Left            =   3900
      TabIndex        =   2
      Top             =   105
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Ok"
      AccessKey       =   "O"
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
      Left            =   3900
      TabIndex        =   3
      Top             =   585
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Exit"
      AccessKey       =   "E"
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
      Left            =   3900
      TabIndex        =   4
      Top             =   1065
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Add"
      AccessKey       =   "A"
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
      Index           =   3
      Left            =   3900
      TabIndex        =   5
      Top             =   1545
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Modify"
      AccessKey       =   "M"
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
      Index           =   4
      Left            =   3900
      TabIndex        =   6
      Top             =   2010
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Delete"
      AccessKey       =   "D"
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
      Caption         =   "Double click to insert "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   45
      TabIndex        =   7
      Top             =   3645
      Width           =   3555
   End
End
Attribute VB_Name = "frmUserTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub CargarTemplates()

    On Error GoTo ErrorCargarTemplates
    
    Dim k
    Dim j As Integer
    Dim plantilla As String
    Dim Archivo As String
    Dim ini As String
    Dim Path As String
    Dim valido As String
    
    util.Hourglass hwnd, True
    
    lstjs.Clear
    
    ini = IniPath
    Path = util.StripPath(App.Path) & "utemplates\"
    
    k = util.LeeIni(ini, "plantillas", "cantidad")
    
    If k = "" Or Not IsNumeric(k) Then
        util.Hourglass hwnd, False
        Exit Sub
    End If
    
    For j = 1 To CInt(k)
        plantilla = util.LeeIni(ini, "plantillas", "utemplate" & j)
        If Len(plantilla) > 0 Then
            If InStr(plantilla, ";") > 0 Then
                valido = Explode(plantilla, 3, ";")
                If valido = "y" Then
                    Archivo = Explode(plantilla, 2, ";")
                    Archivo = Path & Archivo
                    If util.ArchivoExiste(Archivo) Then
                        lstjs.AddItem Explode(plantilla, 1, ";")
                        lstjs.ItemData(lstjs.NewIndex) = j
                    End If
                End If
            End If
        End If
    Next j
    
    util.Hourglass hwnd, False
    
    Exit Sub
ErrorCargarTemplates:
    lstjs.Clear
    MsgBox "CargarTemplates : " & Err & " " & Error$, vbCritical
    Err = 0
    
End Sub

Private Sub InsertarCodigo()

    On Error GoTo ErrorInsertarCodigo
    
    Dim texto As String
    Dim Index As Integer
    Dim src As New cStringBuilder
    Dim valido As String
    
    Index = lstjs.ListIndex
    
    If Index = -1 Then Exit Sub
        
    Dim plantilla As String
    Dim Archivo As String
    Dim ini As String
    Dim Path As String
    Dim nfreefile As Long
    Dim k As Variant
    ini = IniPath
    Path = util.StripPath(App.Path) & "utemplates\"
    
    k = util.LeeIni(ini, "plantillas", "cantidad")
    
    If k = "" Or Not IsNumeric(k) Then
        MsgBox "Error at try to load template from : " & ini, vbCritical
        Exit Sub
    End If
    
    plantilla = util.LeeIni(ini, "plantillas", "utemplate" & lstjs.ItemData(lstjs.ListIndex))
    
    If Len(plantilla) > 0 Then
        If InStr(plantilla, ";") > 0 Then
            valido = Explode(plantilla, 3, ";")
            If valido = "y" Then
                Archivo = Explode(plantilla, 2, ";")
                Archivo = Path & Archivo
                If util.ArchivoExiste(Archivo) Then
                    nfreefile = FreeFile
                    Open Archivo For Input As #nfreefile
                        src.Append Input(LOF(nfreefile), nfreefile)
                    Close #nfreefile
                End If
            End If
        End If
    End If
    
    If frmMain.ActiveForm.Name = "frmEdit" Then
        Call frmMain.ActiveForm.Insertar(src.ToString)
    End If
    
    Exit Sub
ErrorInsertarCodigo:
    MsgBox "InsertarCodigo : " & Err & " " & Error$, vbCritical
    
End Sub

Private Sub cmd_Click(Index As Integer)

    On Error Resume Next
    
    Dim plantilla As String
    Dim Archivo As String
    
    If Index = 0 Then
        Call InsertarCodigo
        Unload Me
    ElseIf Index = 1 Then
        Unload Me
    ElseIf Index = 2 Then
        frmCrearPlantilla.Show vbModal
    ElseIf Index = 3 Then
        'modificar
        If lstjs.ListCount - 1 > -1 Then
            If lstjs.ListIndex <> -1 Then
                plantilla = util.LeeIni(IniPath, "plantillas", "utemplate" & lstjs.ItemData(lstjs.ListIndex))
                Archivo = Explode(plantilla, 2, ";")
                frmCrearPlantilla.template_name = lstjs.Text
                frmCrearPlantilla.template_file = Archivo
                frmCrearPlantilla.Show vbModal
            End If
        End If
    ElseIf Index = 4 Then
        If lstjs.ListCount - 1 > -1 Then
            If lstjs.ListIndex <> -1 Then
                If Confirma("Are you sure to remove this template") = vbYes Then
                    plantilla = util.LeeIni(IniPath, "plantillas", "utemplate" & lstjs.ItemData(lstjs.ListIndex))
                    Archivo = Explode(plantilla, 2, ";")
                    util.BorrarArchivo util.StripPath(App.Path) & "utemplates\" & Archivo
                    Call util.GrabaIni(IniPath, "plantillas", "utemplate" & lstjs.ItemData(lstjs.ListIndex), Replace(plantilla, ";y", ";n"))
                    lstjs.RemoveItem lstjs.ListIndex
                End If
            End If
        End If
    End If
    
    Err = 0
    
End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    util.CenterForm Me
    
    set_color_form Me
    
    Call CargarTemplates
    
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    cmd(0).Refresh
    cmd(1).Refresh
    cmd(2).Refresh
    cmd(3).Refresh
    cmd(4).Refresh
    
    Debug.Print "load"
    
    DrawXPCtl Me
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmUserTemplate = Nothing
End Sub


Private Sub lstjs_DblClick()
    InsertarCodigo
End Sub

