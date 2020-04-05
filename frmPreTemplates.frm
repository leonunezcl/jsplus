VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Begin VB.Form frmPreTemplates 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Predefined Template"
   ClientHeight    =   6165
   ClientLeft      =   1365
   ClientTop       =   1890
   ClientWidth     =   10260
   Icon            =   "frmPreTemplates.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   4530
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2835
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lst 
      Appearance      =   0  'Flat
      Height          =   5295
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   270
      Width           =   3660
   End
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   8160
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6645
      Visible         =   0   'False
      Width           =   2250
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   1
      Left            =   4575
      TabIndex        =   1
      Top             =   5640
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
      Index           =   0
      Left            =   2580
      TabIndex        =   2
      Top             =   5640
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
   Begin CodeSenseCtl.CodeSense txtCode 
      Height          =   5310
      Left            =   3735
      OleObjectBlob   =   "frmPreTemplates.frx":000C
      TabIndex        =   3
      Top             =   270
      Width           =   6495
   End
   Begin jsplus.MyButton cmd 
      Height          =   405
      Index           =   2
      Left            =   6510
      TabIndex        =   7
      Top             =   5640
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Copy"
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
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Code Preview"
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
      Index           =   0
      Left            =   3735
      TabIndex        =   6
      Top             =   60
      Width           =   1185
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
      TabIndex        =   5
      Top             =   75
      Width           =   3555
   End
End
Attribute VB_Name = "frmPreTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CargaTemplates()

    Dim k As Integer
    
    File1.Path = util.StripPath(App.Path) & "templates"
        
    For k = 0 To File1.ListCount - 1
        lst.AddItem Explode(File1.List(k), 1, ".")
    Next k
    
End Sub
Private Sub InsertTemplate()

    'Dim k As Integer
    Dim ini As String
    Dim glosa As String
    Dim Archivo As String
    Dim nfreefile As Long
    
    util.Hourglass hwnd, True
    
    ini = IniPath
    
    If Not util.ArchivoExiste(ini) Then
        MsgBox "I can't find : " & ini, vbCritical
        Exit Sub
    End If
    
    Archivo = util.StripPath(App.Path) & "templates\" & lst.Text & ".txt"

    If Not util.ArchivoExiste(Archivo) Then
        MsgBox "The template file " & Archivo & " there is doesn't exists.", vbCritical
        Exit Sub
    End If
    
    On Error Resume Next
    nfreefile = FreeFile
    Open Archivo For Input As #nfreefile
        glosa = Input(LOF(nfreefile), nfreefile)
    Close #1
    Err = 0
    
    Call frmMain.ActiveForm.Insertar(glosa)
    
    util.Hourglass hwnd, False
    
End Sub
Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If lst.ListIndex <> -1 Then
            Call InsertTemplate
            Unload Me
        Else
            MsgBox "Select template to insert", vbCritical
        End If
    ElseIf Index = 2 Then
        If txtCode.CanCopy Then
            txtCode.Copy
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

Private Sub CargarConfiguracion()

    'bookmark
    Call txtCode.SetColor(cmClrBookmark, frmMain.ActiveForm.txtCode.GetColor(cmClrBookmark))
    Call txtCode.SetColor(cmClrBookmarkBk, frmMain.ActiveForm.txtCode.GetColor(cmClrBookmarkBk))
            
    'resalte de linea
    Call txtCode.SetColor(cmClrHighlightedLine, frmMain.ActiveForm.txtCode.GetColor(cmClrHighlightedLine))
    
    'lineas divisoras
    Call txtCode.SetColor(cmClrHDividerLines, frmMain.ActiveForm.txtCode.GetColor(cmClrHDividerLines))
    Call txtCode.SetColor(cmClrVDividerLines, frmMain.ActiveForm.txtCode.GetColor(cmClrVDividerLines))
   
    'comentarios
    Call txtCode.SetColor(cmClrComment, frmMain.ActiveForm.txtCode.GetColor(cmClrComment))
    Call txtCode.SetColor(cmClrCommentBk, frmMain.ActiveForm.txtCode.GetColor(cmClrCommentBk))
    Call txtCode.SetFontStyle(cmStyComment, frmMain.ActiveForm.txtCode.GetFontStyle(cmStyComment))
        
    'keywords
    Call txtCode.SetColor(cmClrKeyword, frmMain.ActiveForm.txtCode.GetColor(cmClrKeyword))
    Call txtCode.SetColor(cmClrKeywordBk, frmMain.ActiveForm.txtCode.GetColor(cmClrKeywordBk))
    Call txtCode.SetFontStyle(cmStyKeyword, frmMain.ActiveForm.txtCode.GetFontStyle(cmStyKeyword))
            
    'numeros de linea
    Call txtCode.SetColor(cmClrLineNumber, frmMain.ActiveForm.txtCode.GetColor(cmClrLineNumber))
    Call txtCode.SetColor(cmClrLineNumberBk, frmMain.ActiveForm.txtCode.GetColor(cmClrLineNumberBk))
    Call txtCode.SetFontStyle(cmStyLineNumber, frmMain.ActiveForm.txtCode.GetFontStyle(cmStyLineNumber))
            
    'numeros
    Call txtCode.SetColor(cmClrNumber, frmMain.ActiveForm.txtCode.GetColor(cmClrNumber))
    Call txtCode.SetColor(cmClrNumberBk, frmMain.ActiveForm.txtCode.GetColor(cmClrNumberBk))
    Call txtCode.SetFontStyle(cmStyNumber, frmMain.ActiveForm.txtCode.GetFontStyle(cmStyNumber))
    
    'operadores
    Call txtCode.SetColor(cmClrOperator, frmMain.ActiveForm.txtCode.GetColor(cmClrOperator))
    Call txtCode.SetColor(cmClrOperatorBk, frmMain.ActiveForm.txtCode.GetColor(cmClrOperatorBk))
    Call txtCode.SetFontStyle(cmStyOperator, frmMain.ActiveForm.txtCode.GetFontStyle(cmStyOperator))
    
    'alcance (scope)
    Call txtCode.SetColor(cmClrScopeKeyword, frmMain.ActiveForm.txtCode.GetColor(cmClrScopeKeyword))
    Call txtCode.SetColor(cmClrScopeKeywordBk, frmMain.ActiveForm.txtCode.GetColor(cmClrScopeKeywordBk))
    Call txtCode.SetFontStyle(cmStyScopeKeyword, frmMain.ActiveForm.txtCode.GetFontStyle(cmStyScopeKeyword))
    
    'cadenas
    Call txtCode.SetColor(cmClrString, frmMain.ActiveForm.txtCode.GetColor(cmClrString))
    Call txtCode.SetColor(cmClrStringBk, frmMain.ActiveForm.txtCode.GetColor(cmClrStringBk))
    Call txtCode.SetFontStyle(cmStyString, frmMain.ActiveForm.txtCode.GetFontStyle(cmStyString))
    
    'texto
    Call txtCode.SetColor(cmClrText, frmMain.ActiveForm.txtCode.GetColor(cmClrText))
    Call txtCode.SetColor(cmClrTextBk, frmMain.ActiveForm.txtCode.GetColor(cmClrTextBk))
    Call txtCode.SetFontStyle(cmStyText, frmMain.ActiveForm.txtCode.GetFontStyle(cmStyText))
    
    'fondo de la ventana
    Call txtCode.SetColor(cmClrWindow, frmMain.ActiveForm.txtCode.GetColor(cmClrWindow))
    
    'otros
    txtCode.Font.Name = frmMain.ActiveForm.txtCode.Font.Name
    txtCode.Font.Size = frmMain.ActiveForm.txtCode.Font.Size
    txtCode.Font.Bold = frmMain.ActiveForm.txtCode.Font.Bold
    txtCode.Font.Italic = frmMain.ActiveForm.txtCode.Font.Italic
    txtCode.TabSize = frmMain.ActiveForm.txtCode.TabSize
    
End Sub


Private Sub Form_Load()

    util.CenterForm Me
    set_color_form Me
    util.Hourglass hwnd, True
    CargaTemplates
    SetearLenguaje
    util.Hourglass hwnd, False
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    cmd(0).Refresh
    cmd(1).Refresh
    cmd(2).Refresh
    'SetLayered hwnd, True
    
    Debug.Print "load"
    DrawXPCtl Me
    
End Sub

Private Sub SetearLenguaje()

    If frmMain.ActiveForm.Name = "frmEdit" Then
        txtCode.Language = frmMain.ActiveForm.txtCode.Language
    End If
    
    Call CargarConfiguracion
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmPreTemplates = Nothing
End Sub


Private Sub lst_Click()

    On Error Resume Next
    
    Dim Archivo As String
    Dim k As Integer
    
    For k = 0 To File1.ListCount - 1
        If Left$(File1.List(k), Len(lst.Text)) = lst.Text Then
            Archivo = util.StripPath(App.Path) & "templates\" & File1.List(k)
            txtCode.OpenFile (Archivo)
        End If
    Next k
    
    Err = 0
    
End Sub

Private Sub lst_DblClick()
    cmd_Click 0
End Sub


Private Function txtCode_RClick(ByVal Control As CodeSenseCtl.ICodeSense) As Boolean
    txtCode_RClick = True
End Function


