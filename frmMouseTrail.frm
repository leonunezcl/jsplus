VERSION 5.00
Begin VB.Form frmMouseTrail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MouseTrail"
   ClientHeight    =   5385
   ClientLeft      =   5175
   ClientTop       =   2970
   ClientWidth     =   4380
   Icon            =   "frmMouseTrail.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   292
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MyButtonDefSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   8325
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3465
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picImg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   45
      ScaleHeight     =   1785
      ScaleWidth      =   2865
      TabIndex        =   2
      Top             =   3390
      Width           =   2895
   End
   Begin VB.ListBox lstImg 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   45
      TabIndex        =   1
      Top             =   270
      Width           =   2895
   End
   Begin jsplus.MyButton cmdAgregar 
      Height          =   375
      Left            =   3060
      TabIndex        =   4
      Top             =   270
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
   Begin jsplus.MyButton cmdEliminar 
      Height          =   375
      Left            =   3060
      TabIndex        =   5
      Top             =   735
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Remove"
      AccessKey       =   "R"
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
   Begin jsplus.MyButton cmdGenerar 
      Height          =   375
      Left            =   3060
      TabIndex        =   6
      Top             =   1230
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      SPN             =   "MyButtonDefSkin"
      Text            =   "A&pply"
      AccessKey       =   "p"
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
   Begin jsplus.MyButton cmdPreview 
      Height          =   375
      Left            =   3060
      TabIndex        =   7
      Top             =   1665
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      SPN             =   "MyButtonDefSkin"
      Text            =   "P&review"
      AccessKey       =   "r"
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
   Begin jsplus.MyButton cmdSalir 
      Height          =   375
      Left            =   3060
      TabIndex        =   8
      Top             =   2130
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
   Begin VB.Image imgTmp 
      Height          =   495
      Left            =   3105
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Preview"
      Height          =   195
      Index           =   1
      Left            =   45
      TabIndex        =   3
      Top             =   3150
      Width           =   570
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Images"
      Height          =   195
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   510
   End
End
Attribute VB_Name = "frmMouseTrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function add_image() As String

    Dim ret As String
    
    If Not Cdlg.VBGetOpenFileName(ret, , , , , , , , App.Path, "Open Image File") Then
        Exit Function
    End If
    
    add_image = ret
    
End Function

Private Sub preview_mousetrail(ByVal mostrar As Boolean)

    Dim buffer As New cStringBuilder
    Dim archivo As String
    Dim k As Integer
    Dim nFreeFile As Long
    Dim pathapp As String
    Dim arr_data() As String
    Dim linea As String
    Dim pathdes As String
    Dim glosa As String
    Dim imgfile As String
    
    pathapp = Util.StripPath(App.Path) & "plus\mousetrail\"
    
    glosa = "Hypertext files (*.htm)|*.htm|"
    glosa = glosa & "All Files (*.*)|*.*"
        
    k = 1
    nFreeFile = FreeFile
    If lstImg.ListCount - 1 <> -1 Then
        If Not Util.ArchivoExiste(pathapp & "header.js") Then
            MsgBox pathapp & "header.js doesn't exists", vbCritical
            Exit Sub
        End If
        
        If Not Util.ArchivoExiste(pathapp & "footer.js") Then
            MsgBox pathapp & "footer.js doesn't exists", vbCritical
            Exit Sub
        End If
        
        If mostrar Then
            archivo = Util.StripPath(App.Path) & "mousetrail.html"
        Else
            If Not Cdlg.VBGetSaveFileName(archivo, , , glosa, , App.Path, "Save As ...", "htm") Then
                MsgBox "Canceled by user", vbCritical
                Exit Sub
            End If
        End If
            
        'copiar las imagenes al path destino del archivo
        pathdes = Util.StripPath(Util.PathArchivo(archivo))
        For k = 0 To lstImg.ListCount - 1
            Util.CopiarArchivo lstImg.List(k), pathdes & Util.VBArchivoSinPath(lstImg.List(k))
        Next k
        
        Open pathapp & "header.js" For Input As #nFreeFile
            Do While Not EOF(nFreeFile)
                Line Input #nFreeFile, linea
                ReDim Preserve arr_data(k)
                arr_data(k) = linea
                k = k + 1
            Loop
        Close #nFreeFile
        
        nFreeFile = FreeFile

        Open archivo For Output As #nFreeFile
            Print #nFreeFile, "<html>"
            Print #nFreeFile, "<body>"
            
            For k = 1 To UBound(arr_data)
                Print #nFreeFile, arr_data(k)
            Next k
            
            buffer.Append "T1=new Array("

            For k = 0 To lstImg.ListCount - 1
                imgfile = Util.VBArchivoSinPath(lstImg.List(k))
                Set imgTmp.Picture = LoadPicture(imgfile)
                DoEvents
                If k < lstImg.ListCount - 1 Then
                    buffer.Append Chr$(34) & imgfile & Chr$(34) & "," & imgTmp.Height & "," & imgTmp.Width & ","
                Else
                    buffer.Append Chr$(34) & imgfile & Chr$(34) & "," & imgTmp.Height & "," & imgTmp.Width
                End If
            Next k
            buffer.Append vbNewLine
            buffer.Append ")" & vbNewLine

            Print #nFreeFile, buffer.ToString
        Close #nFreeFile
    
        ReDim arr_data(0)
                        
        k = 1
        nFreeFile = FreeFile
        
        Open pathapp & "footer.js" For Input As #nFreeFile
            Do While Not EOF(nFreeFile)
                Line Input #nFreeFile, linea
                ReDim Preserve arr_data(k)
                arr_data(k) = linea
                k = k + 1
            Loop
        Close #nFreeFile
        
        nFreeFile = FreeFile

        Open archivo For Append As #nFreeFile
            For k = 1 To UBound(arr_data)
                Print #nFreeFile, arr_data(k)
            Next k
            Print #nFreeFile, "</body>"
            Print #nFreeFile, "</html>"
        Close #nFreeFile
        
        Util.ShellFunc archivo, vbNormalFocus
    Else
        MsgBox "Nothing to do", vbCritical
    End If
    
    Set buffer = Nothing
    
End Sub
Private Sub cmdAgregar_Click()

    Dim archivo As String
    
    archivo = add_image()
    
    If Len(archivo) > 0 Then
        lstImg.AddItem archivo
    End If
    
End Sub

Private Sub cmdEliminar_Click()

    If lstImg.ListIndex <> -1 Then
        lstImg.RemoveItem lstImg.ListIndex
    End If
    
End Sub

Private Sub cmdGenerar_Click()
    Call preview_mousetrail(False)
End Sub

Private Sub cmdPreview_Click()
    preview_mousetrail True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Load()

    Util.CenterForm Me
    
    Set MyButtonDefSkin.Picture = LoadResPicture(1002, vbResBitmap)
    cmdAgregar.Refresh
    cmdEliminar.Refresh
    cmdGenerar.Refresh
    cmdPreview.Refresh
    cmdSalir.Refresh
    
    DrawXPCtl Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMouseTrail = Nothing
End Sub


Private Sub lstImg_Click()

    If lstImg.ListIndex <> -1 Then
        Set picImg.Picture = LoadPicture(lstImg.List(lstImg.ListIndex))
    End If
    
End Sub


