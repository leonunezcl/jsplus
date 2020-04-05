VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSlideShow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SlideShow Wizard"
   ClientHeight    =   9000
   ClientLeft      =   4020
   ClientTop       =   1995
   ClientWidth     =   7935
   ControlBox      =   0   'False
   Icon            =   "frmSlideShow.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "Remove All"
      Height          =   375
      Index           =   5
      Left            =   6600
      TabIndex        =   39
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Remove Image"
      Height          =   375
      Index           =   3
      Left            =   6600
      TabIndex        =   38
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Add Folder"
      Height          =   375
      Index           =   4
      Left            =   6600
      TabIndex        =   37
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Update Image"
      Height          =   375
      Index           =   7
      Left            =   6600
      TabIndex        =   36
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Add Image"
      Height          =   375
      Index           =   2
      Left            =   6600
      TabIndex        =   35
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Preview"
      Height          =   375
      Index           =   6
      Left            =   6600
      TabIndex        =   34
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Exit"
      Height          =   375
      Index           =   1
      Left            =   6600
      TabIndex        =   33
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   6600
      TabIndex        =   32
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Images Settings"
      Height          =   1170
      Index           =   2
      Left            =   30
      TabIndex        =   22
      Top             =   2055
      Width           =   6405
      Begin VB.ComboBox cboTransitions 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   705
         Width           =   4560
      End
      Begin VB.ComboBox cboPrefetch 
         Height          =   315
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   300
         Width           =   1470
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Transitions (IE Only)"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   24
         Top             =   735
         Width           =   1410
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Prefetch images :"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   23
         Top             =   375
         Width           =   1230
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Slide Controls"
      Height          =   795
      Index           =   3
      Left            =   30
      TabIndex        =   21
      Top             =   1215
      Width           =   6405
      Begin VB.CheckBox chkControls 
         Caption         =   "Stop"
         Height          =   255
         Index           =   5
         Left            =   5460
         TabIndex        =   9
         Top             =   360
         Value           =   1  'Checked
         Width           =   720
      End
      Begin VB.CheckBox chkControls 
         Caption         =   "Play"
         Height          =   255
         Index           =   4
         Left            =   4395
         TabIndex        =   8
         Top             =   360
         Value           =   1  'Checked
         Width           =   750
      End
      Begin VB.CheckBox chkControls 
         Caption         =   "Random"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   7
         Top             =   360
         Value           =   1  'Checked
         Width           =   1005
      End
      Begin VB.CheckBox chkControls 
         Caption         =   "View"
         Height          =   255
         Index           =   2
         Left            =   2250
         TabIndex        =   6
         Top             =   360
         Value           =   1  'Checked
         Width           =   825
      End
      Begin VB.CheckBox chkControls 
         Caption         =   "Next"
         Height          =   255
         Index           =   1
         Left            =   1365
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chkControls 
         Caption         =   "Previous"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   1005
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Slide Configuration"
      Height          =   1065
      Index           =   0
      Left            =   30
      TabIndex        =   18
      Top             =   60
      Width           =   6405
      Begin VB.CheckBox chkSuffle 
         Caption         =   "Shuffle/randomize the slideshow"
         Height          =   255
         Left            =   3345
         TabIndex        =   3
         Top             =   645
         Width           =   2730
      End
      Begin VB.CheckBox chkRepeat 
         Caption         =   "Repeat or loop the slideshow"
         Height          =   255
         Left            =   135
         TabIndex        =   1
         Top             =   675
         Width           =   2535
      End
      Begin VB.CheckBox chkAutoStart 
         Caption         =   "Automatically start playing"
         Height          =   255
         Left            =   3345
         TabIndex        =   2
         Top             =   315
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.TextBox txtTimeout 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1290
         MaxLength       =   255
         TabIndex        =   0
         Text            =   "3000"
         Top             =   270
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "1000 = 1 Second"
         Height          =   195
         Index           =   3
         Left            =   1935
         TabIndex        =   20
         Top             =   315
         Width           =   1230
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Slide duration:"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   19
         Top             =   270
         Width           =   1005
      End
   End
   Begin MSComctlLib.ListView lvwImages 
      Height          =   2340
      Left            =   30
      TabIndex        =   16
      Top             =   6240
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   4128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "FileName"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Link"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Target"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Attributes"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fra 
      Caption         =   "Slide Settings"
      Height          =   2670
      Index           =   1
      Left            =   30
      TabIndex        =   25
      Top             =   3270
      Width           =   6405
      Begin VB.TextBox txtAttributes 
         Height          =   285
         Left            =   135
         MaxLength       =   255
         TabIndex        =   15
         Top             =   2220
         Width           =   6075
      End
      Begin VB.TextBox txtTarget 
         Height          =   285
         Left            =   135
         MaxLength       =   255
         TabIndex        =   14
         Top             =   1650
         Width           =   6075
      End
      Begin VB.TextBox txtLink 
         Height          =   285
         Left            =   135
         MaxLength       =   255
         TabIndex        =   13
         Top             =   1065
         Width           =   6075
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   135
         MaxLength       =   255
         TabIndex        =   12
         Top             =   480
         Width           =   6075
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Required"
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
         Height          =   195
         Index           =   0
         Left            =   5445
         TabIndex        =   30
         Top             =   255
         Width           =   780
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Window attributes : (width:600,height:400)"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   29
         Top             =   1995
         Width           =   2985
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Link target : (windowname, _blank)"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   28
         Top             =   1410
         Width           =   2475
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Link: (http://www.mywebsite.com/)"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   27
         Top             =   840
         Width           =   2505
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Image filename : (pic1.jpg)"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   26
         Top             =   270
         Width           =   1845
      End
   End
   Begin VB.Image imgOP 
      Height          =   255
      Left            =   7755
      Top             =   8715
      Width           =   300
   End
   Begin VB.Image imgNE 
      Height          =   255
      Left            =   7425
      Top             =   8715
      Width           =   300
   End
   Begin VB.Image imgIE 
      Height          =   255
      Left            =   7095
      Top             =   8715
      Width           =   300
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Browser Compatibility"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   5490
      TabIndex        =   31
      Top             =   8730
      Width           =   1485
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Images"
      Height          =   195
      Index           =   1
      Left            =   30
      TabIndex        =   17
      Top             =   6015
      Width           =   510
   End
End
Attribute VB_Name = "frmSlideShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private arr_filters() As String
Private idx As Integer
Private Sub add_folder()

    Dim Path As String
    Dim arr_files() As String
    Dim k As Integer
    
    Path = util.BrowseFolder(hwnd)
    
    If Len(Path) > 0 Then
        get_files_from_folder Path, arr_files
        For k = 1 To UBound(arr_files)
            idx = idx + 1
            lvwImages.ListItems.Add , "k" & idx, util.VBArchivoSinPath(arr_files(k))
            lvwImages.ListItems("k" & idx).SubItems(1) = vbNullString
            lvwImages.ListItems("k" & idx).SubItems(2) = vbNullString
            lvwImages.ListItems("k" & idx).SubItems(3) = vbNullString
        Next k
    End If
    
End Sub

Private Sub add_image()
    
    If txtFileName.Text = "" Then
        txtFileName.SetFocus
        Exit Sub
    End If
    
    idx = idx + 1
    lvwImages.ListItems.Add , "k" & idx, txtFileName.Text
    lvwImages.ListItems("k" & idx).SubItems(1) = txtLink.Text
    lvwImages.ListItems("k" & idx).SubItems(2) = txtTarget.Text
    lvwImages.ListItems("k" & idx).SubItems(3) = txtAttributes.Text
        
End Sub

Private Function GeneraSlideShow(ByVal preview As Boolean) As Boolean

    Dim k As Integer
    Dim buffer As New cStringBuilder
    Dim buffer2 As New cStringBuilder
    Dim buffer3 As New cStringBuilder
    Dim nFreeFile As Long
    Dim Archivo As String
    Dim linea As String
    
    util.Hourglass hwnd, True
    
    'cargar lineas del archivo fuente
    Archivo = util.StripPath(App.Path) & "plus\slideshow\slideshow.js"
    If Not ArchivoExiste2(Archivo) Then
        MsgBox "File : " & Archivo & " doesn't exists", vbCritical
        Exit Function
    End If
        
    Archivo = util.StripPath(App.Path) & "plus\slideshow\slidebase2.js"
    If Not ArchivoExiste2(Archivo) Then
        MsgBox "File : " & Archivo & " doesn't exists", vbCritical
        Exit Function
    End If
            
    nFreeFile = FreeFile
    
    Open Archivo For Input As #nFreeFile
        Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, linea
            buffer2.Append linea & vbNewLine
        Loop
    Close #nFreeFile
    
    Archivo = vbNullString
    
    buffer.Append "<HTML>" & vbNewLine
    buffer.Append "<HEAD>" & vbNewLine
    buffer.Append "<TITLE>JavaScript Slideshow Template</TITLE>" & vbNewLine
    buffer.Append "<meta http-equiv=" & Chr$(34) & "Content-Type" & Chr$(34) & " content=" & Chr$(34) & "text/html; charset=iso-8859-1" & Chr$(34) & ">" & vbNewLine
    buffer.Append "<SCRIPT TYPE=" & Chr(34) & "text/javascript" & Chr$(34) & " SRC=" & Chr$(34) & "slideshow.js" & Chr$(34) & ">" & vbNewLine
    buffer.Append "<!--" & vbNewLine
    buffer.Append "" & vbNewLine
    buffer.Append "" & "// JavaScript Slideshow by Patrick Fitzgerald" & vbNewLine
    buffer.Append "" & "// http://slideshow.barelyfitz.com/" & vbNewLine
    buffer.Append "// -->" & vbNewLine
    buffer.Append "</SCRIPT>" & vbNewLine
    buffer.Append "" & vbNewLine
    buffer.Append "<SCRIPT TYPE=" & Chr$(34) & "text/javascript" & Chr$(34) & ">" & vbNewLine
    buffer.Append "<!--" & vbNewLine
    buffer.Append "SLIDES = new slideshow(" & Chr$(34) & "SLIDES" & Chr$(34) & ");" & vbNewLine
    buffer.Append "SLIDES.timeout = " & txtTimeout.Text & ";" & vbNewLine
    buffer.Append "SLIDES.prefetch = " & cboPrefetch.ItemData(cboPrefetch.ListIndex) & ";" & vbNewLine
    
    If chkRepeat.Value Then
        buffer.Append "SLIDES.repeat = true;" & vbNewLine
    Else
        buffer.Append "SLIDES.repeat = false;" & vbNewLine
    End If

    For k = 1 To lvwImages.ListItems.count
        buffer.Append "s = new slide();" & vbNewLine
        buffer.Append "s.src = " & Chr$(34) & lvwImages.ListItems(k).Text & Chr$(34) & ";" & vbNewLine
        buffer.Append "s.text = unescape(" & Chr$(34) & Chr$(34) & ");" & vbNewLine
        buffer.Append "s.link = " & Chr$(34) & lvwImages.ListItems(k).SubItems(1) & Chr$(34) & ";" & vbNewLine
        buffer.Append "s.target = " & Chr$(34) & lvwImages.ListItems(k).SubItems(2) & Chr$(34) & ";" & vbNewLine
        buffer.Append "s.attr = " & Chr$(34) & lvwImages.ListItems(k).SubItems(3) & Chr$(34) & ";" & vbNewLine
        buffer.Append "s.filter = " & Chr$(34) & Chr$(34) & ";" & vbNewLine
        buffer.Append "SLIDES.add_slide(s);" & vbNewLine
        buffer.Append "" & vbNewLine
    Next k
    
    If chkSuffle.Value Then
        buffer.Append "if (true) SLIDES.shuffle();" & vbNewLine
    Else
        buffer.Append "if (false) SLIDES.shuffle();" & vbNewLine
    End If
    
    buffer.Append "//-->" & vbNewLine
    buffer.Append "</SCRIPT>" & vbNewLine
    
    buffer3.Append "<BODY ONLOAD=" & Chr$(34) & "SLIDES.restore_position('SS_POSITION');SLIDES.update();" & Chr$(34)
    buffer3.Append " ONUNLOAD=" & Chr$(34) & "SLIDES.save_position('SS_POSITION');" & Chr$(34) & ">" & vbNewLine
    
    If chkControls(0) Then buffer.Append "<STRONG><A HREF=" & Chr$(34) & "javascript:SLIDES.previous()" & Chr$(34) & ">&lt;previous</A></STRONG>" & vbNewLine
    If chkControls(1) Then buffer.Append "<STRONG><A HREF=" & Chr$(34) & "javascript:SLIDES.next()" & Chr$(34) & ">next&gt;</A></STRONG>" & vbNewLine
    If chkControls(1) Then buffer.Append "<STRONG><A HREF=" & Chr$(34) & "javascript:SLIDES.goto_random_slide()" & Chr$(34) & ">random</A></STRONG>" & vbNewLine
    If chkControls(1) Then buffer.Append "<STRONG><A HREF=" & Chr$(34) & "javascript:SLIDES.hotlink()" & Chr$(34) & ">view</A></STRONG>" & vbNewLine
    If chkControls(1) Then buffer.Append "<STRONG><A HREF=" & Chr$(34) & "javascript:SLIDES.play()" & Chr$(34) & ">play</A></STRONG>" & vbNewLine
    If chkControls(1) Then buffer.Append "<STRONG><A HREF=" & Chr$(34) & "javascript:SLIDES.pause()" & Chr$(34) & ">stop</A></STRONG>" & vbNewLine

    buffer3.Append "<P>" & vbNewLine
    buffer3.Append "<a href=" & "javascript:SLIDES.hotlink()" & "><img name=" & Chr$(34) & "SLIDESIMG" & Chr$(34)
    buffer3.Append " src=" & Chr$(34) & lvwImages.ListItems(1).Text & Chr$(34) & " STYLE=" & Chr$(34) & arr_filters(cboTransitions.ListIndex) & Chr$(34) & " BORDER=0 alt=" & Chr$(34) & "Slideshow image" & Chr$(34) & "></A>" & vbNewLine

    buffer3.Append "<SCRIPT type=" & "text/javascript" & ">" & vbNewLine
    buffer3.Append "<!--" & vbNewLine
    buffer3.Append "if (document.images) {" & vbNewLine
    buffer3.Append "SLIDES.image = document.images.SLIDESIMG;" & vbNewLine
    buffer3.Append "SLIDES.textid = " & "SLIDESTEXT" & ";" & vbNewLine
    buffer3.Append "SLIDES.update();" & vbNewLine
    buffer3.Append "SLIDES.play();" & vbNewLine
    buffer3.Append "}" & vbNewLine
    buffer3.Append "//-->" & vbNewLine
    buffer3.Append "</SCRIPT>" & vbNewLine
    buffer3.Append "" & vbNewLine
    buffer3.Append "<BR CLEAR=all>" & vbNewLine
    buffer3.Append "" & vbNewLine
    buffer3.Append "<NOSCRIPT>" & vbNewLine
    buffer3.Append "<HR>" & vbNewLine
    buffer3.Append "Since your web browser does not support JavaScript," & vbNewLine
    buffer3.Append "here is a non-JavaScript version of the image slideshow:" & vbNewLine
    buffer3.Append "<P>" & vbNewLine
    buffer3.Append "<P>" & vbNewLine
    
    For k = 1 To lvwImages.ListItems.count
        buffer3.Append "<IMG SRC=" & lvwImages.ListItems(k).Text & " ALT=" & "slideshow image" & "><BR>" & vbNewLine
    Next k
    
    buffer3.Append "</P>" & vbNewLine
    buffer3.Append "<HR>" & vbNewLine
    buffer3.Append "" & vbNewLine
    buffer3.Append "</NOSCRIPT>" & vbNewLine
    
    buffer3.Append "</body>" & vbNewLine
    buffer3.Append "</html>" & vbNewLine
    
    If preview Then
        Archivo = util.StripPath(App.Path) & "plus\slideshow\slideshowtest.htm"
        Open Archivo For Output As #nFreeFile
            Print #nFreeFile, buffer.ToString
            Print #nFreeFile, buffer2.ToString
            Print #nFreeFile, buffer3.ToString
        Close #nFreeFile
    Else
        If Cdlg.VBGetSaveFileName(Archivo, , , strGlosa(), , LastPath, "Save File As ...", "htm", hwnd) Then
        
            MsgBox "Now you must copy this file and the slideshow.js file inside the your source images folder.", vbInformation
            
            util.CopiarArchivo util.StripPath(App.Path) & "plus\slideshow\slideshow.js", util.PathArchivo(Archivo) & "slideshow.js"
            Open Archivo For Output As #nFreeFile
                Print #nFreeFile, buffer.ToString
                Print #nFreeFile, buffer2.ToString
                Print #nFreeFile, buffer3.ToString
            Close #nFreeFile
        End If
    End If
        
    util.Hourglass hwnd, False
    
    Set buffer = Nothing
    Set buffer2 = Nothing
    Set buffer3 = Nothing
    
    GeneraSlideShow = True
    
End Function

Private Function Validar() As Boolean

    Dim k As Integer
    Dim flag As Boolean
    
    If txtTimeout.Text = "" Then
        txtTimeout.SetFocus
        Exit Function
    End If
    
    For k = 0 To 5
        If chkControls(k).Value Then
            flag = True
            Exit For
        End If
    Next k
    
    If Not flag Then
        MsgBox "You must select a slide control to display", vbCritical
        Exit Function
    End If
    
    If lvwImages.ListItems.count = 0 Then
        MsgBox "There is no images to slide", vbCritical
        Exit Function
    End If
    
    Validar = True
    
End Function

Private Sub cmd_Click(Index As Integer)

    Select Case Index
        Case 0  'ok
            If Validar() Then
                If GeneraSlideShow(False) Then
                    Unload Me
                End If
            End If
        Case 1  'exit
            Unload Me
        Case 2  'add image
            Call add_image
        Case 3  'remove image
            If lvwImages.ListItems.count > 0 Then
                If Not lvwImages.SelectedItem Is Nothing Then
                    lvwImages.ListItems.Remove lvwImages.SelectedItem.key
                End If
            End If
        Case 4  'add folder
            Call add_folder
        Case 5  'remove all
            If lvwImages.ListItems.count > 0 Then
                If Confirma("Are you sure") = vbYes Then
                    lvwImages.ListItems.Clear
                    idx = 0
                End If
            End If
        Case 6  'preview
            If Validar() Then
                If GeneraSlideShow(True) Then
                    util.ShellFunc util.StripPath(App.Path) & "plus\slideshow\slideshowtest.htm", vbNormalFocus
                End If
            End If
        Case 7  'update image
            If lvwImages.ListItems.count > 0 Then
                If Not lvwImages.SelectedItem Is Nothing Then
                    lvwImages.SelectedItem.Text = txtFileName.Text
                    lvwImages.SelectedItem.SubItems(1) = txtLink.Text
                    lvwImages.SelectedItem.SubItems(2) = txtTarget.Text
                    lvwImages.SelectedItem.SubItems(3) = txtAttributes.Text
                End If
            End If
    End Select
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
        
    util.SetNumber txtTimeout.hwnd
    
    cboPrefetch.AddItem "all slides"
    cboPrefetch.ItemData(cboPrefetch.NewIndex) = -1
    
    cboPrefetch.AddItem "0 slides"
    cboPrefetch.ItemData(cboPrefetch.NewIndex) = 0
    
    cboPrefetch.AddItem "1 slides"
    cboPrefetch.ItemData(cboPrefetch.NewIndex) = 1
    
    cboPrefetch.AddItem "2 slides"
    cboPrefetch.ItemData(cboPrefetch.NewIndex) = 2
    
    cboPrefetch.ListIndex = 0
    
    cboTransitions.AddItem "none"
    cboTransitions.AddItem "Barn"
    cboTransitions.AddItem "Blinds"
    cboTransitions.AddItem "CheckerBoard"
    cboTransitions.AddItem "Fade"
    cboTransitions.AddItem "GradientWipe"
    cboTransitions.AddItem "Inset"
    cboTransitions.AddItem "Iris"
    cboTransitions.AddItem "Pixelate"
    cboTransitions.AddItem "RadialWipe"
    cboTransitions.AddItem "RandomBars"
    cboTransitions.AddItem "RandomDissolve"
    cboTransitions.AddItem "Slide"
    cboTransitions.AddItem "Spiral"
    cboTransitions.AddItem "Stretch"
    cboTransitions.AddItem "Strips"
    cboTransitions.AddItem "Wheel"
    cboTransitions.AddItem "ZigZag"
    cboTransitions.ListIndex = 4
    
    ReDim arr_filters(17)
    
    arr_filters(0) = vbNullString
    arr_filters(1) = "filter:progid:DXImageTransform.Microsoft.Barn()"
    arr_filters(2) = "filter:progid:DXImageTransform.Microsoft.Blinds()"
    arr_filters(3) = "filter:progid:DXImageTransform.Microsoft.CheckerBoard()"
    arr_filters(4) = "filter:progid:DXImageTransform.Microsoft.Fade()"
    arr_filters(5) = "filter:progid:DXImageTransform.Microsoft.GradientWipe()"
    arr_filters(6) = "filter:progid:DXImageTransform.Microsoft.Inset()"
    arr_filters(7) = "filter:progid:DXImageTransform.Microsoft.Iris()"
    arr_filters(8) = "filter:progid:DXImageTransform.Microsoft.Pixelate()"
    arr_filters(9) = "filter:progid:DXImageTransform.Microsoft.RadialWipe()"
    arr_filters(10) = "filter:progid:DXImageTransform.Microsoft.RandomBars()"
    arr_filters(11) = "filter:progid:DXImageTransform.Microsoft.RandomDissolve()"
    arr_filters(12) = "filter:progid:DXImageTransform.Microsoft.Slide()"
    arr_filters(13) = "filter:progid:DXImageTransform.Microsoft.Spiral()"
    arr_filters(14) = "filter:progid:DXImageTransform.Microsoft.Stretch()"
    arr_filters(15) = "filter:progid:DXImageTransform.Microsoft.Strips()"
    arr_filters(16) = "filter:progid:DXImageTransform.Microsoft.Wheel()"
    arr_filters(17) = "filter:progid:DXImageTransform.Microsoft.ZigZag()"
  
    Set imgIE.Picture = LoadResPicture(1007, vbResBitmap)
    Set imgNE.Picture = LoadResPicture(1009, vbResBitmap)
    Set imgOP.Picture = LoadResPicture(1010, vbResBitmap)
        
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload :" & Me.Name
    Set frmSlideShow = Nothing
End Sub



Private Sub lvwImages_ItemClick(ByVal ITem As MSComctlLib.ListItem)

    If lvwImages.ListItems.count > 0 Then
        If Not ITem Is Nothing Then
            txtFileName.Text = ITem.Text
            txtLink.Text = ITem.SubItems(1)
            txtTarget.Text = ITem.SubItems(2)
            txtAttributes.Text = ITem.SubItems(3)
        End If
    End If
    
End Sub


