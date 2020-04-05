VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "ccrpftv6.ocx"
Begin VB.Form frmThumbs 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Image Browser"
   ClientHeight    =   6120
   ClientLeft      =   825
   ClientTop       =   3705
   ClientWidth     =   8265
   Icon            =   "Thumbs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   551
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3225
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   8
      Top             =   2685
      Width           =   900
      Begin VB.PictureBox picProgressSlide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   0
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   9
         Top             =   0
         Width           =   255
      End
   End
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   5805
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   529
            MinWidth        =   529
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13520
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picThumb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   3915
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   4
      Top             =   2265
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.FileListBox filHidden 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   3225
      Pattern         =   "*.bmp;*.dib;*.rle;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur;*.jpeg"
      TabIndex        =   6
      Top             =   2130
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox picLoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   4995
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   3
      Top             =   900
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4890
      Left            =   3090
      ScaleHeight     =   326
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   0
      Top             =   390
      Width           =   4125
      Begin VB.PictureBox picSlide 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4710
         Left            =   555
         ScaleHeight     =   314
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   231
         TabIndex        =   2
         Top             =   60
         Width           =   3465
         Begin VB.OptionButton optThumb 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   1440
            Index           =   0
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   45
            Visible         =   0   'False
            Width           =   1200
         End
      End
      Begin VB.VScrollBar vsbSlide 
         Height          =   2100
         Left            =   3180
         TabIndex        =   1
         Top             =   405
         Width           =   210
      End
   End
   Begin MSComctlLib.ImageList imlToolsHover 
      Left            =   3555
      Top             =   1395
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Thumbs.frx":014A
            Key             =   "Browse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Thumbs.frx":04E6
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Thumbs.frx":0882
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Thumbs.frx":0C1E
            Key             =   "Refresh"
         EndProperty
      EndProperty
   End
   Begin CCRPFolderTV6.FolderTreeview dirMain 
      Height          =   2880
      Left            =   30
      TabIndex        =   10
      ToolTipText     =   "Seleccione el directorio a desplegar"
      Top             =   405
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   5080
      Appearance      =   0
      VirtualFolders  =   0   'False
   End
End
Attribute VB_Name = "frmThumbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tipo_ins As Integer
Public Imagen As String

Private Type POINTAPI
    X  As Long
    y  As Long
End Type

Private mbActive                As Boolean
Private mlCurThumb              As Long
Private Const SRCCOPY           As Long = &HCC0020
Private Const STRETCH_HALFTONE  As Long = &H4&
Private Const SW_RESTORE        As Long = &H9&

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lpPt As POINTAPI) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Browse(ByVal sPath As String)

    On Error Resume Next
    
    'Dim lRet    As Long
    
    If Len(sPath) > 0 Then
        filHidden.Path = sPath
        If Err.Number = 0 Then
            CreateThumbs
            Call util.GrabaIni(IniPath, "Main", "Last Path", sPath)
        Else
            CreateThumbs
            Call util.GrabaIni(IniPath, "Main", "Last Path", App.Path)
        End If
    End If
            
    Err = 0
    
End Sub

Private Sub CreateThumbPic(picSource As PictureBox, picThumb As PictureBox)

'This sub uses the halftone stretch mode, which produces the highest
'quality possible, when stretching the bitmap.

Dim lRet            As Long
Dim lLeft           As Long
Dim lTop            As Long
Dim lWidth          As Long
Dim lHeight         As Long
Dim lForeColor      As Long
Dim hBrush          As Long
Dim hDummyBrush     As Long
Dim lOrigMode       As Long
Dim fScale          As Single
Dim uBrushOrigPt    As POINTAPI

    picThumb.Width = 64
    picThumb.Height = 64
    'picThumb.BackColor = vbButtonFace
    picThumb.AutoRedraw = True
    picThumb.Cls
    
    If picSource.Width <= picThumb.Width - 2 And picSource.Height <= picThumb.Height - 2 Then
        fScale = 1
    Else
        fScale = IIf(picSource.Width > picSource.Height, (picThumb.Width - 2) / picSource.Width, (picThumb.Height - 2) / picSource.Height)
    End If
    lWidth = picSource.Width * fScale
    lHeight = picSource.Height * fScale
    lLeft = Int((picThumb.Width - lWidth) / 2)
    lTop = Int((picThumb.Height - lHeight) / 2)
    
    'Store the original ForeColor
    lForeColor = picThumb.ForeColor
    
    'Set picEdit's stretch mode to halftone (this may cause misalignment of the brush)
    lOrigMode = SetStretchBltMode(picThumb.hdc, STRETCH_HALFTONE)
    
    'Realign the brush...
    'Get picEdit's brush by selecting a dummy brush into the DC
    hDummyBrush = CreateSolidBrush(lForeColor)
    hBrush = SelectObject(picThumb.hdc, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    lRet = UnrealizeObject(hBrush)
    'Set picEdit's brush alignment coordinates to the left-top of the bitmap
    lRet = SetBrushOrgEx(picThumb.hdc, lLeft, lTop, uBrushOrigPt)
    'Now put the original brush back into the DC at the new alignment
    hDummyBrush = SelectObject(picThumb.hdc, hBrush)
    
    'Stretch the bitmap
    lRet = StretchBlt(picThumb.hdc, lLeft, lTop, lWidth, lHeight, _
            picSource.hdc, 0, 0, picSource.Width, picSource.Height, SRCCOPY)
    
    'Set the stretch mode back to it's original mode
    lRet = SetStretchBltMode(picThumb.hdc, lOrigMode)
    
    'Reset the original alignment of the brush...
    'Get picEdit's brush by selecting the dummy brush back into the DC
    hBrush = SelectObject(picThumb.hdc, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    lRet = UnrealizeObject(hBrush)
    'Set the brush alignment back to the original coordinates
    lRet = SetBrushOrgEx(picThumb.hdc, uBrushOrigPt.X, uBrushOrigPt.y, uBrushOrigPt)
    'Now put the original brush back into picEdit's DC at the original alignment
    hDummyBrush = SelectObject(picThumb.hdc, hBrush)
    'Get rid of the dummy brush
    lRet = DeleteObject(hDummyBrush)
    
    'Restore the original ForeColor
    picThumb.ForeColor = lForeColor

    picThumb.Line (lLeft - 1, lTop - 1)-Step(lWidth + 1, lHeight + 1), &H0&, B
    
End Sub

Private Sub CreateThumbs()

Dim iMaxLen As Integer
'Dim x       As Long
'Dim y       As Long
Dim lIdx    As Long
Dim lPicCnt As Long
Dim lFilCnt As Long
Dim sPath   As String
Dim sText   As String

    Screen.MousePointer = vbHourglass
    filHidden.Refresh
    
    picSlide.Move 0, 0, optThumb(0).Width, optThumb(0).Height
    picSlide.Visible = False
    'picSlide.BackColor = vbButtonFace
    Set picSlide.Font = optThumb(0).Font
    While optThumb.count > 1
        Unload optThumb(optThumb.count - 1)
    Wend
    optThumb(0).Visible = False
    DoEvents
    On Error Resume Next
    sPath = filHidden.Path
    sPath = sPath & IIf(VBA.Right$(sPath, 1) <> "\", "\", "")
    lFilCnt = filHidden.ListCount
    Me.Caption = "Image Browser : " & "Images: " & Format$(lFilCnt, "#,##0")
    If Len(sPath) > 0 Then
        Call StartProgress
        For lIdx = 0 To filHidden.ListCount - 1
            Call UpdateProgress((CSng(lIdx + 1) / CSng(lFilCnt)) * 100, filHidden.List(lIdx))
            Set picLoad.Picture = LoadPicture()
            picLoad.Cls
            Err.Clear
            If InStr(1, LCase$(filHidden.List(lIdx)), ".ico") > 0 _
              Or InStr(1, LCase$(filHidden.List(lIdx)), ".cur") > 0 Then
                Set picLoad.Picture = LoadPicture(sPath & filHidden.List(lIdx), vbLPLargeShell, vbLPDefault)
            Else
                Set picLoad.Picture = LoadPicture(sPath & filHidden.List(lIdx))
            End If
            If Err.Number = 0 Then
                Call CreateThumbPic(picLoad, picThumb)
                If lPicCnt > 0 Then
                    Load optThumb(lPicCnt)
                    Set optThumb(lPicCnt).Container = picSlide
                End If
                optThumb(lPicCnt).Tag = filHidden.List(lIdx)
                Set optThumb(lPicCnt).Picture = picThumb.Image
                sText = filHidden.List(lIdx)
                iMaxLen = optThumb(lPicCnt).Width - 15
                If picSlide.TextWidth(sText) > iMaxLen Then
                    iMaxLen = iMaxLen - picSlide.TextWidth("...")
                End If
                While picSlide.TextWidth(sText) > iMaxLen
                    sText = Left$(sText, Len(sText) - 1)
                Wend
                If iMaxLen < optThumb(lPicCnt).Width - 15 Then
                    sText = sText & "..."
                End If
                optThumb(lPicCnt).Caption = sText
                optThumb(lPicCnt).Visible = True
                lPicCnt = lPicCnt + 1
            Else
                Debug.Print Error$
            End If
        Next lIdx
        
        picProgress.Visible = False
        
        'Free the unneeded resources
        Set picLoad.Picture = LoadPicture()
        Set picThumb.Picture = LoadPicture()
        optThumb(0).Value = True
        mlCurThumb = 0
        Call Form_Resize
        picSlide.Visible = True
        'lblCount.Caption = "Images: " & Format$(lPicCnt, "#,##0")
        
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Edit()

Dim lRet    As Long
Dim sFIle   As String

    sFIle = filHidden.Path
    If Len(sFIle) > 0 Then
        sFIle = sFIle & IIf(VBA.Right$(sFIle, 1) <> "\", "\", "")
        sFIle = sFIle & optThumb(mlCurThumb).Tag
        lRet = ShellExecute(Me.hWnd, "Open", sFIle, &H0&, &H0&, SW_RESTORE)
    End If

End Sub

Private Sub StartProgress()

    With picProgress
        .Cls
        .BackColor = vbButtonFace
        .ForeColor = vbButtonText
        .Move sbrMain.Left + sbrMain.Panels(2).Left, sbrMain.Top + 1, _
            sbrMain.Panels(2).Width, sbrMain.Height - 1
    End With
    
    With picProgressSlide
        .Cls
        .BackColor = vbHighlight
        .ForeColor = vbHighlightText
        .Move 0, 0, 1, picProgress.ScaleHeight
    End With
    
    picProgress.Visible = True
    
End Sub

Private Sub UpdateProgress(ByVal iPercent As Integer, ByVal sCaption As String)

Dim lTextTop    As Long

    picProgress.Cls
    picProgressSlide.Cls
    picProgressSlide.Width = picProgress.ScaleWidth * (CSng(iPercent) / 100!)
    lTextTop = (picProgress.ScaleHeight - picProgress.TextHeight(sCaption)) / 2
    picProgress.CurrentX = 3
    picProgress.CurrentY = lTextTop
    picProgress.Print sCaption
    picProgressSlide.CurrentX = 3
    picProgressSlide.CurrentY = lTextTop
    picProgressSlide.Print sCaption
    DoEvents
    
End Sub

Private Sub dirMain_FolderClick(Folder As CCRPFolderTV6.Folder, location As CCRPFolderTV6.ftvHitTestConstants)

    On Error GoTo errordir
    
    Dim Path As String
    
    If Folder.Name = "Escritorio" Or Folder.Name = "Desktop" Then
        Path = util.GetSpecialfolder(eCSIDL_DESKTOPDIRECTORY)
    ElseIf Folder.Name = "Mi PC" Or Folder.Name = "My PC" Then
        Path = util.GetSpecialfolder(eCSIDL_DRIVES)
    End If
    
    If Len(Path) = 0 Then Path = Folder.FullPath
    
    Browse Path
    
    Exit Sub
    
errordir:
    Exit Sub
    
End Sub

Private Sub filHidden_PathChange()

Dim sPath As String

    sPath = filHidden.Path
    If Len(sPath) > 0 Then
        sPath = sPath & IIf(VBA.Right$(sPath, 1) <> "\", "\", "")
    End If
    sbrMain.Panels(2).Text = sPath
    'lblCount.Caption = "Images: 0"
    
End Sub


Private Sub Form_Activate()

    If Not mbActive Then
        picSlide.Visible = False
        DoEvents
        mbActive = True
        Call Browse(True)
    End If
        
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub


Private Sub Form_Load()

    util.CenterForm Me
    DrawXPCtl Me
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim lIdx As Long

    For lIdx = 1 To optThumb.count - 1
        Set optThumb(lIdx).Picture = Nothing
        Unload optThumb(lIdx)
    Next lIdx
    Set optThumb(0).Picture = Nothing
    
    Call clear_memory(Me)
    
End Sub

Private Sub Form_Resize()

Dim X       As Long
Dim y       As Long
Dim lIdx    As Long
Dim lCols   As Long
        
    If Me.WindowState <> vbMinimized Then
        If Me.Width < 346 * Screen.TwipsPerPixelX Then
            Me.Width = 346 * Screen.TwipsPerPixelX
        ElseIf Me.Height < 378 * Screen.TwipsPerPixelY Then
            Me.Height = 378 * Screen.TwipsPerPixelY
        Else
            dirMain.Move 0, 0, dirMain.Width, Me.ScaleHeight - sbrMain.Height
            picFrame.Move dirMain.Width + 1, 0, Me.ScaleWidth - dirMain.Width - 1, Me.ScaleHeight - sbrMain.Height
            
            vsbSlide.Move picFrame.ScaleWidth - vsbSlide.Width, 0, vsbSlide.Width, picFrame.ScaleHeight
            lCols = Int((picFrame.ScaleWidth - vsbSlide.Width) / optThumb(0).Width)
            For lIdx = 0 To optThumb.count - 1
                X = (lIdx Mod lCols) * optThumb(0).Width
                y = Int(lIdx / lCols) * optThumb(0).Height
                optThumb(lIdx).Move X, y
            Next lIdx
            picSlide.Width = lCols * optThumb(0).Width
            picSlide.Height = Int(optThumb.count / lCols) * optThumb(0).Height
            If Int(optThumb.count / lCols) < (optThumb.count / lCols) Then
                picSlide.Height = picSlide.Height + optThumb(0).Height
            End If
            vsbSlide.Value = 0
            vsbSlide.Max = picSlide.Height - picFrame.ScaleHeight
            If vsbSlide.Max < 0 Then
                vsbSlide.Max = 0
                vsbSlide.Enabled = False
            Else
                vsbSlide.Enabled = True
                vsbSlide.SmallChange = optThumb(0).Height
                vsbSlide.LargeChange = picFrame.ScaleHeight
            End If
        End If
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
   
    Set frmThumbs = Nothing
    
End Sub

Private Sub optThumb_Click(Index As Integer)
    
    Dim sPath As String

    sPath = filHidden.Path
    If Len(sPath) > 0 Then
        sPath = sPath & IIf(VBA.Right$(sPath, 1) <> "\", "\", "")
    End If
    sbrMain.Panels(2).Text = sPath & optThumb(Index).Tag
    Imagen = sbrMain.Panels(2).Text
    mlCurThumb = Index

End Sub

Private Sub optThumb_DblClick(Index As Integer)

    Dim sPath As String

    sPath = filHidden.Path
    If Len(sPath) > 0 Then
        sPath = sPath & IIf(VBA.Right$(sPath, 1) <> "\", "\", "")
    End If
    sbrMain.Panels(2).Text = sPath & optThumb(Index).Tag
    Imagen = sbrMain.Panels(2).Text
    mlCurThumb = Index
    
    If tipo_ins = 0 Then
        Edit
    Else
        If Len(optThumb(Index).Tag) > 0 Then
            frmImage.archivo_img = Imagen
        End If
        Unload Me
    End If
    
End Sub


Private Sub vsbSlide_Change()

    picSlide.Top = -vsbSlide.Value
    picFrame.SetFocus
    
End Sub


Private Sub vsbSlide_Scroll()

    vsbSlide_Change

End Sub


