VERSION 5.00
Begin VB.Form frmImageEffect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Special Effects - (none)"
   ClientHeight    =   5160
   ClientLeft      =   1050
   ClientTop       =   3075
   ClientWidth     =   11250
   Icon            =   "frmImageEffect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Preview"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9960
      TabIndex        =   10
      Top             =   1500
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   5160
      ScaleHeight     =   295
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   311
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.PictureBox picSave 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   270
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   207
      TabIndex        =   9
      Top             =   6105
      Width           =   3135
   End
   Begin VB.PictureBox picOri 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   3600
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   8
      Top             =   6120
      Width           =   1560
   End
   Begin VB.CommandButton cmdCopyImage 
      Caption         =   "Copy Image"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9960
      TabIndex        =   7
      Top             =   930
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply Effect"
      Height          =   495
      Left            =   9960
      TabIndex        =   3
      Top             =   345
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   360
      Max             =   255
      Min             =   -255
      TabIndex        =   2
      Top             =   4830
      Width           =   9480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   5160
      ScaleHeight     =   295
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   311
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   360
      ScaleHeight     =   295
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   311
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Image with effect applied"
      Height          =   195
      Left            =   5160
      TabIndex        =   5
      Top             =   120
      Width           =   1770
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Original"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   525
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save As..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnutraco 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEffects 
      Caption         =   "Effects"
      Enabled         =   0   'False
      Begin VB.Menu mnu1 
         Caption         =   "Color Adjustment"
         Begin VB.Menu mnuBrilho 
            Caption         =   "Brightness"
         End
         Begin VB.Menu mnuContraste 
            Caption         =   "Contrast"
         End
         Begin VB.Menu mnuNitidez 
            Caption         =   "Sharpening"
         End
         Begin VB.Menu mnuGamma 
            Caption         =   "Gamma Adjust"
         End
         Begin VB.Menu mnuReduceColors 
            Caption         =   "Reduce Colors"
         End
         Begin VB.Menu mnuEightColors 
            Caption         =   "Reduce to 8 colors"
         End
         Begin VB.Menu mnuShift 
            Caption         =   "Shift Effect"
         End
         Begin VB.Menu mnuSaturation 
            Caption         =   "Saturation"
         End
         Begin VB.Menu mnuHue 
            Caption         =   "Hue Adjust"
         End
         Begin VB.Menu mnuColBalance 
            Caption         =   "Color Balance"
         End
         Begin VB.Menu mnuWebColors 
            Caption         =   "WebColors Mode"
         End
         Begin VB.Menu mnuMediumTones 
            Caption         =   "Medium Tones"
         End
         Begin VB.Menu mnuStretchHisto 
            Caption         =   "Stretch Histogram"
         End
      End
      Begin VB.Menu mnu2 
         Caption         =   "Blur"
         Begin VB.Menu mnuAlias 
            Caption         =   "AntiAlias"
         End
         Begin VB.Menu mnuBlur 
            Caption         =   "Blur"
         End
         Begin VB.Menu mnuSmartBlur 
            Caption         =   "SmartBlur"
         End
         Begin VB.Menu mnuMoreBlur 
            Caption         =   "More Blur"
         End
         Begin VB.Menu mnuSoftnerBlur 
            Caption         =   "Softner Blur"
         End
         Begin VB.Menu mnuMotionBlur 
            Caption         =   "Motion Blur"
         End
         Begin VB.Menu mnuFarBlur 
            Caption         =   "Far Blur"
         End
         Begin VB.Menu mnuRadialBlur 
            Caption         =   "Radial Blur"
         End
         Begin VB.Menu mnuZoomBlur 
            Caption         =   "Zoom Blur"
         End
         Begin VB.Menu mnuUnsharpMask 
            Caption         =   "Unsharp Mask"
         End
      End
      Begin VB.Menu mnu3 
         Caption         =   "Tones"
         Begin VB.Menu mnuGrayScale 
            Caption         =   "Gray Tones"
         End
         Begin VB.Menu mnuSepia 
            Caption         =   "Sepia Effect"
         End
         Begin VB.Menu mnuAmbient 
            Caption         =   "Ambient Light"
         End
         Begin VB.Menu mnuTone 
            Caption         =   "Tone Adjust"
         End
      End
      Begin VB.Menu mnu4 
         Caption         =   "Distortion"
         Begin VB.Menu mnuMosaico 
            Caption         =   "Mosaic"
         End
         Begin VB.Menu mnuDiffuse 
            Caption         =   "Diffuse"
         End
         Begin VB.Menu mnuRock 
            Caption         =   "Rock Effect"
         End
         Begin VB.Menu mnuNoise 
            Caption         =   "Noise"
         End
         Begin VB.Menu mnuMelt 
            Caption         =   "Melt"
         End
         Begin VB.Menu mnuFishEye 
            Caption         =   "Fish Eye"
         End
         Begin VB.Menu mnuFishEyeEx 
            Caption         =   "Fish Eye Ex"
         End
         Begin VB.Menu mnuTwirl 
            Caption         =   "Twirl"
         End
         Begin VB.Menu mnuTwirlEx 
            Caption         =   "TwirlEx"
         End
         Begin VB.Menu mnuSwirl 
            Caption         =   "Swirl"
         End
         Begin VB.Menu mnu3D 
            Caption         =   "Make 3D"
         End
         Begin VB.Menu mnu4Corners 
            Caption         =   "Four Corners"
         End
         Begin VB.Menu mnuCaricature 
            Caption         =   "Caricature"
         End
         Begin VB.Menu mnuRoll 
            Caption         =   "Enroll"
         End
         Begin VB.Menu mnuPolar 
            Caption         =   "Polar Coordinates"
         End
         Begin VB.Menu mnuCilindrical 
            Caption         =   "Cilindrical"
         End
      End
      Begin VB.Menu mnuW 
         Caption         =   "Waves"
         Begin VB.Menu mnuWave 
            Caption         =   "Waves"
         End
         Begin VB.Menu mnuBlockWaves 
            Caption         =   "Block Waves"
         End
         Begin VB.Menu mnuCircularWaves 
            Caption         =   "Circular Waves"
         End
         Begin VB.Menu mnuCircularWavesEx 
            Caption         =   "Circular Waves Enhanced"
         End
      End
      Begin VB.Menu mnu5 
         Caption         =   "Borders"
         Begin VB.Menu mnuBack 
            Caption         =   "Backdrop Removal"
         End
         Begin VB.Menu mnuEmbEng 
            Caption         =   "Emboss / Engrave"
         End
         Begin VB.Menu mnuNeon 
            Caption         =   "Neon"
         End
         Begin VB.Menu mnuBorders 
            Caption         =   "Detect Borders"
         End
         Begin VB.Menu mnuFindEdges 
            Caption         =   "Find Edges"
         End
         Begin VB.Menu mnuNotePaper 
            Caption         =   "Note Paper"
         End
      End
      Begin VB.Menu mnuBlend 
         Caption         =   "Blend Modes"
         Visible         =   0   'False
         Begin VB.Menu mnuAlphaBlend 
            Caption         =   "AlphaBlend"
         End
         Begin VB.Menu mnuBlendModes 
            Caption         =   "Blend Modes"
         End
         Begin VB.Menu mnuGlassBlendMode 
            Caption         =   "Glass Blend Mode"
         End
      End
      Begin VB.Menu mnuMetal 
         Caption         =   "Metallic Effects"
         Begin VB.Menu mnuMetallic 
            Caption         =   "Metallic"
         End
         Begin VB.Menu mnuGold 
            Caption         =   "Gold"
         End
         Begin VB.Menu mnuIce 
            Caption         =   "Ice"
         End
      End
      Begin VB.Menu mnu6 
         Caption         =   "Other Effects"
         Begin VB.Menu mnuInvertion 
            Caption         =   "Invertion Adjust"
         End
         Begin VB.Menu mnuMono 
            Caption         =   "Monochrome"
         End
         Begin VB.Menu mnuBackN 
            Caption         =   "Replace Color"
         End
         Begin VB.Menu mnuAscii 
            Caption         =   "Ascii Effect"
         End
         Begin VB.Menu mnuRandomPoints 
            Caption         =   "Random Points"
         End
         Begin VB.Menu mnuSol 
            Caption         =   "Solarize"
         End
         Begin VB.Menu mnuCanvas 
            Caption         =   "Canvas Adjust"
         End
         Begin VB.Menu mnuRelief 
            Caption         =   "Relief"
         End
         Begin VB.Menu mnuTile 
            Caption         =   "Tile Effect"
         End
         Begin VB.Menu mnuFragment 
            Caption         =   "Fragment"
         End
         Begin VB.Menu mnuFog 
            Caption         =   "Fog Effect"
         End
         Begin VB.Menu mnuOilPaint 
            Caption         =   "Oil Paint"
         End
         Begin VB.Menu mnuFrostGlass 
            Caption         =   "Frost Glass"
         End
         Begin VB.Menu mnuRainDrop 
            Caption         =   "Rain Drop"
         End
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "Image"
      Enabled         =   0   'False
      Begin VB.Menu mnuFlipH 
         Caption         =   "Flip Horizontal"
      End
      Begin VB.Menu mnuFlipV 
         Caption         =   "Flip Vertical"
      End
      Begin VB.Menu mnuFlipB 
         Caption         =   "Flip Both"
      End
   End
End
Attribute VB_Name = "frmImageEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
    ByVal Y As Long) As Long

Private ColorPickTool As Boolean
Private LastPath As String
Private resp As Long
Private Resp2 As Long

Private Const TwipsPerPixel As Long = 15
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, _
ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Sub cmdCopyImage_Click()
    GPX_BitBlt Picture1.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.hdc, 0, 0, vbSrcCopy, resp
    GPX_BitBlt picOri.hdc, 0, 0, picSave.ScaleWidth, picSave.ScaleHeight, picSave.hdc, 0, 0, vbSrcCopy, Resp2
    Picture1.Refresh
    picOri.Refresh
End Sub



Private Sub Command1_Click()
    frmPreview.Show vbModal
End Sub

Private Sub Command2_Click()
    Dim Ticks As Long
    Dim Color As Long
    
    On Error Resume Next
    If (Effect <> 14) Then
        Picture2.Cls
        picSave.Cls
    End If
    Ticks = GetTickCount
    Select Case Effect
        Case 4
            GPX_AntiAlias Picture2.hdc, Picture1.hdc, 50, resp
            GPX_AntiAlias picSave.hdc, picOri.hdc, 50, Resp2
        Case 9
            GPX_ReduceTo8Colors Picture2.hdc, Picture1.hdc, resp
            GPX_ReduceTo8Colors picSave.hdc, picOri.hdc, Resp2
        Case 11
            GPX_Sepia Picture2.hdc, Picture1.hdc, resp
            GPX_Sepia picSave.hdc, picOri.hdc, Resp2
        Case 14
            Dim sBuffer As String
            Dim sBuffer2 As String
            
            GPX_AllocBufferSize Picture2.hdc, resp
            sBuffer = Space(resp)
            GPX_AsciiMorph Picture2.hdc, sBuffer, resp
            
            GPX_AllocBufferSize picSave.hdc, Resp2
            sBuffer2 = Space(Resp2)
            GPX_AsciiMorph picSave.hdc, sBuffer, Resp2
        Case 18
            GPX_Diffuse Picture2.hdc, Picture1.hdc, resp
            GPX_Diffuse picSave.hdc, picOri.hdc, Resp2
        Case 23
            GPX_Solarize Picture2.hdc, Picture1.hdc, False, resp
            GPX_Solarize picSave.hdc, picOri.hdc, False, Resp2
        Case 25
            GPX_Melt Picture2.hdc, Picture1.hdc, resp
            GPX_Melt picSave.hdc, picOri.hdc, Resp2
        Case 26
            GPX_FishEye Picture2.hdc, Picture1.hdc, resp
            GPX_FishEye picSave.hdc, picOri.hdc, Resp2
        Case 33
            GPX_Blur Picture2.hdc, Picture1.hdc, resp
            GPX_Blur picSave.hdc, picOri.hdc, Resp2
        Case 34
            GPX_Relief Picture2.hdc, Picture1.hdc, resp
            GPX_Relief picSave.hdc, picOri.hdc, Resp2
        Case 39
            GPX_Make3DEffect Picture2.hdc, Picture1.hdc, 6, resp
            GPX_Make3DEffect picSave.hdc, picOri.hdc, 6, Resp2
        Case 40
            GPX_FourCorners Picture2.hdc, Picture1.hdc, resp
            GPX_FourCorners picSave.hdc, picOri.hdc, Resp2
        Case 41
            GPX_Caricature Picture2.hdc, Picture1.hdc, resp
            GPX_Caricature picSave.hdc, picOri.hdc, Resp2
        Case 43
            GPX_Roll Picture2.hdc, Picture1.hdc, resp
            GPX_Roll picSave.hdc, picOri.hdc, Resp2
        Case 44
            GPX_SmartBlur Picture2.hdc, Picture1.hdc, 20, resp
            GPX_SmartBlur picSave.hdc, picOri.hdc, 20, Resp2
        Case 46
            GPX_SoftnerBlur Picture2.hdc, Picture1.hdc, resp
            GPX_SoftnerBlur picSave.hdc, picOri.hdc, Resp2
        Case 53
            GPX_WebColors Picture2.hdc, Picture1.hdc, resp
            GPX_WebColors picSave.hdc, picOri.hdc, Resp2
        Case 58
            GPX_PolarCoordinates Picture2.hdc, Picture1.hdc, 0, resp
            GPX_PolarCoordinates picSave.hdc, picOri.hdc, 0, Resp2
        Case 60
            GPX_FrostGlass Picture2.hdc, Picture1.hdc, 3, resp
            GPX_FrostGlass picSave.hdc, picOri.hdc, 3, Resp2
        Case 63
            GPX_RainDrops Picture2.hdc, Picture1.hdc, 40, 50, 40, resp
            GPX_RainDrops picSave.hdc, picOri.hdc, 40, 50, 40, Resp2
        Case 67
            GPX_StretchHistogram Picture2.hdc, Picture1.hdc, HST_COLOR, 1, resp
            GPX_StretchHistogram picSave.hdc, picOri.hdc, HST_COLOR, 1, Resp2
    End Select
    Ticks = GetTickCount - Ticks
    Picture2.Refresh
    picSave.Refresh
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Effect = 0
    LastPath = App.Path
    ColorPickTool = False
    HScroll1.Enabled = False
    Command2.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim ctrl As Control
    
    For Each ctrl In Controls
        If TypeOf ctrl Is PictureBox Then
            ctrl.Cls
            ctrl.Picture = LoadPicture()
            Set ctrl.Picture = Nothing
        End If
    Next
    
End Sub

Private Sub Form_Resize()
    'PleaseDontResize Me, 11805, 8835
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmImageEffect = Nothing
End Sub

Private Sub HScroll1_Change()
    Dim Color As Long
    Dim Ticks As Long
    
    On Error Resume Next
    Ticks = GetTickCount
    Select Case Effect
        Case 1
            GPX_Brightness Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_Brightness picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 2
            GPX_Contrast Picture2.hdc, Picture1.hdc, HScroll1.Value / 100, HScroll1.Value / 100, HScroll1.Value / 100, resp
            GPX_Contrast picSave.hdc, picOri.hdc, HScroll1.Value / 100, HScroll1.Value / 100, HScroll1.Value / 100, Resp2
        Case 3
            GPX_Sharpening Picture2.hdc, Picture1.hdc, HScroll1.Value / 100, resp
            GPX_Sharpening picSave.hdc, picOri.hdc, HScroll1.Value / 100, Resp2
        Case 5
            GPX_Gamma Picture2.hdc, Picture1.hdc, HScroll1.Value / 25, resp
            GPX_Gamma picSave.hdc, picOri.hdc, HScroll1.Value / 25, Resp2
        Case 6
            GPX_GrayScale Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_GrayScale picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 7
            GPX_Invert Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_Invert picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 8
            GPX_ReduceColors Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_ReduceColors picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 10
            GPX_Stamp Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_Stamp picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 12
            GPX_Mosaic Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_Mosaic picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 13
            GPX_BackDropRemoval Picture2.hdc, Picture1.hdc, Color, Not Color, HScroll1.Value, resp
            GPX_BackDropRemoval picSave.hdc, picOri.hdc, Color, Not Color, HScroll1.Value, Resp2
        Case 15
            GPX_AmbientLight Picture2.hdc, Picture1.hdc, Color, HScroll1.Value, resp
            GPX_AmbientLight picSave.hdc, picOri.hdc, Color, HScroll1.Value, Resp2
        Case 16
            GPX_Tone Picture2.hdc, Picture1.hdc, Color, HScroll1.Value, resp
            GPX_Tone picSave.hdc, picOri.hdc, Color, HScroll1.Value, Resp2
        Case 17
            GPX_BackDropRemovalEx Picture2.hdc, Picture1.hdc, Color, Not Color, HScroll1.Value, True, True, True, False, resp
            GPX_BackDropRemovalEx picSave.hdc, picOri.hdc, Color, Not Color, HScroll1.Value, True, True, True, False, Resp2
        Case 19
            GPX_Rock Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_Rock picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 20
            GPX_Emboss Picture2.hdc, Picture1.hdc, HScroll1.Value / 100, resp
            GPX_Emboss picSave.hdc, picOri.hdc, HScroll1.Value / 100, Resp2
        Case 21
            GPX_ColorRandomize Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_ColorRandomize picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 22
            'Label1.Caption = HScroll1.Value - 2
            GPX_RandomicalPoints Picture2.hdc, Picture1.hdc, HScroll1.Value, Color, resp
            GPX_RandomicalPoints picSave.hdc, picOri.hdc, HScroll1.Value, Color, Resp2
        Case 24
            GPX_Shift Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_Shift picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 27
            GPX_Twirl Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_Twirl picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 28
            GPX_Swirl Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_Swirl picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 29
            GPX_Neon Picture2.hdc, Picture1.hdc, HScroll1.Value, 2, resp
            GPX_Neon picSave.hdc, picOri.hdc, HScroll1.Value, 2, Resp2
        Case 30
            GPX_Canvas Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_Canvas picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 31
            GPX_Waves Picture2.hdc, Picture1.hdc, HScroll1.Value, HScroll1.Value, HScroll1.Value, True, resp
            GPX_Waves picSave.hdc, picOri.hdc, HScroll1.Value, HScroll1.Value, HScroll1.Value, True, Resp2
        Case 32
            GPX_DetectBorders Picture2.hdc, Picture1.hdc, HScroll1.Value, Color, Not Color, resp
            GPX_DetectBorders picSave.hdc, picOri.hdc, HScroll1.Value, Color, Not Color, Resp2
        Case 35
            GPX_Saturation Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_Saturation picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 36
            GPX_FindEdges Picture2.hdc, Picture1.hdc, HScroll1.Value, 2, resp
            GPX_FindEdges picSave.hdc, picOri.hdc, HScroll1.Value, 2, Resp2
        Case 37
            GPX_Hue Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_Hue picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 38
            GPX_AlphaBlend Picture2.hdc, Picture3.hdc, Picture1.hdc, HScroll1.Value, resp
        Case 42
            Picture2.Picture = LoadPicture("")
            Picture2.BackColor = vbWhite
            GPX_Tile Picture2.hdc, Picture1.hdc, HScroll1.Value, HScroll1.Value, 6, resp
            
            picSave.Picture = LoadPicture("")
            picSave.BackColor = vbWhite
            GPX_Tile picSave.hdc, picOri.hdc, HScroll1.Value, HScroll1.Value, 6, Resp2
        Case 45
            GPX_AdvancedBlur Picture2.hdc, Picture1.hdc, HScroll1.Value, 25, True, resp
            GPX_AdvancedBlur picSave.hdc, picOri.hdc, HScroll1.Value, 25, True, Resp2
        Case 47
            GPX_MotionBlur Picture2.hdc, Picture1.hdc, HScroll1.Value, 15, resp
            GPX_MotionBlur picSave.hdc, picOri.hdc, HScroll1.Value, 15, Resp2
        Case 48
            GPX_ColorBalance Picture2.hdc, Picture1.hdc, 0, 0, HScroll1.Value, resp
            GPX_ColorBalance picSave.hdc, picOri.hdc, 0, 0, HScroll1.Value, Resp2
        Case 49
            GPX_Fragment Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_Fragment picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 50
            GPX_FarBlur Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_FarBlur picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 51
            GPX_RadialBlur Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_RadialBlur picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 52
            GPX_ZoomBlur Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_ZoomBlur picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 54
            GPX_Fog Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_Fog picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 55
            GPX_MediumTones Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_MediumTones picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 56
            GPX_CircularWaves Picture2.hdc, Picture1.hdc, HScroll1.Value, HScroll1.Value, resp
            GPX_CircularWaves picSave.hdc, picOri.hdc, HScroll1.Value, HScroll1.Value, Resp2
        Case 57
            GPX_CircularWavesEx Picture2.hdc, Picture1.hdc, HScroll1.Value * 4, HScroll1.Value, resp
            GPX_CircularWavesEx picSave.hdc, picOri.hdc, HScroll1.Value * 4, HScroll1.Value, Resp2
        Case 59
            GPX_OilPaint Picture2.hdc, Picture1.hdc, HScroll1.Value / 50, HScroll1.Value, resp
            GPX_OilPaint picSave.hdc, picOri.hdc, HScroll1.Value / 50, HScroll1.Value, Resp2
        Case 61
            GPX_NotePaper Picture2.hdc, Picture1.hdc, HScroll1.Value, 2, 20, 1, Color, Not Color, resp
            GPX_NotePaper picSave.hdc, picOri.hdc, HScroll1.Value, 2, 20, 1, Color, Not Color, Resp2
        Case 62
            GPX_FishEyeEx Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_FishEyeEx picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 64
            GPX_Cilindrical Picture2.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_Cilindrical picSave.hdc, picOri.hdc, HScroll1.Value, Resp2
        Case 65
            GPX_UnsharpMask Picture2.hdc, Picture1.hdc, 2, HScroll1.Value / 2, resp
            GPX_UnsharpMask picSave.hdc, picOri.hdc, 2, HScroll1.Value / 2, Resp2
        Case 66
            GPX_BlockWaves Picture2.hdc, Picture1.hdc, HScroll1.Value, HScroll1.Value / 2, 1, resp
            GPX_BlockWaves picSave.hdc, picOri.hdc, HScroll1.Value, HScroll1.Value / 2, 1, Resp2
        Case 68
            GPX_BlendMode Picture2.hdc, Picture3.hdc, Picture1.hdc, HScroll1.Value, resp
            GPX_BlendMode picSave.hdc, picOri.hdc, Picture1.hdc, HScroll1.Value, Resp2
        Case 69
            GPX_TwirlEx Picture2.hdc, Picture1.hdc, -HScroll1.Value, HScroll1.Value, resp
            GPX_TwirlEx picSave.hdc, picOri.hdc, -HScroll1.Value, HScroll1.Value, Resp2
        Case 70
            GPX_GlassBlendMode Picture2.hdc, Picture3.hdc, Picture1.hdc, HScroll1.Value / 200, 3, resp
            GPX_GlassBlendMode picSave.hdc, picOri.hdc, Picture1.hdc, HScroll1.Value / 200, 3, Resp2
        Case 71
            GPX_Metallic Picture2.hdc, Picture1.hdc, 4, HScroll1.Value, 1, resp
            GPX_Metallic picSave.hdc, picOri.hdc, 4, HScroll1.Value, 1, Resp2
        Case 72
            GPX_Metallic Picture2.hdc, Picture1.hdc, 4, HScroll1.Value, 2, resp
            GPX_Metallic picSave.hdc, picOri.hdc, 4, HScroll1.Value, 2, Resp2
        Case 73
            GPX_Metallic Picture2.hdc, Picture1.hdc, 4, HScroll1.Value, 3, resp
            GPX_Metallic picSave.hdc, picOri.hdc, 4, HScroll1.Value, 3, Resp2
    End Select
    Ticks = GetTickCount - Ticks
    Picture2.Refresh
    picSave.Refresh
    'lblTime.Caption = Ticks & " ms"
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub mnu3D_Click()
    ChangeControls 39, Command2, HScroll1
    Me.Caption = "Special Effects - (3D)"
End Sub

Private Sub mnu4Corners_Click()
    ChangeControls 40, Command2, HScroll1
    Me.Caption = "Special Effects - (4 Corners)"
End Sub

Private Sub mnuAlias_Click()
    ChangeControls 4, Command2, HScroll1
    Me.Caption = "Special Effects - (AntiAlias)"
End Sub

Private Sub mnuAlphaBlend_Click()
'    With CommonDialog1
'        .CancelError = False
'        .DialogTitle = "Open..."
'        .InitDir = LastPath
'        .Filter = "Compatible Image Files|*.bmp;*.jpg;*.emf;*.gif;*.rle;*.wmf"
'        .Flags = cdlOFNHideReadOnly
'        .ShowOpen
'        If Len(.FileName) > 0 Then
'            Picture3.Picture = LoadPicture()
'            Picture3.PaintPicture LoadPicture(.FileName), 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight
'        End If
'        LastPath = Left(.FileName, Len(.FileName) - Len(.FileTitle))
'    End With
'
'    ChangeControls 38, Command2, HScroll1, 0, 255, 0
'    Me.Caption = "Special Effects - (Alpha Blend)"
'    HScroll1_Change
End Sub

Private Sub mnuAmbient_Click()
    ChangeControls 15, Command2, HScroll1, 0, 255, 255
    Me.Caption = "Special Effects - (Ambient Light)"
    HScroll1_Change
End Sub

Private Sub mnuAscii_Click()
    ChangeControls 14, Command2, HScroll1
    Me.Caption = "Special Effects - (Ascii Morph)"
End Sub

Private Sub mnuBack_Click()
    ChangeControls 17, Command2, HScroll1, 0, 255, 0
    Me.Caption = "Special Effects - (Replace Colors)"
    HScroll1_Change
End Sub

Private Sub mnuBackN_Click()
    ChangeControls 13, Command2, HScroll1, 0, 255, 0
    Me.Caption = "Special Effects - (Backdrop Removal)"
    HScroll1_Change
End Sub

Private Sub mnuBlendModes_Click()
'    With CommonDialog1
'        .CancelError = False
'        .DialogTitle = "Open..."
'        .InitDir = LastPath
'        .Filter = "Compatible Image Files|*.bmp;*.jpg;*.emf;*.gif;*.rle;*.wmf"
'        .Flags = cdlOFNHideReadOnly
'        .ShowOpen
'        If Len(.FileName) > 0 Then
'            Picture3.Picture = LoadPicture()
'            Picture3.PaintPicture LoadPicture(.FileName), 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight
'        End If
'        LastPath = Left(.FileName, Len(.FileName) - Len(.FileTitle))
'    End With
'
'    ChangeControls 68, Command2, HScroll1, 0, 24, 0
'    Me.Caption = "Special Effects - (Blend Modes)"
'    HScroll1_Change
End Sub

Private Sub mnuBlockWaves_Click()
    ChangeControls 66, Command2, HScroll1, 0, 20, 0
    Me.Caption = "Special Effects - (Block Waves)"
    HScroll1_Change
End Sub

Private Sub mnuBlur_Click()
    ChangeControls 33, Command2, HScroll1
    Me.Caption = "Special Effects - (Blur)"
End Sub

Private Sub mnuBorders_Click()
    ChangeControls 32, Command2, HScroll1, 0, 255, 0
    Me.Caption = "Special Effects - (Detect Borders)"
    HScroll1_Change
End Sub

Private Sub mnuBrilho_Click()
    ChangeControls 1, Command2, HScroll1, -255, 255, 0
    Me.Caption = "Special Effects - (Brightness)"
    HScroll1_Change
End Sub

Private Sub mnuCanvas_Click()
    ChangeControls 30, Command2, HScroll1, 0, Picture1.ScaleWidth, 0
    Me.Caption = "Special Effects - (Canvas)"
    HScroll1_Change
End Sub

Private Sub mnuCaricature_Click()
    ChangeControls 41, Command2, HScroll1
    Me.Caption = "Special Effects - (Caricature)"
End Sub

Private Sub mnuCilindrical_Click()
    ChangeControls 64, Command2, HScroll1, -30, 30, 0
    Me.Caption = "Special Effects - (Cilindrical)"
    HScroll1_Change
End Sub

Private Sub mnuCircularWaves_Click()
    ChangeControls 56, Command2, HScroll1, 0, 20, 0
    Me.Caption = "Special Effects - (Circular Waves)"
    HScroll1_Change
End Sub

Private Sub mnuCircularWavesEx_Click()
    ChangeControls 57, Command2, HScroll1, 0, 20, 0
    Me.Caption = "Special Effects - (Circular Waves Ex)"
    HScroll1_Change
End Sub

Private Sub mnuColBalance_Click()
    ChangeControls 48, Command2, HScroll1, -255, 255, 0
    Me.Caption = "Special Effects - (Color Balance)"
    HScroll1_Change
End Sub

Private Sub mnuContraste_Click()
    ChangeControls 2, Command2, HScroll1, 0, 255, 100
    Me.Caption = "Special Effects - (Contrast)"
    HScroll1_Change
End Sub

Private Sub mnuDiffuse_Click()
    ChangeControls 18, Command2, HScroll1
    Me.Caption = "Special Effects - (Diffuse)"
End Sub

Private Sub mnuEightColors_Click()
    ChangeControls 9, Command2, HScroll1
    Me.Caption = "Special Effects - (8 colors)"
End Sub

Private Sub mnuEmbEng_Click()
    ChangeControls 20, Command2, HScroll1, -255, 255, 0
    Me.Caption = "Special Effects - (Emboss / Engrave)"
    HScroll1_Change
End Sub

Private Sub mnuFarBlur_Click()
    ChangeControls 50, Command2, HScroll1, 0, 50, 0
    Me.Caption = "Special Effects - (Far Blur)"
    HScroll1_Change
End Sub

Private Sub mnuFindEdges_Click()
    ChangeControls 36, Command2, HScroll1, 0, 5, 0
    Me.Caption = "Special Effects - (Find Edges)"
    HScroll1_Change
End Sub

Private Sub mnuFishEye_Click()
    ChangeControls 26, Command2, HScroll1
    Me.Caption = "Special Effects - (FishEye)"
End Sub

Private Sub mnuFishEyeEx_Click()
    ChangeControls 62, Command2, HScroll1, -255, 255, 0
    Me.Caption = "Special Effects - (FishEye Ex)"
    HScroll1_Change
End Sub

Private Sub mnuFlipB_Click()
    GPX_Flip Picture1.hdc, Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight, 1, 1, resp
    GPX_Flip picOri.hdc, picOri.hdc, picOri.ScaleWidth, picOri.ScaleHeight, 1, 1, Resp2
    Picture1.Refresh
    picOri.Refresh
End Sub

Private Sub mnuFlipH_Click()
    GPX_Flip Picture1.hdc, Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight, 1, 0, resp
    GPX_Flip picOri.hdc, picOri.hdc, picOri.ScaleWidth, picOri.ScaleHeight, 1, 0, Resp2
    Picture1.Refresh
    picOri.Refresh
End Sub

Private Sub mnuFlipV_Click()
    GPX_Flip Picture1.hdc, Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 1, resp
    GPX_Flip picOri.hdc, picOri.hdc, picOri.ScaleWidth, picOri.ScaleHeight, 0, 1, Resp2
    Picture1.Refresh
    picOri.Refresh
End Sub

Private Sub mnuFog_Click()
    ChangeControls 54, Command2, HScroll1, 0, 127, 0
    HScroll1_Change
End Sub

Private Sub mnuFragment_Click()
    ChangeControls 49, Command2, HScroll1, 0, 50, 0
    Me.Caption = "Special Effects - (Fragment)"
    HScroll1_Change
End Sub

Private Sub mnuFrostGlass_Click()
    ChangeControls 60, Command2, HScroll1
    Me.Caption = "Special Effects - (Frost Glass)"
End Sub

Private Sub mnuGamma_Click()
    ChangeControls 5, Command2, HScroll1, 0, 255, 25
    Me.Caption = "Special Effects - (Gamma)"
    HScroll1_Change
End Sub

Private Sub mnuGlassBlendMode_Click()
'    With CommonDialog1
'        .CancelError = False
'        .DialogTitle = "Open..."
'        .InitDir = LastPath
'        .Filter = "Compatible Image Files|*.bmp;*.jpg;*.emf;*.gif;*.rle;*.wmf"
'        .Flags = cdlOFNHideReadOnly
'        .ShowOpen
'        If Len(.FileName) > 0 Then
'            Picture3.Picture = LoadPicture()
'            Picture3.PaintPicture LoadPicture(.FileName), 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight
'        End If
'        LastPath = Left(.FileName, Len(.FileName) - Len(.FileTitle))
'    End With
'
'    ChangeControls 70, Command2, HScroll1, -255, 255, 0
'    Me.Caption = "Special Effects - (Glass Blend Mode)"
'    HScroll1_Change
End Sub

Private Sub mnuGold_Click()
    ChangeControls 72, Command2, HScroll1, 0, 255, 0
    Me.Caption = "Special Effects - (Gold)"
    HScroll1_Change
End Sub

Private Sub mnuGrayScale_Click()
    ChangeControls 6, Command2, HScroll1, -255, 255, 0
    Me.Caption = "Special Effects - (Gray Scale)"
    HScroll1_Change
End Sub

Private Sub mnuHue_Click()
    ChangeControls 37, Command2, HScroll1, 0, 350, 0
    Me.Caption = "Special Effects - (Hue)"
    HScroll1_Change
End Sub

Private Sub mnuIce_Click()
    ChangeControls 73, Command2, HScroll1, 0, 255, 0
    Me.Caption = "Special Effects - (Ice)"
    HScroll1_Change
End Sub

Private Sub mnuInvertion_Click()
    ChangeControls 7, Command2, HScroll1, 0, 255, 0
    Me.Caption = "Special Effects - (Invertion)"
    HScroll1_Change
End Sub

Private Sub mnuMediumTones_Click()
    ChangeControls 55, Command2, HScroll1, 0, 255, 0
    Me.Caption = "Special Effects - (Medium Tones)"
    HScroll1_Change
End Sub

Private Sub mnuMelt_Click()
    ChangeControls 25, Command2, HScroll1
    Me.Caption = "Special Effects - (Melt)"
End Sub

Private Sub mnuMetallic_Click()
    ChangeControls 71, Command2, HScroll1, 0, 255, 0
    Me.Caption = "Special Effects - (Metallic)"
    HScroll1_Change
End Sub

Private Sub mnuMono_Click()
    ChangeControls 10, Command2, HScroll1, 0, 255, 0
    Me.Caption = "Special Effects - (Monochrome)"
    HScroll1_Change
End Sub

Private Sub mnuMoreBlur_Click()
    ChangeControls 45, Command2, HScroll1, 0, 10, 0
    Me.Caption = "Special Effects - (More Blur)"
    HScroll1_Change
End Sub

Private Sub mnuMosaico_Click()
    ChangeControls 12, Command2, HScroll1, 0, 255, 0
    Me.Caption = "Special Effects - (Mosaic)"
    HScroll1_Change
End Sub

Private Sub mnuMotionBlur_Click()
    ChangeControls 47, Command2, HScroll1, 0, 360, 0
    Me.Caption = "Special Effects - (Motion Blur)"
    HScroll1_Change
End Sub

Private Sub mnuNeon_Click()
    ChangeControls 29, Command2, HScroll1, 0, 5, 0
    Me.Caption = "Special Effects - (Neon)"
    HScroll1_Change
End Sub

Private Sub mnuNitidez_Click()
    ChangeControls 3, Command2, HScroll1, 0, 60, 0
    Me.Caption = "Special Effects - (Sharpen)"
    HScroll1_Change
End Sub

Private Sub mnuNoise_Click()
    ChangeControls 21, Command2, HScroll1, -255, 255, 0
    Me.Caption = "Special Effects - (Noise)"
    HScroll1_Change
End Sub

Private Sub mnuNotePaper_Click()
    ChangeControls 61, Command2, HScroll1, 0, 255, 0
    Me.Caption = "Special Effects - (Note Paper)"
    HScroll1_Change
End Sub

Private Sub mnuOilPaint_Click()
    ChangeControls 59, Command2, HScroll1, 0, 255, 0
    Me.Caption = "Special Effects - (OilPaint)"
    HScroll1_Change
End Sub

Private Sub mnuOpen_Click()
    
    Dim Cdlg As New cCommonDialog
    Dim Archivo As String
    Dim Filtro As String
    
    Filtro = ""
    
    If LastPath = "" Then LastPath = App.Path
    
    Filtro = "Bitmap (*.bmp)|*.bmp|"
    Filtro = Filtro & "JPG (*.jpg)|*.jpg|"
    Filtro = Filtro & "JPEG (*.jpeg)|*.jpeg|"
    Filtro = Filtro & "EMF (*.emf)|*.emf|"
    Filtro = Filtro & "GIF (*.gif)|*.gif|"
    Filtro = Filtro & "RLE (*.rle)|*.rle|"
    Filtro = Filtro & "WMF (*.wmf)|*.wmf|"
    Filtro = Filtro & "All files (*.*)|*.*"
        
    If Not Cdlg.VBGetOpenFileName(Archivo, , , , , , Filtro, , LastPath, "Open File ...", "jpeg") Then
        Exit Sub
    End If
    
    On Error GoTo ErrorOpenFile
    
    Picture1.Picture = LoadPicture()
    Picture2.Picture = LoadPicture()
    Picture1.PaintPicture LoadPicture(Archivo), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
    Picture2.PaintPicture LoadPicture(Archivo), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
    
    picOri.Picture = LoadPicture(Archivo)
    picSave.Picture = LoadPicture(Archivo)
    
    LastPath = PathArchivo(Archivo)
    
    mnuSave.Enabled = True
    mnuEffects.Enabled = True
    mnuImage.Enabled = True
    'Command2.Enabled = True
    cmdCopyImage.Enabled = True
    Command1.Enabled = True
    
    Set Cdlg = Nothing
    
    Exit Sub
    
ErrorOpenFile:
    MsgBox "Failed to open selected file. " & vbNewLine & vbNewLine & "Description : " & Err.Description & " Number : " & Err.Number, vbCritical
    mnuSave.Enabled = False
    mnuEffects.Enabled = False
    mnuImage.Enabled = False
    Command2.Enabled = False
    cmdCopyImage.Enabled = False
    Command1.Enabled = False
    
    Set Cdlg = Nothing
    
End Sub
Private Sub mnuPolar_Click()
    ChangeControls 58, Command2, HScroll1
    Me.Caption = "Special Effects - (Polar Coordinates)"
End Sub

Private Sub mnuRadialBlur_Click()
    ChangeControls 51, Command2, HScroll1, 0, 30, 0
    Me.Caption = "Special Effects - (Radial Blur)"
    HScroll1_Change
End Sub

Private Sub mnuRainDrop_Click()
    ChangeControls 63, Command2, HScroll1
    Me.Caption = "Special Effects - (Rain Drop)"
End Sub

Private Sub mnuRandomPoints_Click()
    ChangeControls 22, Command2, HScroll1, 2, 102, 2
    Me.Caption = "Special Effects - (Random Points)"
    HScroll1_Change
End Sub

Private Sub mnuReduceColors_Click()
    ChangeControls 8, Command2, HScroll1, 0, 255, 0
    Me.Caption = "Special Effects - (Reduce Colors)"
    HScroll1_Change
End Sub

Private Sub mnuRelief_Click()
    ChangeControls 34, Command2, HScroll1
    Me.Caption = "Special Effects - (Relief)"
End Sub

Private Sub mnuRock_Click()
    ChangeControls 19, Command2, HScroll1, 0, 6, 0
    Me.Caption = "Special Effects - (Rock)"
    HScroll1_Change
End Sub

Private Sub mnuRoll_Click()
    ChangeControls 43, Command2, HScroll1
    Me.Caption = "Special Effects - (Roll)"
End Sub

Private Sub mnuSair_Click()
    Unload Me
End Sub

Private Sub mnuSaturation_Click()
    ChangeControls 35, Command2, HScroll1, -255, 512, 0
    HScroll1_Change
End Sub

Private Sub mnuSave_Click()
    
    Dim Cdlg As New cCommonDialog
    Dim Archivo As String
    Dim Filtro As String
    
    Filtro = ""
    
    If LastPath = "" Then LastPath = App.Path
    
    Filtro = "Jpeg (*.jpeg)|*.jpeg|"
    Filtro = Filtro & "JPG (*.jpg)|*.jpg|"
            
    If Not Cdlg.VBGetSaveFileName(Archivo, , , Filtro, , LastPath, "Save As ...", "jpg") Then
        Exit Sub
    End If
    
    Set Cdlg = Nothing
    
    On Error GoTo ErrorSaveFile
    
    Dim locImageLoad As cImage
    Dim locImageSave As cImage
    Dim MyPicLoad As StdPicture
    Dim MyPicSave As StdPicture
    Dim locJpeg As cJpeg
   
    SavePicture picSave.Image, Archivo

    Set MyPicSave = LoadPicture(Archivo)
    Set locImageSave = New cImage
    locImageSave.CopyStdPicture MyPicSave
    
    PaintImage locImageSave
    
    Set locJpeg = New cJpeg
    locJpeg.Quality = 75
    locJpeg.SetSamplingFrequencies 2, 2, 1, 1, 1, 1
    locJpeg.SampleHDC locImageSave.hdc, locImageSave.Width, locImageSave.Height
    
    DeleteFile Archivo
    
    locJpeg.SaveFile Archivo
      
    Set MyPicLoad = Nothing
    Set MyPicSave = Nothing
    Set locImageLoad = Nothing
    Set locImageSave = Nothing
    Set locJpeg = Nothing
      
    Exit Sub
    
ErrorSaveFile:
    MsgBox "Failed to save current file." & vbNewLine & vbNewLine & "Description :" & Err.Description & " Number :" & Err.Number, vbCritical
    Set Cdlg = Nothing
    
End Sub

Private Sub PaintImage(TheImage As cImage)
    If ObjPtr(TheImage) = 0 Then
        picSave.Cls
    Else
        TheImage.PaintHDC picSave.hdc
        picSave.Refresh
    End If
End Sub
Private Sub mnuSepia_Click()
    ChangeControls 11, Command2, HScroll1
End Sub

Private Sub mnuShift_Click()
    ChangeControls 24, Command2, HScroll1, 0, 255, 0
    Me.Caption = "Special Effects - (Shift)"
    HScroll1_Change
End Sub

Private Sub mnuSmartBlur_Click()
    ChangeControls 44, Command2, HScroll1
    Me.Caption = "Special Effects - (Smart Blur)"
    ''Label1.Caption = ""
End Sub

Private Sub mnuSoftnerBlur_Click()
    ChangeControls 46, Command2, HScroll1
    Me.Caption = "Special Effects - (Softner Blur)"
    'Label1.Caption = ""
End Sub

Private Sub mnuSol_Click()
    ChangeControls 23, Command2, HScroll1
    Me.Caption = "Special Effects - (Solarize)"
    'Label1.Caption = ""
End Sub

Private Sub mnuStretchHisto_Click()
    ChangeControls 67, Command2, HScroll1
    Me.Caption = "Special Effects - (Stretch Histogram)"
    'Label1.Caption = ""
End Sub

Private Sub mnuSwirl_Click()
    ChangeControls 28, Command2, HScroll1, -255, 255, 0
    Me.Caption = "Special Effects - (Swirl)"
    HScroll1_Change
End Sub

Private Sub mnuTile_Click()
    ChangeControls 42, Command2, HScroll1, 0, 100, 0
    Me.Caption = "Special Effects - (Tile)"
    HScroll1_Change
End Sub

Private Sub mnuTone_Click()
    ChangeControls 16, Command2, HScroll1, 0, 255, 0
    Me.Caption = "Special Effects - (Tone)"
    HScroll1_Change
End Sub


Private Sub mnuTwirl_Click()
    ChangeControls 27, Command2, HScroll1, -100, 100, 0
    Me.Caption = "Special Effects - (Twirl)"
    HScroll1_Change
End Sub

Private Sub mnuTwirlEx_Click()
    ChangeControls 69, Command2, HScroll1, -100, 100, 0
    Me.Caption = "Special Effects - (Twirl Ex)"
    HScroll1_Change
End Sub

Private Sub mnuUnsharpMask_Click()
    ChangeControls 65, Command2, HScroll1, 0, 10, 0
    Me.Caption = "Special Effects - (Unsharp Mask)"
    HScroll1_Change
End Sub

Private Sub mnuWave_Click()
    ChangeControls 31, Command2, HScroll1, 0, 20, 0
    Me.Caption = "Special Effects - (Waves)"
    HScroll1_Change
End Sub

Private Sub mnuWebColors_Click()
    ChangeControls 53, Command2, HScroll1
    Me.Caption = "Special Effects - (Web Colors)"
    'Label1.Caption = ""
End Sub

Private Sub mnuZoomBlur_Click()
    ChangeControls 52, Command2, HScroll1, 0, 200, 0
    Me.Caption = "Special Effects - (Zoom Blur)"
    HScroll1_Change
End Sub

