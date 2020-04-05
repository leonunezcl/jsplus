VERSION 5.00
Begin VB.UserControl vbsColorName 
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3555
   ScaleHeight     =   5460
   ScaleWidth      =   3555
   Begin VB.PictureBox picScroll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5010
      Left            =   0
      ScaleHeight     =   5010
      ScaleWidth      =   3540
      TabIndex        =   4
      Top             =   390
      Width           =   3540
      Begin VB.VScrollBar vsPreview 
         Height          =   1215
         Left            =   3270
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picBk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   0
         ScaleHeight     =   240
         ScaleWidth      =   3390
         TabIndex        =   5
         Top             =   0
         Width           =   3390
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   2385
            ScaleHeight     =   165
            ScaleWidth      =   840
            TabIndex        =   6
            Top             =   15
            Width           =   870
         End
         Begin VB.Label lblColorName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "LBLCOLOR"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   8
            Top             =   15
            Width           =   840
         End
         Begin VB.Label lblColorHex 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "COLORHEX"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   1275
            TabIndex        =   7
            Top             =   15
            Width           =   885
         End
      End
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   3540
      TabIndex        =   0
      Top             =   0
      Width           =   3540
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Color Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   3
         Top             =   75
         Width           =   990
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Color HEX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1275
         TabIndex        =   2
         Top             =   75
         Width           =   885
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Web Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2370
         TabIndex        =   1
         Top             =   90
         Width           =   900
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   3240
         Y1              =   315
         Y2              =   315
      End
   End
End
Attribute VB_Name = "vbsColorName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ColorString As String
Private ColorNameUC As String
Private RValue As String
Private GValue As String
Private BValue As String
Private StrLength As Integer
Private TZP, RP, GP, BP As Integer
Private ColorS(89) As String

Public Event SelectColor(ByVal ColorName As String, ByVal ColorValue As String, ByVal RedValue As String, ByVal GreenValue As String, ByVal BlueValue As String)

'propiedades
Private m_ColorName As String
Private m_ColorValue As String
Private m_RedValue As String
Private m_GreenValue As String
Private m_BlueValue As String

Private Sub lblColorHex_Click(Index As Integer)
    picColor_Click Index
End Sub

Private Sub lblColorName_Click(Index As Integer)
    picColor_Click Index
End Sub

Private Sub picColor_Click(Index As Integer)

    m_ColorName = Explode(picColor(Index).Tag, 1, "|")
    m_ColorValue = Explode(picColor(Index).Tag, 2, "|")
    m_RedValue = Explode(picColor(Index).Tag, 3, "|")
    m_GreenValue = Explode(picColor(Index).Tag, 4, "|")
    m_BlueValue = Explode(picColor(Index).Tag, 5, "|")
    
    RaiseEvent SelectColor(m_ColorName, m_ColorValue, m_RedValue, m_GreenValue, m_BlueValue)
    
End Sub

Private Sub UserControl_Initialize()

    Dim k As Integer
    
    ColorS(0) = "Black;RC0GC0BC0"
    ColorS(1) = "White;RC255GC255BC255"
    ColorS(2) = "Blue;RC0GC0BC255"
    ColorS(3) = "Cyan;RC0GC255BC255"
    ColorS(4) = "Green;RC0GC255BC0"
    ColorS(5) = "Yellow;RC255GC255BC0"
    ColorS(6) = "Red;RC255GC0BC0"
    ColorS(7) = "Magenta;RC255GC0BC255"
    ColorS(8) = "Purple;RC153GC0BC204"
    ColorS(9) = "Orange;RC255GC102BC0"
    ColorS(10) = "Pink;RC255GC153BC204"
    ColorS(11) = "Dark Brown;RC102GC51BC51"
    ColorS(12) = "Powder Blue;RC204GC204BC255"
    ColorS(13) = "Pastel Blue;RC153GC153BC255"
    ColorS(14) = "Baby Blue;RC102GC153BC255"
    ColorS(15) = "Electric Blue;RC102GC102BC255"
    ColorS(16) = "Twilight Blue;RC102GC102BC204"
    ColorS(17) = "Navy Blue;RC0GC51BC153"
    ColorS(18) = "Deep Navy Blue;RC0GC0BC102"
    ColorS(19) = "Desert Blue;RC51GC102BC153"
    ColorS(20) = "Sky Blue;RC0GC204BC255"
    ColorS(21) = "Ice Blue;RC153GC255BC255"
    ColorS(22) = "Light BlueGreen;RC153GC204BC204"
    ColorS(23) = "Ocean Green;RC102GC153BC153"
    ColorS(24) = "Moss Green;RC51GC102BC102"
    ColorS(25) = "Dark Green;RC0GC51BC51"
    ColorS(26) = "Forest Green;RC0GC102BC51"
    ColorS(27) = "Grass Green;RC0GC153BC51"
    ColorS(28) = "Kentucky Green;RC51GC153BC102"
    ColorS(29) = "Light Green;RC51GC204BC102"
    ColorS(30) = "Spring Green;RC51GC204BC51"
    ColorS(31) = "Turquoise;RC102GC255BC204"
    ColorS(32) = "Sea Green;RC51GC204BC153"
    ColorS(33) = "Faded Green;RC153GC204BC153"
    ColorS(34) = "Ghost Green;RC204GC255BC204"
    ColorS(35) = "Mint Green;RC153GC255BC153"
    ColorS(36) = "Army Green;RC102GC153BC102"
    ColorS(37) = "Avocado Green;RC102GC153BC51"
    ColorS(38) = "Martian Green;RC153GC204BC51"
    ColorS(39) = "Dull Green;RC153GC204BC102"
    ColorS(40) = "Chartreuse;RC153GC255BC0"
    ColorS(41) = "Moon Green;RC204GC255BC102"
    ColorS(42) = "Murky Green;RC51GC51BC0"
    ColorS(43) = "Olive Drab;RC102GC102BC51"
    ColorS(44) = "Khaki;RC153GC153BC102"
    ColorS(45) = "Olive;RC153GC153BC51"
    ColorS(46) = "Banana Yellow;RC204GC204BC51"
    ColorS(47) = "Light Yellow;RC255GC255BC102"
    ColorS(48) = "Chalk;RC255GC255BC153"
    ColorS(49) = "Pale Yellow;RC255GC255BC204"
    ColorS(50) = "Brown;RC153GC102BC51"
    ColorS(51) = "Red Brown;RC204GC102BC51"
    ColorS(52) = "Gold;RC204GC153BC51"
    ColorS(53) = "Autumn Orange;RC255GC102BC51"
    ColorS(54) = "Light Orange;RC255GC153BC51"
    ColorS(55) = "Peach;RC255GC153BC102"
    ColorS(56) = "Deep Yellow;RC255GC204BC0"
    ColorS(57) = "Sand;RC255GC204BC153"
    ColorS(58) = "Walnut;RC102GC51BC0"
    ColorS(59) = "Ruby Red;RC153GC0BC0"
    ColorS(60) = "Brick Red;RC204GC51BC0"
    ColorS(61) = "Tropical Pink;RC255GC102BC102"
    ColorS(62) = "Soft Pink;RC255GC153BC153"
    ColorS(63) = "Faded Pink;RC255GC204BC204"
    ColorS(64) = "Crimson;RC153GC51BC102"
    ColorS(65) = "Regal Red;RC204GC51BC102"
    ColorS(66) = "Deep Rose;RC204GC51BC153"
    ColorS(67) = "Neon Red;RC255GC0BC102"
    ColorS(68) = "Deep Pink;RC255GC102BC153"
    ColorS(69) = "Hot Pink;RC255GC51BC153"
    ColorS(70) = "Dusty Rose;RC204GC102BC153"
    ColorS(71) = "Plum;RC102GC0BC102"
    ColorS(72) = "Deep Violet;RC153GC0BC153"
    ColorS(73) = "Light Violet;RC255GC153BC255"
    ColorS(74) = "Violet;RC204GC102BC204"
    ColorS(75) = "Dusty Plum;RC153GC102BC153"
    ColorS(76) = "Pale Purple;RC204GC153BC204"
    ColorS(77) = "Majestic Purple;RC153GC51BC204"
    ColorS(78) = "Neon Purple;RC204GC51BC255"
    ColorS(79) = "Light Purple;RC204GC102BC255"
    ColorS(80) = "Twilight Violet;RC153GC102BC204"
    ColorS(81) = "Easter Purple;RC204GC153BC255"
    ColorS(82) = "Deep Purple;RC51GC0BC102"
    ColorS(83) = "Grape;RC102GC51BC153"
    ColorS(84) = "Blue Violet;RC153GC102BC255"
    ColorS(85) = "Blue Purple;RC153GC0BC255"
    ColorS(86) = "Deep River;RC102GC0BC204"
    ColorS(87) = "Deep Azure;RC102GC51BC255"
    ColorS(88) = "Storm Blue;RC51GC0BC153"
    ColorS(89) = "Deep Blue;RC51GC0BC204"
    
    UserControl.ScaleMode = vbPixels
    picBk.ScaleMode = vbPixels
    picScroll.ScaleMode = vbPixels
 
    For k = 0 To UBound(ColorS)
        CalculateColors (k)
      
        If k > 0 Then
            Load lblColorName(k)
            
            lblColorName(k).Height = lblColorName(k - 1).Height
            lblColorName(k).Width = lblColorName(k - 1).Width
            lblColorName(k).Top = lblColorName(k - 1).Top + lblColorName(k - 1).Height
            lblColorName(k).Visible = True
            
            Load lblColorHex(k)
            lblColorHex(k).Height = lblColorHex(k - 1).Height
            lblColorHex(k).Width = lblColorHex(k - 1).Width
            lblColorHex(k).Top = lblColorHex(k - 1).Top + lblColorHex(k - 1).Height
            lblColorHex(k).Visible = True
            
            Load picColor(k)
            picColor(k).Height = picColor(k - 1).Height
            picColor(k).Width = picColor(k - 1).Width
            picColor(k).Top = picColor(k - 1).Top + picColor(k - 1).Height
            picColor(k).Visible = True
            
        End If
        
        lblColorName(k).Caption = ColorNameUC
        lblColorHex(k).Caption = "#" & pvHex(RValue) & pvHex(GValue) & pvHex(BValue)
        picColor(k).BackColor = RGB(RValue, GValue, BValue)
        picColor(k).Tag = lblColorName(k).Caption & "|" & lblColorHex(k).Caption & "|" & RValue & "|" & GValue & "|" & BValue
        picBk.Height = picBk.Height + 240
    Next k

    vsPreview.Max = picScroll.Height + 506
    vsPreview.LargeChange = vsPreview.Max \ 10
    vsPreview.SmallChange = vsPreview.Max \ 25
  
End Sub
Private Function pvHex(ByVal lValue As Long, Optional lCount As Long = 2) As String
'--- convert hex and pad with zeroes
    pvHex = VBA.Right(String(lCount, "0") & Hex(lValue), lCount)
End Function
Private Sub CalculateColors(ByVal CI As Integer)
    
    ColorString = ColorS(CI)
    
    StrLength = Len(ColorString)
    TZP = InStr(1, ColorString, ";")
    RP = InStr(1, ColorString, "RC")
    GP = InStr(1, ColorString, "GC")
    BP = InStr(1, ColorString, "BC")
    
    ColorNameUC = Left(ColorString, TZP - 1)
    RValue = Mid(ColorString, RP + 2, GP - RP - 2)
    GValue = Mid(ColorString, GP + 2, BP - GP - 2)
    BValue = Mid(ColorString, BP + 2, Len(ColorString) - BP - 1)

End Sub

Private Sub UserControl_Resize()
    
    'picScroll.Height = UserControl.Height - picHead.Height
    vsPreview.Height = picScroll.Height - 1
    
    'vsPreview.Max = picScroll.Height + 506
    'vsPreview.LargeChange = vsPreview.Max \ 10
    'vsPreview.SmallChange = vsPreview.Max \ 25
    
End Sub


Private Sub vsPreview_Change()
    picBk.Top = -vsPreview.Value '* 14.4
End Sub



Public Property Get ColorName() As String
    ColorName = m_ColorName
End Property

Public Property Let ColorName(ByVal pColorName As String)
    m_ColorName = pColorName
End Property

Public Property Get ColorValue() As String
    ColorValue = m_ColorValue
End Property

Public Property Let ColorValue(ByVal pColorValue As String)
    m_ColorValue = pColorValue
End Property

Public Property Get RedValue() As String
    RedValue = m_RedValue
End Property

Public Property Let RedValue(ByVal pRedValue As String)
    m_RedValue = pRedValue
End Property

Public Property Get GreenValue() As String
    GreenValue = m_GreenValue
End Property

Public Property Let GreenValue(ByVal pGreenValue As String)
    m_GreenValue = pGreenValue
End Property

Public Property Get BlueValue() As String
    BlueValue = m_BlueValue
End Property

Public Property Let BlueValue(ByVal pBlueValue As String)
    m_BlueValue = pBlueValue
End Property
