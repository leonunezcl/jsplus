VERSION 5.00
Begin VB.Form frmCPick 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose Color"
   ClientHeight    =   4095
   ClientLeft      =   4845
   ClientTop       =   4650
   ClientWidth     =   4065
   HasDC           =   0   'False
   Icon            =   "frmCPick.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pStan 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   60
      MousePointer    =   2  'Cross
      Picture         =   "frmCPick.frx":058A
      ScaleHeight     =   420
      ScaleWidth      =   1680
      TabIndex        =   11
      Top             =   330
      Width           =   1740
      Begin VB.Shape shpTmp 
         BorderStyle     =   3  'Dot
         Height          =   210
         Left            =   0
         Top             =   0
         Width           =   210
      End
      Begin VB.Shape shpMov 
         BorderWidth     =   2
         Height          =   210
         Left            =   0
         Top             =   0
         Width           =   210
      End
   End
   Begin VB.PictureBox pHolder 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2865
      Left            =   60
      MousePointer    =   2  'Cross
      Picture         =   "frmCPick.frx":0C2C
      ScaleHeight     =   2805
      ScaleWidth      =   2625
      TabIndex        =   10
      Top             =   1140
      Width           =   2685
      Begin VB.Line lY 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3005
      End
      Begin VB.Line lX 
         X1              =   0
         X2              =   2700
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lY2 
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3005
      End
      Begin VB.Line lX2 
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         X1              =   0
         X2              =   2700
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox pPrev2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   2880
      ScaleHeight     =   480
      ScaleWidth      =   840
      TabIndex        =   5
      Top             =   3330
      Width           =   870
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   660
      Width           =   1005
   End
   Begin VB.PictureBox pPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   2880
      ScaleHeight     =   480
      ScaleWidth      =   840
      TabIndex        =   3
      Top             =   2475
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   195
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Selected"
      Height          =   195
      Index           =   1
      Left            =   2880
      TabIndex        =   9
      Top             =   2250
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Current"
      Height          =   195
      Index           =   0
      Left            =   2880
      TabIndex        =   8
      Top             =   3105
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Standard Colors"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   105
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Palette Colors"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   945
      Width           =   975
   End
   Begin VB.Label lbHex 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "#FFFFFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   1
      Top             =   1905
      Width           =   780
   End
   Begin VB.Label lbRGB 
      Caption         =   "Red:              Green:         Blue:"
      Height          =   690
      Left            =   2880
      TabIndex        =   0
      Top             =   1170
      Width           =   1050
   End
End
Attribute VB_Name = "frmCPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Sub Command1_Click()

gCodeColor = lbHex.Caption
gSelectColor = pPrev.BackColor
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCPick = Nothing
End Sub


Private Sub pHolder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim cl As Long
If Not lX2.Visible Then lX2.Visible = True
If Not lY2.Visible Then lY2.Visible = True
If shpTmp.Visible Then shpTmp.Visible = False
cl = GetPixel(pHolder.hdc, x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY)
pPrev2.BackColor = cl
lY2.x1 = x
lY2.x2 = x
lX2.y1 = y
lX2.y2 = y
If Button = 1 Then
lbRGB.Caption = GetClrRGBVal(cl)
lbHex.Caption = GetHexVal(cl)
End If
End Sub

Function GetHexVal(Color As Long) As String
Dim R, G, B
R = Hex(Color Mod 256)
G = Hex((Color \ 256) Mod 256)
B = Hex(Color \ 65536)
If Len(R) < 2 Then R = "0" & R
If Len(G) < 2 Then G = "0" & G
If Len(B) < 2 Then B = "0" & B
GetHexVal = "#" & R & G & B
End Function

Function GetClrRGBVal(Color As Long, Optional delim As String = vbCrLf) As String
Dim R, G, B
R = Color Mod 256
G = (Color \ 256) Mod 256
B = Color \ 65536
GetClrRGBVal = "Red: " & R & delim & "Green: " & G & delim & "Blue: " & B
End Function
Private Sub pHolder_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim cl As Long
cl = GetPixel(pHolder.hdc, x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY)
lbRGB.Caption = GetClrRGBVal(cl)
lbHex.Caption = GetHexVal(cl)
pPrev.BackColor = cl
lY.x1 = x
lY.x2 = x
lX.y1 = y
lX.y2 = y
End Sub

Private Sub pStan_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
shpTmp.Left = (x \ 210) * 210 'rounded
shpTmp.Top = (y \ 210) * 210 'rounded
pPrev2.BackColor = GetPixel(pStan.hdc, x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
If lX2.Visible Then lX2.Visible = False
If lY2.Visible Then lY2.Visible = False
If Not shpTmp.Visible Then shpTmp.Visible = True
End Sub

Private Sub pStan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
shpMov.Left = (x \ 210) * 210 'rounded
shpMov.Top = (y \ 210) * 210 'rounded
pStan.Refresh
Dim Color As Long
Color = GetPixel(pStan.hdc, x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
lbHex.Caption = GetHexVal(Color)
lbRGB.Caption = GetClrRGBVal(Color)
pPrev.BackColor = Color
End Sub

