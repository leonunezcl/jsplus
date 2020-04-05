VERSION 5.00
Begin VB.UserControl HeaderPicture 
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   PropertyPages   =   "HeaderPicture.ctx":0000
   ScaleHeight     =   255
   ScaleWidth      =   2595
   ToolboxBitmap   =   "HeaderPicture.ctx":0023
   Begin VB.PictureBox picHeader 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   0
      Width           =   2595
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblCaption"
         Height          =   195
         Left            =   855
         TabIndex        =   1
         Top             =   0
         Width           =   705
      End
      Begin VB.Image imgIcon 
         Height          =   255
         Left            =   120
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Image imgInit 
      Height          =   480
      Left            =   1800
      Picture         =   "HeaderPicture.ctx":0335
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "HeaderPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'By Jim K in April, 2004.
'UserControl Portions of this code Peter Hart, April 2004.
'Use PictureBox as guiding info headers.

'Properties
'----------
' Alignment
' Caption
' Font
' FontBold
' FontSize
' FontColor
' Gradient            (Vertical/Horizontal)
' GradientStart       (Color)
' GradientFinish      (Color)
' GradientFinishStyle (Transparent/Opaque)
' MultiLine
' Picture
' Shape               (Rectangle/Rounded/RoundedTop)

Private sCaption As String
Private SGC As Long
Private EGC As Long
Private bMulti As Boolean         ' Multiline boolean

Public Enum Style                 ' Border Shape
    Rectangle = 0
    Rounded = 1
    RoundedTop = 2
    RoundedBottom = 3
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Rectangle, Rounded, RoundedTop, RoundedBottom
#End If

Public Enum Align                 ' Alignment
     Left = 0
     LeftCenter = 1
     Right = 2
     RightCenter = 3
     Bottom = 4
     Top = 5
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Left, LeftCenter, Right, RightCenter, Center, Bottom, Top
#End If

Public Enum Direction             ' Gradient
    Horizontal = 0
    Vertical = 1
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Horizontal, Vertical
#End If
 
Public Enum BackStyle             ' Gradient Finish Style
    Opaque = 0
    Transparent = 1
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Opaque, Transparent
#End If

Private Type UserControlProps
    Alignement              As Align
    GradientDirection       As Direction
    Shape                   As Style
    GradientBackStyle       As BackStyle
End Type

Private myProps             As UserControlProps  ' cached ctrl properties

Private Declare Function GetSysColor Lib "user32" ( _
                   ByVal nIndex As Long) As Long
                                
Public Property Let Alignment(ByVal newVal As Align)
myProps.Alignement = newVal
ResizeInfoHeader
PropertyChanged "Alignment"
End Property
Public Property Get Alignment() As Align
Alignment = myProps.Alignement
End Property
                                
Public Property Let Caption(str As String)
If bMulti Then
   lblCaption.Caption = Replace(str, "|", vbCrLf)
   ResizeInfoHeader
 Else
   lblCaption.Caption = str
End If
PropertyChanged "Caption"
End Property
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "Caption"
Caption = lblCaption.Caption
End Property
 
Public Property Get Font() As Font
Set Font = lblCaption.Font
End Property
Public Property Set Font(ByVal NewFont As Font)
Set lblCaption.Font = NewFont
PropertyChanged "Font"
End Property

Public Property Get FontBold() As Boolean
FontBold = lblCaption.FontBold
End Property
Public Property Let FontBold(b As Boolean)
lblCaption.FontBold = b
PropertyChanged "FontBold"
End Property

Public Property Get FontSize() As Integer
FontSize = lblCaption.FontSize
End Property
Public Property Let FontSize(i As Integer)
lblCaption.FontSize = i
PropertyChanged "FontSize"
End Property

Public Property Let FontColor(nColor As Ole_Color)
lblCaption.ForeColor = nColor
PropertyChanged "FontColor"
End Property
Public Property Get FontColor() As Ole_Color
FontColor = lblCaption.ForeColor
End Property

Public Property Let Gradient(Styles As Direction)
'Gradient Direction
myProps.GradientDirection = Styles
DrawInfoHeader
PropertyChanged "Gradient"
End Property
Public Property Get Gradient() As Direction
Gradient = myProps.GradientDirection
End Property

Public Property Let GradientStart(nColor As Ole_Color)
'Starting Gradient Color
SGC = nColor
DrawInfoHeader
PropertyChanged "GradientStart"
End Property
Public Property Get GradientStart() As Ole_Color
GradientStart = SGC
End Property

Public Property Let GradientFinish(nColor As Ole_Color)
'Finishing Gradient Color
EGC = nColor
DrawInfoHeader
PropertyChanged "GradientFinish"
End Property
Public Property Get GradientFinish() As Ole_Color
GradientFinish = EGC
End Property

Public Property Let GradientFinishStyle(Styles As BackStyle)
'Sets whether or not the finish color is opaque
myProps.GradientBackStyle = Styles
DrawInfoHeader
PropertyChanged "GradientFinishStyle"
End Property
Public Property Get GradientFinishStyle() As BackStyle
GradientFinishStyle = myProps.GradientBackStyle
End Property

Public Property Let MultiLine(b As Boolean)
bMulti = b
If b Then
   lblCaption.Caption = Replace(lblCaption.Caption, "|", vbCrLf)
   ResizeInfoHeader
End If
PropertyChanged "MultiLine"
End Property
Public Property Get MultiLine() As Boolean
MultiLine = bMulti
End Property

Public Property Set Picture(ByVal newVal As StdPicture)
Set imgIcon.Picture = newVal
PropertyChanged "Picture"
ResizeInfoHeader
End Property
Public Property Get Picture() As StdPicture
Set Picture = imgIcon.Picture
End Property

Public Property Let Shape(Styles As Style)
myProps.Shape = Styles
DrawInfoHeader
PropertyChanged "Shape"
End Property
Public Property Get Shape() As Style
Shape = myProps.Shape
End Property

Private Sub UserControl_InitProperties()
Alignment = Align.Left
Caption = "HeaderPicture Control v 1.0.3"
Font = UserControl.Parent.Font
FontBold = True
FontSize = 8
FontColor = vbWhite
Gradient = Horizontal
GradientStart = vbBlue
GradientFinish = vbWhite
GradientFinishStyle = Opaque
MultiLine = False
Set Picture = imgInit.Picture
Shape = Rounded
UserControl.BackColor = Parent.BackColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
     Alignment = .ReadProperty("Alignment", Align.Left)
     Caption = .ReadProperty("Caption", "Title Here")
     Font = .ReadProperty("Font", UserControl.Parent.Font)
     FontBold = .ReadProperty("FontBold", True)
     FontSize = .ReadProperty("FontSize", 8)
     FontColor = .ReadProperty("FontColor", vbWhite)
     Gradient = .ReadProperty("Gradient", 0)
     GradientStart = .ReadProperty("GradientStart", vbBlue)
     GradientFinish = .ReadProperty("GradientFinish", vbWhite)
     GradientFinishStyle = .ReadProperty("GradientFinishStyle", 0)
     MultiLine = .ReadProperty("MultiLine", False)
     Set Picture = .ReadProperty("Picture", Picture)
     Shape = .ReadProperty("Shape", 1)
End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
     .WriteProperty "Alignment", Alignment
     .WriteProperty "Caption", Caption, "Title Here"
     .WriteProperty "Font", Font
     .WriteProperty "FontBold", FontBold
     .WriteProperty "FontSize", FontSize
     .WriteProperty "FontColor", FontColor, vbWhite
     .WriteProperty "Gradient", Gradient, 0
     .WriteProperty "GradientStart", SGC, vbBlue
     .WriteProperty "GradientFinish", GradientFinish, vbWhite
     .WriteProperty "GradientFinishStyle", GradientFinishStyle, 0
     .WriteProperty "MultiLine", MultiLine, False
     .WriteProperty "Picture", Picture
     .WriteProperty "Shape", Shape, 1
End With
End Sub

Private Sub UserControl_Resize()
picHeader.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
DrawInfoHeader
ResizeInfoHeader
End Sub

Private Function ConvertColor(tColor As Long) As Long
'Converts VB color constants to real color values
If tColor < 0 Then
   ConvertColor = GetSysColor(tColor And &HFF&)
 Else
   ConvertColor = tColor
End If
End Function

Public Sub DrawInfoHeader()
Dim x1 As Integer
Dim R, G, b, dr, dg, db As Single
Dim y1 As Integer
Dim R1 As Integer
Dim G1 As Integer
Dim B1 As Integer
Dim R2 As Integer
Dim G2 As Integer
Dim B2 As Integer
Dim r3 As Integer
Dim g3 As Integer
Dim b3 As Integer
Dim BC As Long
Dim i As Integer
Dim cnrInt As Integer
Dim EGCCopy As Long

On Error Resume Next
BC = ConvertColor(Parent.BackColor)

If myProps.GradientBackStyle = Transparent Then
   EGCCopy = BC
 Else
   EGCCopy = EGC
End If

DrawGrad SGC, R1, G1, B1
DrawGrad EGCCopy, R2, G2, B2
DrawGrad BC, r3, g3, b3

With picHeader
   
   Select Case myProps.GradientDirection
   
     Case Horizontal
        'Gradient Horizontal
        dr = (R2 - R1) / (.ScaleWidth / 15)
        dg = (G2 - G1) / (.ScaleWidth / 15)
        db = (B2 - B1) / (.ScaleWidth / 15)
        R = R1
        G = G1
        b = B1
        For x1 = 0 To .ScaleWidth Step 15
            picHeader.Line (x1, 0)-(x1, .ScaleHeight), RGB(R, G, b) + &H2000000
            R = R + dr
            G = G + dg
            b = b + db
        Next x1
     
     Case Vertical
        'Gradient Vertical
        dr = (R2 - R1) / (.ScaleHeight / 15)
        dg = (G2 - G1) / (.ScaleHeight / 15)
        db = (B2 - B1) / (.ScaleHeight / 15)
        R = R1
        G = G1
        b = B1
        For y1 = 0 To .ScaleHeight Step 15
            picHeader.Line (0, y1)-(.ScaleWidth, y1), RGB(R, G, b) + &H2000000
            R = R + dr
            G = G + dg
            b = b + db
        Next y1
   End Select
   
    cnrInt = 20
    'Top Corners
    If myProps.Shape = Rounded Or myProps.Shape = RoundedTop Then
       'Left
       i = 15
       For x1 = 0 To cnrInt
           picHeader.Line (x1, 0)-(x1, i), RGB(r3, g3, b3) + &H2000000
           i = i - 1
       Next x1
       'Right
       i = 0
       For x1 = .ScaleWidth - cnrInt To .ScaleWidth
           picHeader.Line (x1, 0)-(x1, i), RGB(r3, g3, b3) + &H2000000
           i = i + 1
       Next x1
    End If
    
    'Bottom Corners
    If myProps.Shape = Rounded Or myProps.Shape = RoundedBottom Then
       'Right
       i = 0
       For x1 = .ScaleWidth - cnrInt To .ScaleWidth
           picHeader.Line (x1, .ScaleHeight - i)-(x1, .ScaleHeight), _
                                               RGB(r3, g3, b3) + &H2000000
           i = i + 1
       Next x1
       'Left
       i = 15
       For x1 = 0 To cnrInt
           picHeader.Line (x1, .ScaleHeight - i)-(x1, .ScaleHeight), _
                                               RGB(r3, g3, b3) + &H2000000
           i = i - 1
       Next x1
    End If

   .Refresh
End With
End Sub
    
Function DrawGrad(Color As Long, R As Integer, G As Integer, b As Integer)
R = Color Mod 256 'red
G = (Color \ 256) Mod 256 'green
b = Color \ 65536 'blue
End Function

Private Sub ResizeInfoHeader()
Dim hfCtrlHht As Integer
Dim hfLblHht As Integer
Dim hfImgHht As Integer
Dim Buf As Integer

Buf = 120
hfCtrlHht = UserControl.Height / 2
hfLblHht = lblCaption.Height / 2
hfImgHht = imgIcon.Height / 2

Select Case myProps.Alignement
  Case Left
    lblCaption.Alignment = vbLeftJustify
    If imgIcon.Picture = 0 Then
       lblCaption.Move Buf, hfCtrlHht - hfLblHht, UserControl.Width - (Buf * 2)
     Else
       imgIcon.Move Buf, hfCtrlHht - hfImgHht
       lblCaption.Move imgIcon.Left + imgIcon.Width + Buf, hfCtrlHht - hfLblHht, UserControl.Width - (Buf * 2)
    End If
  
  Case LeftCenter
    lblCaption.Alignment = vbCenter
    If imgIcon.Picture = 0 Then
       lblCaption.Move Buf, hfCtrlHht - hfLblHht, UserControl.Width - (Buf * 2)
     Else
       imgIcon.Move Buf, hfCtrlHht - hfImgHht
       lblCaption.Move imgIcon.Left + imgIcon.Width + Buf, hfCtrlHht - hfLblHht, UserControl.Width - ((imgIcon.Left + imgIcon.Width + Buf) * 2)
    End If
  
  Case Right
    lblCaption.Alignment = vbRightJustify
    If imgIcon.Picture = 0 Then
       lblCaption.Move Buf, hfCtrlHht - hfLblHht, UserControl.Width - (Buf * 2)
     Else
       imgIcon.Move UserControl.Width - imgIcon.Width - Buf, hfCtrlHht - hfImgHht
       lblCaption.Move imgIcon.Left - lblCaption.Width - Buf, hfCtrlHht - hfLblHht
    End If
  
  Case RightCenter
    lblCaption.Alignment = vbCenter
    If imgIcon.Picture = 0 Then
       lblCaption.Move Buf, hfCtrlHht - hfLblHht, UserControl.Width - (Buf * 2)
     Else
       imgIcon.Move UserControl.Width - imgIcon.Width - Buf, hfCtrlHht - hfImgHht
       lblCaption.Move Buf + imgIcon.Width, hfCtrlHht - hfLblHht, UserControl.Width - ((imgIcon.Width + Buf) * 2)
    End If
  
  Case Top
    lblCaption.Alignment = vbCenter
    If imgIcon.Picture = 0 Then
       lblCaption.Move Buf, hfCtrlHht - hfLblHht, UserControl.Width - (Buf * 2)
     Else
       imgIcon.Move (UserControl.Width / 2) - (imgIcon.Width / 2), Buf
       lblCaption.Move Buf, imgIcon.Top + imgIcon.Height + 120, UserControl.Width - (Buf * 2)
    End If
  
  Case Bottom
    lblCaption.Alignment = vbCenter
    If imgIcon.Picture = 0 Then
       lblCaption.Move Buf, hfCtrlHht - hfLblHht, UserControl.Width - (Buf * 2)
     Else
       imgIcon.Move (UserControl.Width / 2) - (imgIcon.Width / 2), UserControl.Height - imgIcon.Height - Buf
       lblCaption.Move Buf, imgIcon.Top - lblCaption.Height - 120, UserControl.Width - (Buf * 2)
    End If
End Select
End Sub
