VERSION 5.00
Begin VB.UserControl ctlFrame 
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   ControlContainer=   -1  'True
   ScaleHeight     =   1695
   ScaleWidth      =   3900
   Begin VB.Line LineLeft 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   375
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   2505
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line LineBottom 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   2505
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line LineRight 
      BorderColor     =   &H80000010&
      X1              =   2520
      X2              =   2520
      Y1              =   0
      Y2              =   420
   End
End
Attribute VB_Name = "ctlFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event Resize()
Private Sub UserControl_Resize()
    LineLeft.Y2 = UserControl.ScaleHeight
    LineTop.X2 = UserControl.ScaleWidth - 15
    LineRight.Y2 = UserControl.ScaleHeight
    LineRight.X1 = UserControl.ScaleWidth - 15
    LineRight.X2 = UserControl.ScaleWidth - 15
    LineBottom.X2 = UserControl.ScaleWidth
    LineBottom.Y2 = UserControl.ScaleHeight - 15
    LineBottom.Y1 = UserControl.ScaleHeight - 15
    RaiseEvent Resize
End Sub
