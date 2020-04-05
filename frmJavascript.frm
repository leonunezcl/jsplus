VERSION 5.00
Object = "{462EF1F4-16AF-444F-9DEE-F41BEBEC2FD8}#1.1#0"; "vbalODCL6.ocx"
Object = "{3A709943-58E7-4A77-9E5B-D5333AC98098}#1.2#0"; "vbalToolboxBar6.ocx"
Begin VB.Form frmJsExp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Javascript Browser"
   ClientHeight    =   5640
   ClientLeft      =   6885
   ClientTop       =   3105
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   376
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   197
   ShowInTaskbar   =   0   'False
   Begin vbalToolboxBar6.vbalToolBoxBarCtl tboJs 
      Height          =   2760
      Left            =   30
      TabIndex        =   1
      Top             =   2250
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4868
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ODCboLst6.OwnerDrawComboList lstObj 
      Height          =   1935
      Left            =   -15
      TabIndex        =   0
      Top             =   225
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3413
      ExtendedUI      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   4
      MaxLength       =   0
   End
End
Attribute VB_Name = "frmJsExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Resize2()
    Form_Resize
End Sub


Private Sub Form_Load()

    Dim archivo As String
    Dim v As Variant
    Dim j As Integer
    Dim ele As String
    
    archivo = IniPath
    v = Util.LeeIni(archivo, "objects", "num")
    
    If Len(v) > 0 Then
        For j = 1 To v
            ele = Util.LeeIni(archivo, "objects", "ele" & j)
            If Len(ele) > 0 Then
                lstObj.AddItem ele
            End If
        Next j
    End If
    
    Dim cBar As cToolBoxBar
    
    Set cBar = tboJs.Bars.Add("k1", , "Properties")
    Set cBar = tboJs.Bars.Add("k2", , "Methods")
    Set cBar = tboJs.Bars.Add("k3", , "Events")

End Sub


Private Sub Form_Resize()

    If WindowState <> vbMinimized Then
        On Error Resume Next
        Dim Top As Integer
        Top = 17
        lstObj.Move 0, Top, ScaleWidth, ScaleHeight / 2
        tboJs.Move 0, (lstObj.Height + 1 + Top), ScaleWidth, (ScaleHeight / 2) - Top
        Err = 0
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmJsExp = Nothing
End Sub


