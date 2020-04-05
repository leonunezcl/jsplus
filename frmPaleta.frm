VERSION 5.00
Object = "{246E535D-09D2-4109-80DA-2FF183F4D185}#2.1#0"; "colorpick.ocx"
Object = "{4A3A29A4-F2E3-11D3-B06C-00500427A693}#4.0#0"; "vbalDDFm6.ocx"
Begin VB.Form frmPalette 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Palette Toolbox"
   ClientHeight    =   1275
   ClientLeft      =   4290
   ClientTop       =   5205
   ClientWidth     =   2025
   Icon            =   "frmPaleta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   2025
   ShowInTaskbar   =   0   'False
   Begin ColorPick.ClrPicker ClrPicker1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   885
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   556
   End
   Begin VB.ComboBox cboPalette 
      Height          =   315
      Left            =   15
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   1980
   End
   Begin vbalDropDownForm6.vbalDropDownClient ddcDropDOwn 
      Align           =   1  'Align Top
      Height          =   75
      Left            =   0
      ToolTipText     =   "Drag to make this menu float"
      Top             =   0
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   132
      Caption         =   "Palette"
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Color"
      Height          =   195
      Index           =   1
      Left            =   30
      TabIndex        =   3
      Top             =   690
      Width           =   855
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Palette Color"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   105
      Width           =   1395
   End
End
Attribute VB_Name = "frmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_CodeColor As Long
Private arr_files() As String

Private Sub cboPalette_Click()
    ClrPicker1.PathPaleta = arr_files(cboPalette.ListIndex + 1)
End Sub


Private Sub Form_Load()

    Dim path As String
    
    ReDim arr_files(0)
    ReDim arr_files(9)
    
    path = Util.StripPath(App.path) & "\pal\"
    
    arr_files(1) = path & "2.pal"
    arr_files(2) = path & "8.pal"
    arr_files(3) = path & "16.pal"
    arr_files(4) = path & "256c.pal"
    arr_files(5) = path & "256g.pal"
    arr_files(6) = path & "browser.pal"
    arr_files(7) = path & "default.pal"
    arr_files(8) = path & "named.pal"
    arr_files(9) = path & "windows.pal"
    
    cboPalette.AddItem "2 Colors"
    cboPalette.AddItem "8 Colors"
    cboPalette.AddItem "16 Colors"
    cboPalette.AddItem "256 Colors"
    cboPalette.AddItem "256 Grays"
    cboPalette.AddItem "Browser Colors"
    cboPalette.AddItem "Default Colors"
    cboPalette.AddItem "Named Colors"
    cboPalette.AddItem "Windows Colors"
    cboPalette.ListIndex = 0
    ClrPicker1.PathPaleta = arr_files(1)
    
    m_CodeColor = -1
    
End Sub


Public Property Get ShowState() As EWindowShowState
   ' This is to allow the parent form
   ' to control the current drop-down state:
   ShowState = ddcDropDOwn.ShowState
End Property
Public Property Let ShowState(ByVal eState As EWindowShowState)
   ' This is to allow the parent form
   ' to control the current drop-down state:
   ddcDropDOwn.ShowState = eState
End Property


Public Property Get CodeColor() As Long
    CodeColor = ClrPicker1.code
End Property

Public Property Let CodeColor(ByVal pCodeColor As Long)
    m_CodeColor = pCodeColor
End Property
