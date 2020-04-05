VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmJsERR 
   Caption         =   "Runtime & Sintax Errors"
   ClientHeight    =   5325
   ClientLeft      =   2145
   ClientTop       =   2685
   ClientWidth     =   7245
   ControlBox      =   0   'False
   Icon            =   "frmJsERR.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   7245
   Begin SHDocVwCtl.WebBrowser web1 
      Height          =   2250
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4500
      ExtentX         =   7937
      ExtentY         =   3969
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmJsERR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Archivo As String
Public tipo As String
Private Sub Form_Load()
    
    If tipo = "R" Then
        Me.Caption = "Runtime Errors"
    Else
        Me.Caption = "Sintax Errors"
    End If
    
    web1.Navigate Archivo
End Sub


Private Sub Form_Resize()
    
    If WindowState <> vbMinimized Then
        web1.Move 0, 0, ScaleWidth, ScaleHeight
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmJsERR = Nothing
End Sub


