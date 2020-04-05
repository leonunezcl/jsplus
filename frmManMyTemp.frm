VERSION 5.00
Begin VB.Form frmManMyTemp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4170
   ClientLeft      =   3420
   ClientTop       =   1485
   ClientWidth     =   7035
   Icon            =   "frmManMyTemp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmManMyTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    Set frmManMyTemp = Nothing
End Sub


