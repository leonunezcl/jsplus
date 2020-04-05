VERSION 5.00
Begin VB.Form frmTables 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Table Wizard ..."
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   3930
      Width           =   1290
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   375
      Left            =   2025
      TabIndex        =   31
      Top             =   3930
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3210
      Begin VB.TextBox txtRows 
         Height          =   285
         Left            =   720
         TabIndex        =   28
         Text            =   "1"
         Top             =   3105
         Width           =   660
      End
      Begin VB.TextBox txtColls 
         Height          =   285
         Left            =   2040
         TabIndex        =   27
         Text            =   "1"
         Top             =   3105
         Width           =   645
      End
      Begin VB.PictureBox pic1 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   360
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   26
         Top             =   360
         Width           =   400
      End
      Begin VB.PictureBox pic5 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   2280
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   25
         Top             =   360
         Width           =   400
      End
      Begin VB.PictureBox pic4 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1800
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   24
         Top             =   360
         Width           =   400
      End
      Begin VB.PictureBox pic3 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1320
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   23
         Top             =   360
         Width           =   400
      End
      Begin VB.PictureBox pic2 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   840
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   22
         Top             =   360
         Width           =   400
      End
      Begin VB.PictureBox pic6 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   360
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   21
         Top             =   840
         Width           =   400
      End
      Begin VB.PictureBox pic10 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   2280
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   20
         Top             =   840
         Width           =   400
      End
      Begin VB.PictureBox pic9 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1800
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   19
         Top             =   840
         Width           =   400
      End
      Begin VB.PictureBox pic8 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1320
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   18
         Top             =   840
         Width           =   400
      End
      Begin VB.PictureBox pic7 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   840
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   17
         Top             =   840
         Width           =   400
      End
      Begin VB.PictureBox pic11 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   360
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   16
         Top             =   1320
         Width           =   400
      End
      Begin VB.PictureBox pic15 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   2280
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   15
         Top             =   1320
         Width           =   400
      End
      Begin VB.PictureBox pic14 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1800
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   14
         Top             =   1320
         Width           =   400
      End
      Begin VB.PictureBox pic13 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1320
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   13
         Top             =   1320
         Width           =   400
      End
      Begin VB.PictureBox pic12 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   840
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   12
         Top             =   1320
         Width           =   400
      End
      Begin VB.PictureBox pic16 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   360
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   11
         Top             =   1800
         Width           =   400
      End
      Begin VB.PictureBox pic20 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   2280
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   10
         Top             =   1800
         Width           =   400
      End
      Begin VB.PictureBox pic19 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1800
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   9
         Top             =   1800
         Width           =   400
      End
      Begin VB.PictureBox pic18 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1320
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   8
         Top             =   1800
         Width           =   400
      End
      Begin VB.PictureBox pic17 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   840
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   7
         Top             =   1800
         Width           =   400
      End
      Begin VB.PictureBox pic21 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   360
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   6
         Top             =   2280
         Width           =   400
      End
      Begin VB.PictureBox pic25 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   2280
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   5
         Top             =   2280
         Width           =   400
      End
      Begin VB.PictureBox pic24 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1800
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   4
         Top             =   2280
         Width           =   400
      End
      Begin VB.PictureBox pic23 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   1320
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   3
         Top             =   2280
         Width           =   400
      End
      Begin VB.PictureBox pic22 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   840
         ScaleHeight     =   405
         ScaleWidth      =   405
         TabIndex        =   2
         Top             =   2280
         Width           =   400
      End
      Begin VB.Label Label1 
         Caption         =   "Rows:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   3120
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Colls:"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   29
         Top             =   3120
         Width           =   780
      End
   End
   Begin VB.Label lblTable 
      Height          =   255
      Left            =   450
      TabIndex        =   0
      Top             =   6900
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
' Project:  Casper HTML   v.2.0                              *
' Filename: n/a                                              *
' Author:   Vladimir S. Pekulas Jr.                          *
' Date:     7/22/2000                                        *
' Copyright © 2000 Vladimir S. Pekulas Jr.                   *
'                                                            *
' Use this program as you wish, but please let me know       *
' if you like it. Anyway, you can do whatever you want       *
' with it. I'm not responsible for any demage tough :)       *
'*************************************************************



Function AddTable(ColumnCount As Long, RowCount As Long) As String
On Error Resume Next
'Add table
Dim tmp
Dim j As Long
Dim k As Long
Dim quote$
quote$ = Chr$(34)
tmp = "<TABLE Align=" & Chr(34) & "center" & Chr(34) & " Border=" & Chr(34) & "0" & Chr(34) & ">" & vbCrLf
For j = 1 To RowCount
    tmp = tmp & "<TR>" & vbCrLf & "<TD> </TD>" & vbCrLf
    If ColumnCount > 1 Then
        For k = 2 To ColumnCount
            tmp = tmp & "<TD> </TD>" & vbCrLf
        Next k
    End If
    tmp = tmp & vbTab & "</TR>"
    
Next j

tmp = tmp & "</TABLE>" & vbCrLf
fMainForm.ActiveForm.rtfText.SelText = tmp
End Function



Private Sub cmdCreate_Click()
On Error Resume Next
 Call AddTable(txtColls.Text, txtRows.Text)
 Unload Me
End Sub


Function CleanUp()
 pic1.BackColor = &H80000009
 pic2.BackColor = &H80000009
 pic3.BackColor = &H80000009
 pic4.BackColor = &H80000009
 pic5.BackColor = &H80000009
 pic6.BackColor = &H80000009
 pic7.BackColor = &H80000009
 pic8.BackColor = &H80000009
 pic9.BackColor = &H80000009
 pic10.BackColor = &H80000009
 pic11.BackColor = &H80000009
 pic12.BackColor = &H80000009
 pic13.BackColor = &H80000009
 pic14.BackColor = &H80000009
 pic15.BackColor = &H80000009
 pic16.BackColor = &H80000009
 pic17.BackColor = &H80000009
 pic18.BackColor = &H80000009
 pic19.BackColor = &H80000009
 pic20.BackColor = &H80000009
 pic21.BackColor = &H80000009
 pic22.BackColor = &H80000009
 pic23.BackColor = &H80000009
 pic24.BackColor = &H80000009
 pic25.BackColor = &H80000009
End Function


Private Sub Command2_Click()
 Unload Me
End Sub


Private Sub Form_Load()
pic1_Click
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormDrag(Me)
End Sub


Private Sub pic1_Click()
Call CleanUp
lblTable.Caption = "1"
pic1.BackColor = &H8000000D
txtRows.Text = "1"
txtColls.Text = "1"
End Sub


Private Sub pic11_Click()
Call CleanUp
lblTable.Caption = "11"
pic1.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic11.BackColor = &H8000000D
txtRows.Text = "3"
txtColls.Text = "1"
End Sub

Private Sub pic12_Click()
Call CleanUp
lblTable.Caption = "12"
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
pic11.BackColor = &H8000000D
pic12.BackColor = &H8000000D
txtRows.Text = "3"
txtColls.Text = "2"
End Sub

Private Sub pic13_Click()
lblTable.Caption = "13"
Call CleanUp
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic3.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
pic8.BackColor = &H8000000D
pic11.BackColor = &H8000000D
pic12.BackColor = &H8000000D
pic13.BackColor = &H8000000D
txtRows.Text = "3"
txtColls.Text = "3"
End Sub

Private Sub pic14_Click()
Call CleanUp
lblTable.Caption = "14"
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic3.BackColor = &H8000000D
pic4.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
pic8.BackColor = &H8000000D
pic9.BackColor = &H8000000D
pic11.BackColor = &H8000000D
pic12.BackColor = &H8000000D
pic13.BackColor = &H8000000D
pic14.BackColor = &H8000000D
txtRows.Text = "3"
txtColls.Text = "4"
End Sub

Private Sub pic15_Click()
lblTable.Caption = "15"
Call CleanUp
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic3.BackColor = &H8000000D
pic4.BackColor = &H8000000D
pic5.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
pic8.BackColor = &H8000000D
pic9.BackColor = &H8000000D
pic10.BackColor = &H8000000D
pic11.BackColor = &H8000000D
pic12.BackColor = &H8000000D
pic13.BackColor = &H8000000D
pic14.BackColor = &H8000000D
pic15.BackColor = &H8000000D
txtRows.Text = "3"
txtColls.Text = "5"
End Sub

Private Sub pic16_Click()
lblTable.Caption = "16"
Call CleanUp
pic1.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic11.BackColor = &H8000000D
pic16.BackColor = &H8000000D
txtRows.Text = "4"
txtColls.Text = "1"
End Sub

Private Sub pic17_Click()
lblTable.Caption = "17"
Call CleanUp
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
pic11.BackColor = &H8000000D
pic12.BackColor = &H8000000D
pic16.BackColor = &H8000000D
pic17.BackColor = &H8000000D
txtRows.Text = "4"
txtColls.Text = "2"
End Sub

Private Sub pic18_Click()
lblTable.Caption = "18"
Call CleanUp
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic3.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
pic8.BackColor = &H8000000D
pic11.BackColor = &H8000000D
pic12.BackColor = &H8000000D
pic13.BackColor = &H8000000D
pic16.BackColor = &H8000000D
pic17.BackColor = &H8000000D
pic18.BackColor = &H8000000D
txtRows.Text = "4"
txtColls.Text = "3"
End Sub

Private Sub pic19_Click()
lblTable.Caption = "19"
Call CleanUp
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic3.BackColor = &H8000000D
pic4.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
pic8.BackColor = &H8000000D
pic9.BackColor = &H8000000D
pic11.BackColor = &H8000000D
pic12.BackColor = &H8000000D
pic13.BackColor = &H8000000D
pic14.BackColor = &H8000000D
pic16.BackColor = &H8000000D
pic17.BackColor = &H8000000D
pic18.BackColor = &H8000000D
pic19.BackColor = &H8000000D
txtRows.Text = "4"
txtColls.Text = "4"
End Sub

Private Sub pic2_Click()
Call CleanUp
Call pic1_Click
lblTable.Caption = ""
lblTable.Caption = "2"
pic2.BackColor = &H8000000D
txtRows.Text = "1"
txtColls.Text = "2"
End Sub

Private Sub pic20_Click()
lblTable.Caption = "20"
Call CleanUp
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic3.BackColor = &H8000000D
pic4.BackColor = &H8000000D
pic5.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
pic8.BackColor = &H8000000D
pic9.BackColor = &H8000000D
pic10.BackColor = &H8000000D
pic11.BackColor = &H8000000D
pic12.BackColor = &H8000000D
pic13.BackColor = &H8000000D
pic14.BackColor = &H8000000D
pic15.BackColor = &H8000000D
pic16.BackColor = &H8000000D
pic17.BackColor = &H8000000D
pic18.BackColor = &H8000000D
pic19.BackColor = &H8000000D
pic20.BackColor = &H8000000D
txtRows.Text = "4"
txtColls.Text = "5"
End Sub

Private Sub pic21_Click()
Call CleanUp
lblTable.Caption = "21"
pic1.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic11.BackColor = &H8000000D
pic16.BackColor = &H8000000D
pic21.BackColor = &H8000000D
txtRows.Text = "5"
txtColls.Text = "1"
End Sub

Private Sub pic22_Click()
Call CleanUp
lblTable.Caption = "22"
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
pic11.BackColor = &H8000000D
pic12.BackColor = &H8000000D
pic16.BackColor = &H8000000D
pic17.BackColor = &H8000000D
pic21.BackColor = &H8000000D
pic22.BackColor = &H8000000D
txtRows.Text = "5"
txtColls.Text = "2"
End Sub

Private Sub pic23_Click()
lblTable.Caption = "23"
Call CleanUp
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic3.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
pic8.BackColor = &H8000000D
pic11.BackColor = &H8000000D
pic12.BackColor = &H8000000D
pic13.BackColor = &H8000000D
pic16.BackColor = &H8000000D
pic17.BackColor = &H8000000D
pic18.BackColor = &H8000000D
pic21.BackColor = &H8000000D
pic22.BackColor = &H8000000D
pic23.BackColor = &H8000000D
txtRows.Text = "5"
txtColls.Text = "3"
End Sub

Private Sub pic24_Click()
lblTable.Caption = "24"
Call CleanUp
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic3.BackColor = &H8000000D
pic4.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
pic8.BackColor = &H8000000D
pic9.BackColor = &H8000000D
pic11.BackColor = &H8000000D
pic12.BackColor = &H8000000D
pic13.BackColor = &H8000000D
pic14.BackColor = &H8000000D
pic16.BackColor = &H8000000D
pic17.BackColor = &H8000000D
pic18.BackColor = &H8000000D
pic19.BackColor = &H8000000D
pic21.BackColor = &H8000000D
pic22.BackColor = &H8000000D
pic23.BackColor = &H8000000D
pic24.BackColor = &H8000000D
txtRows.Text = "5"
txtColls.Text = "4"
End Sub

Private Sub pic25_Click()
lblTable.Caption = "25"
Call CleanUp
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic3.BackColor = &H8000000D
pic4.BackColor = &H8000000D
pic5.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
pic8.BackColor = &H8000000D
pic9.BackColor = &H8000000D
pic10.BackColor = &H8000000D
pic11.BackColor = &H8000000D
pic12.BackColor = &H8000000D
pic13.BackColor = &H8000000D
pic14.BackColor = &H8000000D
pic15.BackColor = &H8000000D
pic16.BackColor = &H8000000D
pic17.BackColor = &H8000000D
pic18.BackColor = &H8000000D
pic19.BackColor = &H8000000D
pic20.BackColor = &H8000000D
pic21.BackColor = &H8000000D
pic22.BackColor = &H8000000D
pic23.BackColor = &H8000000D
pic24.BackColor = &H8000000D
pic25.BackColor = &H8000000D
txtRows.Text = "5"
txtColls.Text = "5"
End Sub

Private Sub pic3_Click()
Call CleanUp
Call pic2_Click
lblTable.Caption = ""
lblTable.Caption = "3"
pic3.BackColor = &H8000000D
txtRows.Text = "1"
txtColls.Text = "3"
End Sub

Private Sub pic4_Click()
Call CleanUp
Call pic2_Click
lblTable.Caption = ""
lblTable.Caption = "4"
pic3.BackColor = &H8000000D
pic4.BackColor = &H8000000D
txtRows.Text = "1"
txtColls.Text = "4"
End Sub

Private Sub pic5_Click()
Call CleanUp
Call pic4_Click
lblTable.Caption = ""
lblTable.Caption = "5"
pic3.BackColor = &H8000000D
pic5.BackColor = &H8000000D
txtRows.Text = "1"
txtColls.Text = "5"
End Sub

Private Sub pic6_Click()
Call CleanUp
lblTable.Caption = "6"
pic1.BackColor = &H8000000D
pic6.BackColor = &H8000000D
txtRows.Text = "2"
txtColls.Text = "1"
End Sub

Private Sub pic7_Click()
Call CleanUp
lblTable.Caption = "7"
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
txtRows.Text = "2"
txtColls.Text = "2"
End Sub

Private Sub pic8_Click()
lblTable.Caption = "8"
Call CleanUp
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic3.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
pic8.BackColor = &H8000000D
txtRows.Text = "2"
txtColls.Text = "3"
End Sub

Private Sub pic9_Click()
lblTable.Caption = "9"
Call CleanUp
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic3.BackColor = &H8000000D
pic4.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
pic8.BackColor = &H8000000D
pic9.BackColor = &H8000000D
txtRows.Text = "2"
txtColls.Text = "4"
End Sub

Private Sub pic10_Click()
lblTable.Caption = "10"
Call CleanUp
pic1.BackColor = &H8000000D
pic2.BackColor = &H8000000D
pic3.BackColor = &H8000000D
pic4.BackColor = &H8000000D
pic5.BackColor = &H8000000D
pic6.BackColor = &H8000000D
pic7.BackColor = &H8000000D
pic8.BackColor = &H8000000D
pic9.BackColor = &H8000000D
pic10.BackColor = &H8000000D
txtRows.Text = "2"
txtColls.Text = "5"
End Sub

