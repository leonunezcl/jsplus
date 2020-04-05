VERSION 5.00
Begin VB.Form frmNewPlug 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Add-In"
   ClientHeight    =   7335
   ClientLeft      =   4020
   ClientTop       =   2490
   ClientWidth     =   6855
   Icon            =   "frmNewPlug.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   31
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   30
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Index           =   2
      Left            =   60
      TabIndex        =   15
      Top             =   3360
      Width           =   6780
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "your class.vbp"
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
         Index           =   19
         Left            =   360
         TabIndex        =   29
         Top             =   1950
         Width           =   1245
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Visual Basic proyect with related plugin files."
         Height          =   195
         Index           =   18
         Left            =   1650
         TabIndex        =   28
         Top             =   1950
         Width           =   3105
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "To unregister an add-in use regsvr32 /u [add-in name.dll]"
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
         Index           =   17
         Left            =   135
         TabIndex        =   27
         Top             =   3030
         Width           =   4875
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "To register an add-in use regsvr32 [add-in name.dll]"
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
         Index           =   16
         Left            =   120
         TabIndex        =   26
         Top             =   2790
         Width           =   4410
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   6720
         Y1              =   2235
         Y2              =   2235
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "The project name must be ""plugin"". If this don't has this name the plugin maybe don't works."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   15
         Left            =   120
         TabIndex        =   25
         Top             =   2325
         Width           =   6585
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   $"frmNewPlug.frx":000C
         Height          =   390
         Index           =   13
         Left            =   45
         TabIndex        =   23
         Top             =   750
         Width           =   6780
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Form example that contains a hello world example"
         Height          =   195
         Index           =   12
         Left            =   1650
         TabIndex        =   22
         Top             =   1740
         Width           =   3495
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Aditional add-in information"
         Height          =   195
         Index           =   11
         Left            =   1650
         TabIndex        =   21
         Top             =   1515
         Width           =   1890
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Add-in class that contains the Start method"
         Height          =   195
         Index           =   10
         Left            =   1650
         TabIndex        =   20
         Top             =   1290
         Width           =   3030
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "frmPlugin.frm"
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
         Index           =   9
         Left            =   360
         TabIndex        =   19
         Top             =   1740
         Width           =   1110
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "mPlugin.bas"
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
         Index           =   8
         Left            =   360
         TabIndex        =   18
         Top             =   1515
         Width           =   1035
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "cPlugin.cls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   360
         TabIndex        =   17
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   $"frmNewPlug.frx":0096
         Height          =   390
         Index           =   6
         Left            =   45
         TabIndex        =   16
         Top             =   240
         Width           =   6735
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Class Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Index           =   1
      Left            =   60
      TabIndex        =   13
      Top             =   2100
      Width           =   6780
      Begin VB.TextBox txtClassId 
         Height          =   285
         Left            =   1365
         TabIndex        =   6
         Text            =   "cHelloWorld"
         Top             =   255
         Width           =   5175
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   $"frmNewPlug.frx":0149
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   14
         Left            =   1365
         TabIndex        =   24
         Top             =   555
         Width           =   5385
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   14
         Top             =   285
         Width           =   465
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Add-In Properties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2025
      Index           =   0
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   6780
      Begin VB.OptionButton opt 
         Caption         =   "No"
         Height          =   240
         Index           =   1
         Left            =   2040
         TabIndex        =   4
         Top             =   1290
         Width           =   675
      End
      Begin VB.OptionButton opt 
         Caption         =   "Yes"
         Height          =   240
         Index           =   0
         Left            =   1350
         TabIndex        =   3
         Top             =   1290
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.TextBox txtVersion 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "1.0"
         Top             =   1605
         Width           =   675
      End
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   1350
         TabIndex        =   2
         Text            =   "This add-in will display a hello world message"
         Top             =   915
         Width           =   5175
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   1350
         TabIndex        =   1
         Text            =   "Hello World"
         Top             =   570
         Width           =   5175
      End
      Begin VB.TextBox txtAutor 
         Height          =   285
         Left            =   1365
         TabIndex        =   0
         Text            =   "Javascript Developer"
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Return String:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   12
         Top             =   1305
         Width           =   975
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   11
         Top             =   1650
         Width           =   570
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Caption:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   10
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   9
         Top             =   945
         Width           =   840
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Autor:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   8
         Top             =   270
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmNewPlug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function crea_plugin() As Boolean

    On Error GoTo Errorcrea_plugin
    
    Dim srcCls As New cStringBuilder
    Dim srcBas As New cStringBuilder
    Dim srcFrm As New cStringBuilder
    Dim srcVbp As New cStringBuilder
    
    Dim Path As String
    Dim nFreeFile As Long
    
    If txtAutor.Text = "" Then
        txtAutor.SetFocus
        Exit Function
    End If
    
    If txtCaption.Text = "" Then
        txtCaption.SetFocus
        Exit Function
    End If
    
    If txtDescription.Text = "" Then
        txtDescription.SetFocus
        Exit Function
    End If
    
    If txtClassId.Text = "" Then
        txtClassId.SetFocus
        Exit Function
    End If
    
    Path = util.BrowseFolder(hwnd)
    If Path = "" Then Exit Function
    Path = util.StripPath(Path)
    
    'generar el modulo de clase cPlugin.cls
    
    srcCls.Append "VERSION 1.0 CLASS" & vbNewLine
    srcCls.Append "BEGIN" & vbNewLine
    srcCls.Append "MultiUse = -1  'True" & vbNewLine
    srcCls.Append "Persistable = 0  'NotPersistable" & vbNewLine
    srcCls.Append "DataBindingBehavior = 0  'vbNone" & vbNewLine
    srcCls.Append "DataSourceBehavior = 0   'vbNone" & vbNewLine
    srcCls.Append "MTSTransactionMode = 0   'NotAnMTSObject" & vbNewLine
    srcCls.Append "End" & vbNewLine
    srcCls.Append "Attribute VB_Name = " & Chr$(34) & txtClassId.Text & Chr$(34) & vbNewLine
    srcCls.Append "Attribute VB_GlobalNameSpace = True" & vbNewLine
    srcCls.Append "Attribute VB_Creatable = True" & vbNewLine
    srcCls.Append "Attribute VB_PredeclaredId = False" & vbNewLine
    srcCls.Append "Attribute VB_Exposed = True" & vbNewLine

    srcCls.Append "Option Explicit" & vbNewLine
    srcCls.Append "" & vbNewLine
    srcCls.Append "Public OutputString As New cStringBuilder" & vbNewLine
    srcCls.Append "Private m_ReturnString As Boolean" & vbNewLine
    srcCls.Append "Private m_Autor As String" & vbNewLine
    srcCls.Append "Private m_Caption As String" & vbNewLine
    srcCls.Append "Private m_Description As String" & vbNewLine
    srcCls.Append "Private m_Version As String" & vbNewLine
    srcCls.Append "Private m_ClassId As String" & vbNewLine
    srcCls.Append "Private m_ErrNumber As String" & vbNewLine
    srcCls.Append "Private m_ErrMessage As String" & vbNewLine
    
    srcCls.Append "Public Function Start() As Boolean" & vbNewLine
    srcCls.Append "" & vbNewLine
    srcCls.Append "     On Error GoTo StartError" & vbNewLine
    srcCls.Append "" & vbNewLine
    srcCls.Append "     Dim frm As New frmPlugin" & vbNewLine
    srcCls.Append "" & vbNewLine
    srcCls.Append "     frm.Show vbModal" & vbNewLine
    srcCls.Append "" & vbNewLine
    srcCls.Append "     If Len(glbOutputString.ToString) > 0 Then" & vbNewLine
    srcCls.Append "         OutputString.Append glbOutputString.ToString" & vbNewLine
    srcCls.Append "         m_ReturnString = True" & vbNewLine
    srcCls.Append "     End If" & vbNewLine
    srcCls.Append "" & vbNewLine
    srcCls.Append "     Set frm = Nothing" & vbNewLine
    srcCls.Append "" & vbNewLine
    srcCls.Append "     Start = True" & vbNewLine
    srcCls.Append "" & vbNewLine
    srcCls.Append "     Exit Function" & vbNewLine
    srcCls.Append "StartError:" & vbNewLine
    srcCls.Append "m_ErrNumber = Err.Number" & vbNewLine
    srcCls.Append "m_ErrMessage = Err.Description" & vbNewLine
    srcCls.Append "End Function" & vbNewLine
    srcCls.Append "" & vbNewLine
    srcCls.Append "Private Sub Class_Initialize()" & vbNewLine
    srcCls.Append "    m_Autor = " & Chr$(34) & txtAutor.Text & Chr$(34) & vbNewLine
    srcCls.Append "    m_Caption = " & Chr$(34) & txtCaption.Text & Chr$(34) & vbNewLine
    srcCls.Append "    m_Description = " & Chr$(34) & txtDescription.Text & Chr$(34) & vbNewLine
    srcCls.Append "    m_Version = " & "App.Major & " & Chr$(34) & "." & Chr$(34) & " & App.Minor & " & Chr$(34) & "." & Chr$(34) & " & App.Revision" & vbNewLine
    srcCls.Append "    m_ClassId = " & Chr$(34) & txtClassId.Text & Chr$(34) & vbNewLine
    srcCls.Append "End Sub" & vbNewLine

    srcCls.Append "Public Property Get ClassId() As String" & vbNewLine
    srcCls.Append "     ClassId = m_ClassId" & vbNewLine
    srcCls.Append "End Property" & vbNewLine

    srcCls.Append "Public Property Get ReturnString() As Boolean" & vbNewLine
    srcCls.Append "    ReturnString = m_ReturnString" & vbNewLine
    srcCls.Append "End Property" & vbNewLine

    srcCls.Append "Public Property Get Autor() As String" & vbNewLine
    srcCls.Append "    Autor = m_Autor" & vbNewLine
    srcCls.Append "End Property" & vbNewLine

    srcCls.Append "Public Property Get Caption() As String" & vbNewLine
    srcCls.Append "    Caption = m_Caption" & vbNewLine
    srcCls.Append "End Property" & vbNewLine

    srcCls.Append "Public Property Get Description() As String" & vbNewLine
    srcCls.Append "    Description = m_Description" & vbNewLine
    srcCls.Append "End Property" & vbNewLine

    srcCls.Append "Public Property Get Version() As String" & vbNewLine
    srcCls.Append "    Version = m_Version" & vbNewLine
    srcCls.Append "End Property" & vbNewLine

    srcCls.Append "Public Property Get ErrNumber() As Long" & vbNewLine
    srcCls.Append "    ErrNumber = m_ErrNumber" & vbNewLine
    srcCls.Append "End Property" & vbNewLine

    srcCls.Append "Public Property Let ErrNumber(ByVal pErrNumber As Long)" & vbNewLine
    srcCls.Append "    m_ErrNumber = pErrNumber" & vbNewLine
    srcCls.Append "End Property" & vbNewLine

    srcCls.Append "Public Property Get ErrMessage() As String" & vbNewLine
    srcCls.Append "    ErrMessage = m_ErrMessage" & vbNewLine
    srcCls.Append "End Property" & vbNewLine

    srcCls.Append "Public Property Let ErrMessage(ByVal pErrMessage As String)" & vbNewLine
    srcCls.Append "    m_ErrMessage = pErrMessage" & vbNewLine
    srcCls.Append "End Property" & vbNewLine

    nFreeFile = FreeFile
    Open Path & "cPlugin.cls" For Output As #nFreeFile
        Print #nFreeFile, srcCls.ToString
    Close #nFreeFile
    
    'generar el modulo .bas mPlugin.bas
    nFreeFile = FreeFile
    srcBas.Append "Attribute VB_Name = " & Chr$(34) & "mPlugin" & Chr$(34) & vbNewLine
    srcBas.Append "Option Explicit" & vbNewLine
    srcBas.Append ""
    srcBas.Append "Public glbOutputString As New cStringBuilder" & vbNewLine

    Open Path & "mPlugin.bas" For Output As #nFreeFile
        Print #nFreeFile, srcBas.ToString
    Close #nFreeFile
    
    'generar el formulario
    srcFrm.Append "Version 5.00" & vbNewLine
    srcFrm.Append "Begin VB.Form frmPlugin" & vbNewLine
    srcFrm.Append "     Caption = " & Chr$(34) & "Plugin Example" & Chr$(34) & vbNewLine
    srcFrm.Append "     ClientHeight = 1425" & vbNewLine
    srcFrm.Append "     ClientLeft = 4395" & vbNewLine
    srcFrm.Append "     ClientTop = 2145" & vbNewLine
    srcFrm.Append "     ClientWidth = 2475" & vbNewLine
    srcFrm.Append "     LinkTopic = " & Chr$(34) & "Form1" & Chr$(34) & vbNewLine
    srcFrm.Append "     ScaleHeight = 1425" & vbNewLine
    srcFrm.Append "     ScaleWidth = 2475" & vbNewLine
    srcFrm.Append "     Begin VB.CommandButton cmdOk" & vbNewLine
    srcFrm.Append "         Caption = " & Chr(34) & "Hello World" & Chr$(34) & vbNewLine
    srcFrm.Append "         Height = 525" & vbNewLine
    srcFrm.Append "         Left = 555" & vbNewLine
    srcFrm.Append "         TabIndex = 0" & vbNewLine
    srcFrm.Append "         Top = 375" & vbNewLine
    srcFrm.Append "         Width = 1350" & vbNewLine
    srcFrm.Append "     End" & vbNewLine
    srcFrm.Append "End" & vbNewLine
    srcFrm.Append "" & vbNewLine
    srcFrm.Append "Attribute VB_Name = " & Chr$(34) & "frmPlugin" & Chr$(34) & vbNewLine
    srcFrm.Append "Attribute VB_GlobalNameSpace = False" & vbNewLine
    srcFrm.Append "Attribute VB_Creatable = False" & vbNewLine
    srcFrm.Append "Attribute VB_PredeclaredId = True" & vbNewLine
    srcFrm.Append "Attribute VB_Exposed = False" & vbNewLine
    srcFrm.Append "Option Explicit" & vbNewLine
    srcFrm.Append "Private Util As New cLibrary" & vbNewLine
    srcFrm.Append "Private Sub cmdOk_Click()" & vbNewLine
    srcFrm.Append "     glbOutputString.Append " & Chr$(34) & "Hello World!" & Chr$(34) & vbNewLine
    srcFrm.Append "     Unload Me" & vbNewLine
    srcFrm.Append "End Sub" & vbNewLine

    srcFrm.Append "Private Sub Form_Load()" & vbNewLine
    srcFrm.Append "     Util.CenterForm Me" & vbNewLine
    srcFrm.Append "     'DrawXPCtl Me" & vbNewLine
    srcFrm.Append "End Sub" & vbNewLine

    srcFrm.Append "Private Sub Form_Unload(Cancel As Integer)" & vbNewLine
    srcFrm.Append "     Set frmPlugin = Nothing" & vbNewLine
    srcFrm.Append "End Sub" & vbNewLine

    nFreeFile = FreeFile
    Open Path & "frmPlugin.frm" For Output As #nFreeFile
        Print #nFreeFile, srcFrm.ToString
    Close #nFreeFile
    
    'generar el vbp
    srcVbp.Append "Type=OleDll" & vbNewLine
    srcVbp.Append "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\..\..\WINNT\system32\stdole2.tlb#OLE Automation" & vbNewLine
    srcVbp.Append "Reference=*\G{58FE8683-AE0A-11D7-ACE9-0001030706C1}#3.0#0#C:\Archivos de programa\jsplus\dll\xpstylelib.dll# Windows XP Emulation Library ( Support MessageBox,InputBox and About Dialog Box )" & vbNewLine
    srcVbp.Append "Reference=*\G{B0F6465F-C038-418D-9954-C80B985D847D}#1.0#0#C:\Archivos de programa\jsplus\dll\vbsengine.dll#Javascript Plus! Engine - Powered by vbaccelerator components" & vbNewLine
    srcVbp.Append "Reference=*\G{24042BC3-630E-4340-8906-B3E0DB1FC0BF}#1.0#0#C:\Archivos de programa\jsplus\dll\vbslibrary.dll#VBSoftware - Library" & vbNewLine
    srcVbp.Append "Class=" & txtClassId.Text & "; cPlugin.cls" & vbNewLine
    srcVbp.Append "Module=mPlugin; mPlugin.bas" & vbNewLine
    srcVbp.Append "Form=frmPlugin.frm" & vbNewLine
    srcVbp.Append "Startup=" & Chr$(34) & "(None)" & Chr$(34) & vbNewLine
    srcVbp.Append "HelpFile=" & Chr$(34) & Chr$(34) & vbNewLine
    srcVbp.Append "Title=" & Chr$(34) & "plugin" & Chr$(34) & vbNewLine
    srcVbp.Append "Command32=" & Chr$(34) & Chr$(34) & vbNewLine
    srcVbp.Append "Name=plugin" & vbNewLine
    srcVbp.Append "HelpContextID=" & Chr$(34) & "0" & Chr$(34) & vbNewLine
    srcVbp.Append "Description=" & Chr$(34) & "JavaScript Plus Plugin - " & txtDescription.Text & Chr$(34) & vbNewLine
    srcVbp.Append "CompatibleMode=" & Chr$(34) & "0" & Chr$(34) & vbNewLine
    srcVbp.Append "MajorVer=1" & vbNewLine
    srcVbp.Append "MinorVer=0" & vbNewLine
    srcVbp.Append "RevisionVer=0" & vbNewLine
    srcVbp.Append "AutoIncrementVer=0" & vbNewLine
    srcVbp.Append "ServerSupportFiles=0" & vbNewLine
    srcVbp.Append "VersionCompanyName=" & Chr$(34) & Chr$(34) & vbNewLine
    srcVbp.Append "CompilationType=-1" & vbNewLine
    srcVbp.Append "OptimizationType=0" & vbNewLine
    srcVbp.Append "FavorPentiumPro(tm)=0" & vbNewLine
    srcVbp.Append "CodeViewDebugInfo=0" & vbNewLine
    srcVbp.Append "NoAliasing=0" & vbNewLine
    srcVbp.Append "BoundsCheck=0" & vbNewLine
    srcVbp.Append "OverflowCheck=0" & vbNewLine
    srcVbp.Append "FlPointCheck=0" & vbNewLine
    srcVbp.Append "FDIVCheck=0" & vbNewLine
    srcVbp.Append "UnroundedFP=0" & vbNewLine
    srcVbp.Append "StartMode=1" & vbNewLine
    srcVbp.Append "Unattended=0" & vbNewLine
    srcVbp.Append "Retained=0" & vbNewLine
    srcVbp.Append "ThreadPerObject=0" & vbNewLine
    srcVbp.Append "MaxNumberOfThreads=1" & vbNewLine
    srcVbp.Append "ThreadingModel=1" & vbNewLine
    
    nFreeFile = FreeFile
    Open Path & Mid$(txtClassId.Text, 2) & ".vbp" For Output As #nFreeFile
        Print #nFreeFile, srcVbp.ToString
    Close #nFreeFile
    
    Set srcCls = Nothing
    Set srcBas = Nothing
    Set srcFrm = Nothing
    
    'generar el .bas
    
    crea_plugin = True
    
    Exit Function
    
Errorcrea_plugin:
    MsgBox "crea_plugin : " & Err & " " & Error$, vbCritical
    
End Function
Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        If crea_plugin() Then
            Unload Me
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.CenterForm Me
    
    Debug.Print "load"
        
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmNewPlug = Nothing
End Sub


