VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmActiveX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ActiveX Object Browser"
   ClientHeight    =   6465
   ClientLeft      =   2895
   ClientTop       =   1695
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmActiveX.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwCom 
      Height          =   5655
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9975
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsertar 
      Caption         =   "&Insert"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select ActiveX Component ..."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   2115
   End
End
Attribute VB_Name = "frmActiveX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Populate()

    Dim iSectCount As Long, iSect As Long, sSections() As String
    Dim iVerCount As Long, iVer As Long, sVersions() As String
    Dim iExeSectCount As Long, sExeSect() As String
    Dim iExeSect As Long
    Dim bFoundExeSect As Boolean
    Dim sExists As String
    Dim cTLI As cTypeLibInfo
    
    Dim C As Integer

    Dim cR As New cRegistry
    cR.ClassKey = HKEY_CLASSES_ROOT
    cR.ValueType = REG_SZ
    cR.SectionKey = "TypeLib"
   
    Dim Archivo As String
    Dim linea As String
    Dim nFreeFile As Long
    nFreeFile = FreeFile
        
    Archivo = util.StripPath(App.Path) & "config\activex.ini"
    If ArchivoExiste2(Archivo) Then
        C = 1
        Open Archivo For Input As #nFreeFile
            Do While Not EOF(nFreeFile)
                Line Input #nFreeFile, linea
                
                lvwCom.ListItems.Add , "k" & C, CStr(C)
                
                lvwCom.ListItems(C).SubItems(1) = util.Explode(linea, 1, "|")
                lvwCom.ListItems(C).SubItems(2) = util.Explode(linea, 2, "|")
                lvwCom.ListItems(C).SubItems(3) = util.Explode(linea, 3, "|")
                lvwCom.ListItems(C).SubItems(4) = util.Explode(linea, 4, "|")
                lvwCom.ListItems(C).SubItems(5) = util.Explode(linea, 5, "|")
                
                C = C + 1
            Loop
        Close #nFreeFile
    Else
        C = 1
        Open Archivo For Output As #nFreeFile
        ' Get the registered Type Libs:
        If cR.EnumerateSections(sSections(), iSectCount) Then
           For iSect = 1 To iSectCount
              ' Enumerate the versions for each typelib:
              cR.SectionKey = "TypeLib\" & sSections(iSect)
              If cR.EnumerateSections(sVersions(), iVerCount) Then
                 For iVer = 1 To iVerCount
                    Set cTLI = New cTypeLibInfo
                    cTLI.CLSID = sSections(iSect)
                    cTLI.Ver = sVersions(iVer)
                    cR.SectionKey = "TypeLib\" & sSections(iSect) & "\" & sVersions(iVer)
                    cTLI.Name = cR.Value
                    cR.EnumerateSections sExeSect(), iExeSectCount
                    If iExeSectCount > 0 Then
                       bFoundExeSect = False
                       For iExeSect = 1 To iExeSectCount
                          If IsNumeric(sExeSect(iExeSect)) Then
                             cR.SectionKey = cR.SectionKey & "\" & sExeSect(iExeSect) & "\win32"
                             bFoundExeSect = True
                             Exit For
                          End If
                       Next iExeSect
                       If bFoundExeSect Then
                          cTLI.Path = cR.Value
                          If ArchivoExiste2(cTLI.Path) Then
                             sExists = "Y"
                          Else
                             sExists = "N"
                          End If
                       Else
                          sExists = "N"
                       End If
                    Else
                       sExists = "N"
                    End If
                    cTLI.Exists = (StrComp(sExists, "Y") = 0)
                    lvwCom.ListItems.Add , "k" & C, CStr(C)
                    If Len(cTLI.Name) > 0 Then
                       lvwCom.ListItems(C).SubItems(1) = cTLI.Name
                       'lstTypeLibs.AddItem cTLI.Name & vbTab & sExists
                    Else
                       lvwCom.ListItems(C).SubItems(1) = cTLI.CLSID
                       'lstTypeLibs.AddItem cTLI.CLSID & vbTab & sExists
                    End If
                    lvwCom.ListItems(C).SubItems(2) = cTLI.CLSID
                    lvwCom.ListItems(C).SubItems(3) = cTLI.Ver
                    lvwCom.ListItems(C).SubItems(4) = cTLI.Path
                    lvwCom.ListItems(C).SubItems(5) = sExists
                    
                    Print #nFreeFile, lvwCom.ListItems(C).SubItems(1) & "|" & _
                                      lvwCom.ListItems(C).SubItems(2) & "|" & _
                                      lvwCom.ListItems(C).SubItems(3) & "|" & _
                                      lvwCom.ListItems(C).SubItems(4) & "|" & _
                                      lvwCom.ListItems(C).SubItems(5)
                    
                    C = C + 1
                 Next iVer
              End If
           Next iSect
        End If
        Close #nFreeFile
    End If
    
   'lstTypeLibs.Visible = True
   
End Sub



Private Sub cmdInsertar_Click()

    Dim texto As String
    Dim Itmx As ListItem
    
    If Not lvwCom.SelectedItem Is Nothing Then
        Set Itmx = lvwCom.SelectedItem
        texto = "<OBJECT ID=" & Chr$(34) & Itmx.SubItems(1) & Chr$(34)
        texto = texto & " WIDTH=" & Chr$(34) & Chr$(34) & " "
        texto = texto & " HEIGHT=" & Chr$(34) & Chr$(34) & " "
        texto = texto & " CLASSID=" & Chr$(34) & Itmx.SubItems(2) & Chr$(34) & ">"
        Call frmMain.ActiveForm.Insertar(texto)
    End If
    
    Set Itmx = Nothing
    
    Unload Me
    
End Sub


Private Sub cmdOk_Click()
   Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    util.Hourglass hwnd, True
        
    util.CenterForm Me
        
    Load frmWait
    frmWait.Caption = "Loading ActiveX Information ..."
    frmWait.lbl(0).Caption = "Please wait while JavaScript Plus! load Activex Information"
    
    With lvwCom
        .ColumnHeaders.Add , "k1", "Nº", 500
        .ColumnHeaders.Add , "k2", "Name", 4000
        .ColumnHeaders.Add , "k3", "GUID", 4000
        .ColumnHeaders.Add , "k4", "Version", 1000
        .ColumnHeaders.Add , "k5", "Path", 4000
        .ColumnHeaders.Add , "k6", "Exists", 1000
    End With
    
    frmWait.Show
    DoEvents
    Refresh
    
    Populate
        
    flat_lviews Me
    
    util.Hourglass hwnd, False
    
    Unload frmWait
    
    Debug.Print "load"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "unload"
    Set frmActiveX = Nothing
End Sub




Private Sub lvwCom_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'If ColumnHeader.
    '    ColumnHeader.SortType = eLVSortNumeric
    '    ColumnHeader.SortOrder = eSortOrderDescending
    'Else
    '    ColumnHeader.SortOrder = eSortOrderDescending
    '    ColumnHeader.SortType = eLVSortNumeric
    'End If
    
End Sub

Private Sub lvwCom_DblClick()
    cmdInsertar_Click
End Sub

