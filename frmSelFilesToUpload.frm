VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelFilesToUpload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Upload Files to FTP Server"
   ClientHeight    =   5715
   ClientLeft      =   2685
   ClientTop       =   3015
   ClientWidth     =   9960
   Icon            =   "frmSelFilesToUpload.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   664
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   8640
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   8640
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Files"
      Height          =   5520
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.CheckBox chkAll 
         Caption         =   "Select All"
         Height          =   195
         Left            =   7185
         TabIndex        =   3
         Top             =   270
         Width           =   1065
      End
      Begin MSComctlLib.ListView lvwFiles 
         Height          =   4920
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   8678
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Source"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "FTP Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Select files to upload:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmSelFilesToUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UploadFiles()

    Dim k As Integer
    Dim j As Integer
    Dim File As New cFile
    Dim fCargo As Boolean
    Dim bOk  As Boolean
    Dim frm As Form
    
    For k = 1 To lvwFiles.ListItems.count
        If lvwFiles.ListItems(k).Checked Then
            bOk = True
            Exit For
        End If
    Next k
    
    If Not bOk Then
        MsgBox "You must select files to save first.", vbCritical
        Exit Sub
    End If
    
    util.Hourglass hwnd, True
    
    Me.Hide
    
    For k = 1 To lvwFiles.ListItems.count
        If lvwFiles.ListItems(k).Checked Then
            For j = 1 To Files.Files.count
                Set File = New cFile
                Set File = Files.Files.ITem(j)
                
                If Len(File.filename) > 0 Then
                    If File.IdDoc = CInt(Mid$(lvwFiles.ListItems(k).key, 2)) Then
                    
                        bOk = True
                        If lvwFiles.ListItems(k).SubItems(3) = "Modified" Then
                            If Confirma("File : " + File.filename + " has been modified. Save changes") = vbYes Then
                                For Each frm In Forms
                                    If frm.Name = "frmEdit" Then
                                        If CInt(frm.Tag) = File.IdDoc Then
                                            If Not File.SaveFile(frm, False) Then
                                                MsgBox "Failed to save file : " + File.filename, vbCritical
                                            End If
                                            Exit For
                                        End If
                                    End If
                                Next
                            Else
                                bOk = False
                            End If
                        End If
                        
                        If bOk Then
                            If Not fCargo Then
                                frmFtpFiles.InicializaArreglo
                                fCargo = True
                            End If
                        
                            If File.Ftp Then
                                Call frmFtpFiles.CargaArchivos(File.TempFile, File.filename, File.IdDoc, File.SiteName, File.User, File.RemoteFolder)
                            Else
                                Call frmFtpFiles.CargaArchivos(File.filename, VBArchivoSinPath(File.filename), File.IdDoc, "", "", "")
                            End If
                        End If
                    End If
                End If
                
                Set File = Nothing
            Next j
        End If
    Next k
        
    util.Hourglass hwnd, False
    
    If fCargo Then
        frmFtpFiles.updmulti = True
        frmFtpFiles.Show vbModal
    End If
    
    Unload Me
    
End Sub

Private Sub chkAll_Click()

    Dim ret As Boolean
    
    ret = chkAll.Value
    
    Dim k As Integer
    
    For k = 1 To lvwFiles.ListItems.count
        lvwFiles.ListItems(k).Checked = ret
    Next k
    
End Sub

Private Sub cmd_Click(Index As Integer)
   
    If Index = 0 Then
        Call UploadFiles
    Else
       Unload Me
    End If
   
End Sub


Private Sub Form_Load()

    Dim k As Integer
    Dim File As New cFile
    Dim frm As Form
    
    util.CenterForm Me
    
    util.Hourglass hwnd, True
    
    For k = 1 To Files.Files.count
        Set File = New cFile
        Set File = Files.Files.ITem(k)
        
        If Len(File.filename) > 0 Then
            If File.Ftp Then
                lvwFiles.ListItems.Add , "k" & Files.Files(k).IdDoc, Mid$(File.Caption, 5)
                lvwFiles.ListItems("k" & Files.Files(k).IdDoc).SubItems(1) = "FTP"
                lvwFiles.ListItems("k" & Files.Files(k).IdDoc).SubItems(2) = File.SiteName
            Else
                lvwFiles.ListItems.Add , "k" & Files.Files(k).IdDoc, util.VBArchivoSinPath(File.filename)
                lvwFiles.ListItems("k" & Files.Files(k).IdDoc).SubItems(1) = util.PathArchivo(File.filename)
                lvwFiles.ListItems("k" & Files.Files(k).IdDoc).SubItems(2) = ""
            End If
                
            For Each frm In Forms
                If frm.Name = "frmEdit" Then
                    If CInt(frm.Tag) = File.IdDoc Then
                        If InStr(1, frm.Caption, "*") > 0 Then
                            lvwFiles.ListItems("k" & Files.Files(k).IdDoc).SubItems(3) = "Modified"
                        Else
                            lvwFiles.ListItems("k" & Files.Files(k).IdDoc).SubItems(3) = ""
                        End If
                    End If
                End If
            Next
        End If
        
        Set File = Nothing
    Next k
    
    util.CenterForm Me
            
    util.Hourglass hwnd, False
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set frmSelFilesToUpload = Nothing
End Sub


