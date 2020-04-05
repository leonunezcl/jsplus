VERSION 5.00
Object = "{04DE47C8-1CE9-420F-ABED-109D480907D3}#1.2#0"; "PropertyWindow8.ocx"
Begin VB.Form frmRunTidy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tidy"
   ClientHeight    =   7575
   ClientLeft      =   4365
   ClientTop       =   2340
   ClientWidth     =   4500
   Icon            =   "frmRunTidy.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   5
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ok"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox txtTsk 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   45
      MaxLength       =   40
      TabIndex        =   0
      Top             =   225
      Width           =   4425
   End
   Begin PropertyWindow8.PropertyWindow PropertyWindow1 
      Height          =   6195
      Left            =   30
      TabIndex        =   1
      Top             =   765
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   10927
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
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tidy Properties"
      Height          =   195
      Index           =   1
      Left            =   30
      TabIndex        =   3
      Top             =   525
      Width           =   1050
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Task Name"
      Height          =   195
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   825
   End
End
Attribute VB_Name = "frmRunTidy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public task As String
Public file As String
Private Sub setear_propiedades()

    Dim linea As String
    Dim nFreeFile As Long
    Dim Archivo As String
    Dim k As Integer
    Dim j As Integer
    Dim propiedad As String
    Dim valor As String
    Dim prop As Property
    'Dim itmx As ListItem
    Dim C As Integer
    nFreeFile = FreeFile
    
    Archivo = util.StripPath(App.Path) & "tidy\" & file
    
    If ArchivoExiste2(Archivo) Then
        Open Archivo For Input As #nFreeFile
            Do While Not EOF(nFreeFile)
                Line Input #nFreeFile, linea
                propiedad = util.Explode(linea, 1, ":")
                valor = Trim$(util.Explode(linea, 2, ":"))
                If valor = "False" Then valor = "Falso"
                If valor = "True" Then valor = "Verdadero"
                C = 1
                With PropertyWindow1.Properties
                    For k = 1 To .count
                        If .ITem(k).Name = propiedad Then
                            Set prop = .ITem(k)
                            If prop.ListItems.count > 0 Then
                                For j = 1 To prop.ListItems.count
                                    If prop.ListItems(j).Name = valor Then
                                        prop.Value = j
                                        'prop.Selected = True
                                        Exit For
                                    End If
                                Next j
                            Else
                                prop.Value = valor
                                'prop.Selected = True
                            End If
                            Set prop = Nothing
                            Exit For
                        End If
                    Next k
                End With
            Loop
        Close #nFreeFile
    Else
        MsgBox "File not found : " & file, vbCritical
    End If
    
End Sub

Private Sub tidy_config()

    Dim k As Integer
    'Dim j As Integer
    Dim nFreeFile As Long
    Dim prop As Property
    Dim ini As String
    Dim Path As String
    Dim valor As String
    
    If Len(txtTsk.Text) = 0 Then
        txtTsk.SetFocus
        Exit Sub
    End If
    
    If Not util.ValidPattern(txtTsk.Text) Then
        MsgBox "Invalid task name.", vbCritical
        txtTsk.SetFocus
        Exit Sub
    End If
    
    If Len(file) = 0 Then
        file = txtTsk.Text
    End If
    
    If InStr(file, ".") Then
        file = Left$(file, InStr(1, file, ".") - 1)
    End If
    
    Path = util.StripPath(App.Path) & "tidy\"
    file = file & ".tidy"
    nFreeFile = FreeFile
    
    util.BorrarArchivo Path & file
    
    Open Path & file For Output As #nFreeFile
        With PropertyWindow1.Properties
            For k = 1 To .count
                Set prop = .ITem(k)
                If prop.ListItems.count > 0 Then
                    If Not prop.SelectedListItem Is Nothing Then
                        Print #nFreeFile, prop.Name & ": " & prop.SelectedListItem.Name
                    End If
                Else
                    If prop.Value <> "Falso" Then
                        If prop.Value <> "" Then
                            If prop.Value = "Verdadero" Then
                                valor = "True"
                            Else
                                valor = prop.Value
                            End If
                            
                            Print #nFreeFile, prop.Name & ": " & valor
                        End If
                    End If
                End If
            Next k
        End With
    Close #nFreeFile
    
    'verificar si archivo existe en archivo de configuracion
    Dim arr_tasks() As String
    Dim found As Boolean
    
    ini = util.StripPath(App.Path) & "tidy\tasks.ini"
    get_info_section "tasks", arr_tasks, ini
    
    For k = 1 To UBound(arr_tasks)
        If LCase$(util.Explode(arr_tasks(k), 2, "|")) = LCase$(file) Then
            found = True
            Exit For
        End If
    Next k
    
    If Not found Then
        Dim Num
        Num = util.LeeIni(ini, "tasks", "num") + 1
        On Error Resume Next
        util.GrabaIni ini, "tasks", "tsk" & Num, txtTsk.Text & "|" & file
        Err = 0
    End If
    
    'actualizar lista de tareas
    frmTidyConfig.cargar_tidy
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Then
        Call tidy_config
        Unload Me
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Dim Archivo As String
    Dim arr_grupos() As String
    Dim arr_opciones() As String
    Dim arr_propiedad() As String
    Dim linea As String
    Dim elem As String
    
    Dim k As Integer
    Dim j As Integer
    Dim C As Integer
    
    util.Hourglass hwnd, True
    util.CenterForm Me
        
    txtTsk.Text = task
    If Len(txtTsk.Text) > 0 Then
        txtTsk.Locked = True
    End If
    
    Archivo = util.StripPath(App.Path) & "tidy\tidy.ini"
    
    get_info_section "groups", arr_grupos, Archivo
    get_info_section "options", arr_opciones, Archivo
    
    For k = 1 To UBound(arr_opciones)
        get_info_section arr_opciones(k), arr_propiedad, Archivo
        With PropertyWindow1.Properties
            j = Mid$(arr_propiedad(3), Len(arr_propiedad(3)), 1)
            If arr_propiedad(1) = "boolean" Then
                If arr_propiedad(2) = "no" Then
                    .Add arr_opciones(k), pwboolean, arr_grupos(j), False, , arr_opciones(k), arr_propiedad(4)
                Else
                    .Add arr_opciones(k), pwboolean, arr_grupos(j), True, , arr_opciones(k), arr_propiedad(4)
                End If
            ElseIf arr_propiedad(1) = "String" Then
                .Add arr_opciones(k), pwText, arr_grupos(j), arr_propiedad(2), , arr_opciones(k), arr_propiedad(4)
            ElseIf arr_propiedad(1) = "integer" Then
                .Add arr_opciones(k), pwNumber, arr_grupos(j), arr_propiedad(2), , arr_opciones(k), arr_propiedad(4)
            Else
                linea = util.LeeIni(Archivo, "types", arr_propiedad(1))
                If Len(linea) > 0 Then
                    C = 1
                    'lista
                    With .Add(arr_opciones(k), pwList, arr_grupos(j), arr_propiedad(2), , arr_opciones(k), arr_propiedad(4))
                        Do
                            elem = util.Explode(linea, C, "|")
                            If Len(elem) > 0 Then
                                .ListItems.Add elem, "k" & C
                            Else
                                Exit Do
                            End If
                            C = C + 1
                        Loop
                        '.Value = ListValue
                    End With
                Else
                    .Add arr_opciones(k), pwText, arr_grupos(j), arr_propiedad(2), , arr_opciones(k), arr_propiedad(4)
                End If
            End If
        End With
    Next k
    
    If Len(file) > 0 Then
        Call setear_propiedades
    End If
    
    util.Hourglass hwnd, False
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call clear_memory(Me)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmRunTidy = Nothing
End Sub


