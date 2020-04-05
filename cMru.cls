VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************
' Name: Class_MRUFiles
' Description:An object to maintain the
'     last files opened opened by a program.
'class_MRU gives you a quick and reliable way To provide a Most-Recently Used (MRU) file list In your application.
' By: Waty Thierry
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=1070&lngWId=1'for details.'**************************************

' #VBIDEUtils#**************************
'     **********************************
' * Programmer Name : Waty Thierry
' * Web Site : www.geocities.com/Researc
'     hTriangle/6311/
' * E-Mail: waty.thierry@usa.net
' * Date : 25/09/98
' * Time : 11:15
' * Module Name : class_MRUFiles
' * Module Filename : MRU.cls
' **************************************
'     ********************************
' * Comments : An object to maintain the
'     last files opened
' * opened by a program.
' * class_MRU gives you a quick and reli
'     able way to provide a
' * Most-Recently Used (MRU) file list i
'     n your application.
' *
' *
' **************************************
'     ********************************
Option Explicit
Private colMRUFiles As New Collection
Private Num As Integer

Public Sub Add(file As String)


    If colMRUFiles.count = 0 Then
        colMRUFiles.Add file, UCase(file)
    Else
        On Error Resume Next
        colMRUFiles.Remove UCase(file)
        colMRUFiles.Add file, UCase(file), 1


        If colMRUFiles.count > Num Then
            colMRUFiles.Remove colMRUFiles.count
        End If

    End If

End Sub


Public Sub Clear()

    ' *** Clears all files from the list.

    Do While colMRUFiles.count > 0
        colMRUFiles.Remove 1
    Loop

End Sub


Public Property Get count() As Long

    ' *** Returns the number of files in the
    '     list.
    count = colMRUFiles.count
End Property


Public Property Get ITem(i As Integer) As String

    ' *** Returns the nth item from the list
    '
    On Error GoTo ItemError
    ITem = colMRUFiles(i)
    Exit Property
ItemError:
    ITem = ""
End Property


Public Sub Load(Optional AppName As Variant)

    Dim v As Variant
    Dim i As Integer
    Dim J As Integer
    Dim AppN As String

    If IsMissing(AppName) Then
        AppN = App.ProductName
    Else
        AppN = CStr(AppName)
    End If

    v = GetAllSettings(AppN, "colMRUFiles")


    If Not IsEmpty(v) Then
        i = UBound(v, 1)
        Me.Clear
        colMRUFiles.Add v(i, 1), UCase(v(i, 1))


        For J = i - 1 To LBound(v, 1) Step -1
            colMRUFiles.Add v(J, 1), UCase(v(J, 1)), 1
        Next J

    End If

End Sub


Public Property Get Number() As Integer

    ' *** Gets the maximum size of the list.
    '
    Number = Num
End Property


Public Property Let Number(i As Integer)

    ' *** Sets the maximum size of the list.
    '
    Num = i
End Property


Public Sub Remove(file As String)

    On Error Resume Next
    colMRUFiles.Remove UCase(file)
End Sub


Public Sub Save(Optional AppName As Variant)

    Dim i As Integer
    Dim AppN As String
    On Error Resume Next

    If IsMissing(AppName) Then
        AppN = App.ProductName
    Else
        AppN = CStr(AppName)
    End If

    DeleteSetting AppN, "colMRUFiles"


    For i = 1 To colMRUFiles.count
        SaveSetting AppN, "colMRUFiles", i, colMRUFiles(i)
    Next i

End Sub


Public Sub Update(F As Form)

    ' *** Note: The form must contain a menu
    '     control array
    ' ***named mnuMRUFiles that is at least
    '     as big
    ' ***as Number.
    Dim i As Long
    On Error GoTo NextStep


    For i = 1 To Num
        F.mnuMRUFiles(i).Visible = False
    Next i

NextStep:

On Error GoTo MenuEnd


If colMRUFiles.count > 0 Then
    F.mnuMRUFiles(0).Visible = True

    For i = 1 To colMRUFiles.count
        F.mnuMRUFiles(i).Caption = colMRUFiles(i)
        F.mnuMRUFiles(i).Visible = True
    Next i


    Do
        F.mnuMRUFiles(i).Visible = False
        i = i + 1
    Loop

Else
    i = 0


    Do
        F.mnuMRUFiles(i).Visible = False
        i = i + 1
    Loop

End If

MenuEnd:
End Sub


Private Sub Class_Initialize()

    Num = 5
End Sub
