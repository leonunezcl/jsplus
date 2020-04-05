Attribute VB_Name = "ModResources"
Option Explicit

'Original Coding by: Solomon Manalo

Public fso As New FileSystemObject
Public itn As Node
Public itm As ListItem
Public Msg, inp As Variant
'Public crypt As New cSimpleCrypt

Public gDataPath As String
Public Const invFileName As String = ".x,/x<x>x?x:x;x'x[x]x{x}x=x*x~x`x" & """"

Public Function DataPath() As String
   DataPath = gDataPath
End Function



Public Function OpenFileDoc(LibPath, pathFilename As String) As Boolean
'Dim f As New frmDocument
Dim a As New FileSysObject

    If fso.FileExists(LibPath & pathFilename) = False Then
       OpenFileDoc = False
       Exit Function
    End If
    OpenFileDoc = True
    'f.pathf = pathFilename
    'f.Caption = fso.GetFileName(pathFilename)
    'f.txtDoc.OpenFile (LibPath & pathFilename)
    'f.txtDoc.bChange = False
    'f.Show
End Function

Public Function SaveFileDoc(LibPath As String, pathFilename As String, Data As String) As Boolean
'Dim f As New frmDocument
Dim a As New FileSysObject

    If fso.FileExists(LibPath & pathFilename) = False Then
       SaveFileDoc = False
       Exit Function
    End If
    SaveFileDoc = True
    Call SaveFile(LibPath & pathFilename, Data)
    
End Function

Public Function SaveFileAs(LibPath As String, pathFilename As String, Data As String) As Boolean
'Dim f As New frmDocument
Dim a As New FileSysObject

    SaveFileAs = True
    Call SaveFile(LibPath & pathFilename, Data)
   
End Function
Public Function IsNoFileName(tmp As String) As Boolean
If InStr(LCase(Trim(tmp)), "code-file") _
   And LCase(Left(tmp, Len("code-file"))) = "code-file" Then
   IsNoFileName = True
Else
   IsNoFileName = False
End If
End Function

Public Function ValidName(tmp As String) As Boolean
Dim inv() As String
Dim i As Integer

inv = Split(invFileName, "x")
For i = 1 To UBound(inv)
    If InStr(tmp, inv(i)) Then
       ValidName = False
       Exit For
    Else
       ValidName = True
    End If
Next i
End Function

Public Sub HL(txt As TextBox)
With txt
     .SelStart = 0
     .SelLength = Len(.Text)
     .SetFocus
End With
End Sub
