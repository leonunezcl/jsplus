Attribute VB_Name = "Module8"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''©Rd'

' MRUFiles.bas - module to manage Most Recently Used Files menu.

' This module demonstrates the use of VB's internal registry
' functions:
'                SaveSetting
'                GetSetting
'                GetAllSettings
'                DeleteSetting

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Must limit the max number of file's we record in the Registry
Private Const MAX_RECENT_FILES = 50

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Declare our Windows API function used for file path validation (vb5 compat)
Private Declare Function GetAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpSpec As String) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The mnuRecent(0) menu object must be created within your Form at
' design time. If your parent Form is not named frmMain you will need
' to replace all instances of frmMain in this module.

' All effort has been made to eliminate errors from this module, and
' so these functions should operate reliably and without any unexpected
' runtime exceptions. None-the-less, you should use error handlers in
' all procedures that make calls to these functions before compiling.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Sample Recent File Menu Click procedure in a Form.
' Note - mnuRecent(0) menu item must be created at design time.

'Private Sub mnuRecent_Click(Index As Integer)
'    Dim FileSpec As String
'    FileSpec = mnuRecent(Index).Caption
'
'    OpenAFile (FileSpec)
'    'rtfRichTextBox.LoadFile FileSpec
'End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Initialize the menu within a Form_Load or Sub Main procedure.

'Private Sub MDIForm_Load()
'    LoadRecentFiles
'End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Sample Form Unload procedure in a child form of an MDI application.

'Private Sub Form_Unload(Cancel As Integer)
'    Dim FileSpec As String
'    FileSpec = Me.Caption
'    'FileSpec = rtfRichTextBox.FileName
'
'    ' Add to recent file list
'    AddRecentFile (FileSpec)
'End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Variable used by the NumRecentFiles property to set the number
' of recent files displayed.
Private mNumRecentFiles As Long

Public Property Let NumRecentFiles(ByVal New_Num As Long)
Attribute NumRecentFiles.VB_Description = "Public property allowing the number of recent files displayed in the Recent Files sub-menu to be set as required."
    ' Public property allowing the number of recent files displayed
    ' in the Recent Files sub-menu to be set as required.

    On Error GoTo SkipError
    If (New_Num > MAX_RECENT_FILES) Then
        New_Num = MAX_RECENT_FILES ' Max file's kept track of.
    ElseIf (New_Num < 1) Then
        New_Num = 1
    End If
    mNumRecentFiles = New_Num
    SaveSetting App.Title, "Settings", "NumRecentFiles", CStr(New_Num)
    LoadRecentFiles

SkipError:
    If (Err <> 0) Then
        Debug.Print "Error in MRUFiles.Property Let NumRecentFiles:"
        Debug.Print Err.Number; "-"; Err.description
        Beep
    End If
End Property

Public Property Get NumRecentFiles() As Long
    ' Returns the NumRecentFiles value stored in the Windows registry
    ' using VB's internal GetSetting function. Defaults to 12 entries.

    On Error GoTo SkipError
    mNumRecentFiles = CLng(GetSetting(App.Title, "Settings", "NumRecentFiles", "12"))
    NumRecentFiles = mNumRecentFiles

SkipError:
    If (Err <> 0) Then
        Debug.Print "Error in MRUFiles.Property Get NumRecentFiles:"
        Debug.Print Err.Number; "-"; Err.description
        Beep
    End If
End Property

Private Function FileExists(sFileSpec As String) As Boolean
    ' The FileExists function demonstrates a good use of the Win API
    ' GetAttributes function. This is about the most reliable file
    ' existance test under the sun!

    On Error GoTo SkipError
    If Len(sFileSpec) <> 0 Then
        Dim Attribs As Long
        Attribs = GetAttributes(sFileSpec)
        If (Attribs <> -1) Then ' File exists, and is not a dir
            FileExists = ((Attribs And vbDirectory) <> vbDirectory)
        End If
    End If

SkipError:
    If (Err <> 0) Then
        Debug.Print "Error in MRUFiles.Function FileExists:"
        Debug.Print Err.Number; "-"; Err.description
        Beep
    End If
End Function

Sub LoadRecentFiles()
Attribute LoadRecentFiles.VB_Description = "Validates entries and refreshes the Recent Files menu."
    ' This procedure demonstrates the use of the GetAllSettings function,
    ' which returns a 2-dimensional array of subkeys and their respective
    ' values from the Windows registry. In this case, the registry subkeys
    ' contain the names of the most recently opened files. Note that the
    ' GetAllSettings function stores the array items in the order that
    ' they were written to the registry, not alphabetic or numeric order.

    ' The procedure validates the names of the recently opened files in
    ' case the files have been deleted or renamed, or are otherwise un-
    ' available (removable media for example). The entries are added to
    ' the Recent Files menu until the number of entries is equal to the
    ' NumRecentFiles property defined in this module.

    ' The SaveSetting function writes all valid names back to the Windows
    ' registry in original order. The DeleteSetting function is used here
    ' to remove redundant keys from the Recent Files section but can also
    ' delete whole sections by not specifying a subkey name (by omitting
    ' the third argument).

    On Error GoTo SkipError
    Dim arrFiles() As String ' Used to store returned array.

    ' If entries exist, get all recent files from the registry.
    If Len(GetSetting(App.Title, "Recent Files", "File1")) <> 0 Then
        arrFiles = GetAllSettings(App.Title, "Recent Files")
    Else ' There are no recent files so exit.
        Exit Sub
    End If

    Dim strFile As String, ub As Long
    Dim idx As Long, Num As Long

    'ResetRecentMenu
    ub = UBound(arrFiles, 1)
    
    For idx = 0 To ub
        'Debug.Print arrFiles(Idx, 0) & " - " & arrFiles(Idx, 1)
        strFile = arrFiles(idx, 1)
        If FileExists(strFile) Then
            If (Num < mNumRecentFiles) Then ShowRecentFile Num, strFile
            Num = Num + 1
            SaveSetting App.Title, "Recent Files", "File" & Num, strFile
        End If
    Next idx

    ' If invalid entries were removed
    Do While Num <= ub
        Num = Num + 1
        DeleteSetting App.Title, "Recent Files", "File" & Num
    Loop

SkipError:
    If (Err <> 0) Then
        Debug.Print "Error in MRUFiles.Sub LoadRecentFile:"
        Debug.Print Err.Number; "-"; Err.description
        Beep
    End If
End Sub

Sub AddRecentFile(sFileSpec As String)
Attribute AddRecentFile.VB_Description = "This procedure uses the SaveSettings statement to add the names of recently opened files to the System registry."
    ' This procedure uses the SaveSetting statement to add the names
    ' of recently opened files to the System registry.

    On Error GoTo SkipError
    If Not FileExists(sFileSpec) Then Exit Sub

    Dim arrFiles() As String ' Used to store returned array.

    ' If entries exist, get all recent files from the registry.
    If Len(GetSetting(App.Title, "Recent Files", "File1")) <> 0 Then
        arrFiles = GetAllSettings(App.Title, "Recent Files")
    Else
        ' There are no previous recent files so skip to AddFile.
        GoTo AddFile
    End If

    Dim strFile As String, ub As Long
    Dim idx As Long, Num As Long

    ' If this file is already top-most in the list then exit.
    If (LCase$(arrFiles(0, 1)) = LCase$(sFileSpec)) Then Exit Sub

    ub = UBound(arrFiles, 1)
    Num = 1

    For idx = 0 To ub
        strFile = arrFiles(idx, 1)
        ' Avoid repeated entries.
        If (LCase$(strFile) <> LCase$(sFileSpec)) Then
            ' Copy recent file 1 to recent file 2, and so on.
            If (Num < mNumRecentFiles) Then ShowRecentFile Num, strFile
            Num = Num + 1
            SaveSetting App.Title, "Recent Files", "File" & Num, strFile
            ' Limit how many recent files we keep track of.
            If (Num = MAX_RECENT_FILES) Then Exit For
        Else
            ' Because this file already existed in the list, only the
            ' files above it need to be moved down one, and those below
            ' it are un-affected, so we exit the for loop and add this
            ' file to the top of the list.
            Exit For
        End If
    Next idx

AddFile:
    ' Write the current file to first recent file.
    SaveSetting App.Title, "Recent Files", "File1", sFileSpec
    ShowRecentFile 0, sFileSpec

SkipError:
    If (Err <> 0) Then
        Debug.Print "Error in MRUFiles.Sub AddRecentFile:"
        Debug.Print Err.Number; "-"; Err.description
        Beep
    End If
End Sub

Private Sub ResetRecentMenu()
Attribute ResetRecentMenu.VB_Description = "Sets/resets the File menu's Recent Files sub-menu. The first time called it Loads all menu indicies from 1 to the MaxRecentFiles constant."
    ' This local procedure sets/resets the File menu's Recent Files sub-menu.
    ' The first time this procedure is called it Loads all menu indices from
    ' 1 to the number specified by the NumRecentFiles property.

    On Error GoTo SkipError
    Dim idx As Long, Cnt As Long
    Cnt = frmMain.mnuRecent.count

    frmMain.mnuRecent(0).Caption = "(No Files)"
    frmMain.mnuRecent(0).Enabled = False
    frmMain.mnuRecent(0).Visible = True

    For idx = 1 To NumRecentFiles - 1
        If (Cnt <= idx) Then
            ' Load menu items not yet loaded.
            Load frmMain.mnuRecent(idx)
        End If
        frmMain.mnuRecent(idx).Visible = False
    Next idx

SkipError:
    If (Err <> 0) Then
        Debug.Print "Error in MRUFiles.Sub ResetRecentMenu:"
        Debug.Print Err.Number; "-"; Err.description
        Beep
    End If
End Sub

Private Sub ShowRecentFile(idx As Long, sRecentFile As String)
Attribute ShowRecentFile.VB_Description = "This local procedure adds the recent file it recieves to the File menu's Recent Files sub-menu at the index position specified."
    ' This local procedure adds the recent file it recieves to the File
    ' menu's Recent Files sub-menu at the index position specified.

    On Error GoTo SkipError
    frmMain.mnuRecent(idx).Caption = sRecentFile
    frmMain.mnuRecent(idx).Enabled = True
    frmMain.mnuRecent(idx).Visible = True

SkipError:
    If (Err <> 0) Then
        Debug.Print "Error in MRUFiles.Sub ShowRecentFile:"
        Debug.Print Err.Number; "-"; Err.description
        Beep
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''©Rd'
