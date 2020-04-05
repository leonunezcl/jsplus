Attribute VB_Name = "modLibrary"
Option Explicit

'Original Coding by: Solomon Manalo
Public Enum LanguageOptions
       [New Language] = 0
       [Rename Language] = 1
       [Remove Language] = 2
End Enum
Public Enum CategoryOptions
       [New Category] = 0
       [Rename Category] = 1
       [Remove Category] = 2
End Enum

Public Function Toggle_Language(Options As LanguageOptions, SelectedL As String, Optional NewL As String) As Boolean
Dim cFolder As New FolderSysObject
Dim path As String


If Options = [New Language] Then
   path = DATAPATH & NewL
   If IsLanguageExist(path) = True Then
      Toggle_Language = False
   Else
      'Create
      fso.CreateFolder path
      If IsLanguageExist(path) = True Then
          Toggle_Language = True
      Else
          Toggle_Language = False
      End If
   End If
End If

If Options = [Rename Language] Then
   path = DATAPATH & NewL
   If IsLanguageExist(path) = True Then
      Toggle_Language = False
   Else
      path = DATAPATH & SelectedL
      If IsLanguageExist(path) = True Then
         'Do Rename
         cFolder.RenameFolder path, NewL
         'Check if Renamed Folder Exists
         path = DATAPATH & NewL
         Toggle_Language = IsLanguageExist(path)
      End If
   End If
End If

If Options = [Remove Language] Then
   path = DATAPATH & SelectedL
   If IsLanguageExist(path) = False Then
      Toggle_Language = False
   Else
    path = DATAPATH & SelectedL
    If IsLanguageExist(path) = True Then
       'Do Delete
       fso.DeleteFolder path, True
       Toggle_Language = True
    End If
   End If
End If
End Function

Public Function IsLanguageExist(LanguageName As String) As Boolean
Dim path As String
IsLanguageExist = fso.FolderExists(LanguageName)
End Function

Public Function IsCategoryExist(Category As String) As Boolean
Dim path As String
IsCategoryExist = fso.FolderExists(Category)
End Function

Public Function Toggle_Category(Options As CategoryOptions, SelectedL, SelectedC As String, Optional NewC As String) As Boolean
Dim cFolder As New FolderSysObject
Dim path As String


If Options = [New Category] Then
   path = DATAPATH & SelectedL & "\" & NewC
   If IsCategoryExist(path) = True Then
      Toggle_Category = False
   Else
      'Create
      fso.CreateFolder path
      If IsCategoryExist(path) = True Then
          Toggle_Category = True
      Else
          Toggle_Category = False
      End If
   End If
End If

If Options = [Rename Category] Then
   path = DATAPATH & SelectedL & "\" & NewC
   If IsCategoryExist(path) = True Then
      Toggle_Category = False
   Else
     path = DATAPATH & SelectedL & "\" & SelectedC
      
      If IsCategoryExist(path) = True Then
         'Do Rename
         path = DATAPATH & SelectedL & "\" & SelectedC
         cFolder.RenameFolder path, NewC
         'Check if Renamed Folder Exists
         path = DATAPATH & SelectedL & "\" & NewC
         Toggle_Category = IsCategoryExist(path)
      End If
   End If
End If

If Options = [Remove Category] Then
   path = DATAPATH & SelectedL & "\" & SelectedC
   If IsCategoryExist(path) = False Then
      Toggle_Category = False
   Else
    path = DATAPATH & SelectedL & "\" & SelectedC
    If IsCategoryExist(path) = True Then
       'Do Delete
       fso.DeleteFolder path, True
       Toggle_Category = True
    End If
   End If
End If
End Function

