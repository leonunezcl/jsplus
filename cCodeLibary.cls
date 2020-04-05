VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCodeLibary"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_DataPath As String


Public Sub BrowseLibrary()
   frmMain.Show vbModal
End Sub

Public Sub CategoryWizard()
   frmCategory.Show vbModal
End Sub


Public Sub CheckDataPath()

   CrearDirectorio m_DataPath
   
End Sub

Public Sub LanguageWizard()
   frmLanguage.Show vbModal
End Sub


Public Sub SaveCode()
   frmSaveCode.Show vbModal
End Sub





Public Property Get SetDataPath() As String
   SetDataPath = m_DataPath
   
End Property

Public Property Let SetDataPath(ByVal pDataPath As String)
   m_DataPath = pDataPath
   gDataPath = m_DataPath
End Property
