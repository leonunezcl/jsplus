Attribute VB_Name = "Clipboard_HTML"
Option Explicit

Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal cbLength As Long)
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpData As Long) As Long

Private Const m_sDescription = _
                  "Version:1.0" & vbCrLf & _
                  "StartHTML:aaaaaaaaaa" & vbCrLf & _
                  "EndHTML:bbbbbbbbbb" & vbCrLf & _
                  "StartFragment:cccccccccc" & vbCrLf & _
                  "EndFragment:dddddddddd" & vbCrLf
                  
Private m_cfHTMLClipFormat As Long

Public Function CanPasteHTML() As Boolean
Dim hMemHandle As Long

If RegisterCF() = 0 Then Exit Function

If CBool(OpenClipboard(0)) Then
   
   'Retrieve the data handle the clipboard
   hMemHandle = GetClipboardData(m_cfHTMLClipFormat)

   If CBool(hMemHandle) Then CanPasteHTML = True
   
End If

Call CloseClipboard

End Function
Function RegisterCF() As Long


   'Register the HTML clipboard format
   If (m_cfHTMLClipFormat = 0) Then
      m_cfHTMLClipFormat = RegisterClipboardFormat("HTML Format")
   End If
   RegisterCF = m_cfHTMLClipFormat
   
End Function

Public Sub PutHTMLClipboard(sHtmlFragment As String, _
   Optional sContextStart As String = "<HTML><BODY>", _
   Optional sContextEnd As String = "</BODY></HTML>")
   
   Dim sData As String
   
   If RegisterCF = 0 Then Exit Sub
   
   'Add the starting and ending tags for the HTML fragment
   sContextStart = sContextStart & "<!--StartFragment -->"
   sContextEnd = "<!--EndFragment -->" & sContextEnd
   
   'Build the HTML given the description, the fragment and the context.
   'And, replace the offset place holders in the description with values
   'for the offsets of StartHMTL, EndHTML, StartFragment and EndFragment.
   sData = m_sDescription & sContextStart & sHtmlFragment & sContextEnd
   sData = Replace(sData, "aaaaaaaaaa", _
                   Format(Len(m_sDescription), "0000000000"))
   sData = Replace(sData, "bbbbbbbbbb", Format(Len(sData), "0000000000"))
   sData = Replace(sData, "cccccccccc", Format(Len(m_sDescription & _
                   sContextStart), "0000000000"))
   sData = Replace(sData, "dddddddddd", Format(Len(m_sDescription & _
                   sContextStart & sHtmlFragment), "0000000000"))

   'Add the HTML code to the clipboard
   If CBool(OpenClipboard(0)) Then
   
      Dim hMemHandle As Long, lpData As Long
      
      hMemHandle = GlobalAlloc(0, Len(sData) + 10)
      
      If CBool(hMemHandle) Then
               
         lpData = GlobalLock(hMemHandle)
         If lpData <> 0 Then
            
            CopyMemory ByVal lpData, ByVal sData, Len(sData)
            GlobalUnlock hMemHandle
            EmptyClipboard
            SetClipboardData m_cfHTMLClipFormat, hMemHandle
                        
         End If
      
      End If
   
      Call CloseClipboard
   End If

End Sub

Public Function GetHTMLClipboard() As String

   Dim sData As String
   
   If RegisterCF = 0 Then Exit Function
   
   If CBool(OpenClipboard(0)) Then
   
      Dim hMemHandle As Long, lpData As Long
      Dim nClipSize As Long
      
      GlobalUnlock hMemHandle

      'Retrieve the data from the clipboard
      hMemHandle = GetClipboardData(m_cfHTMLClipFormat)
      
      If CBool(hMemHandle) Then
               
         lpData = GlobalLock(hMemHandle)
         If lpData <> 0 Then
            nClipSize = lstrlen(lpData)
            sData = String(nClipSize + 10, 0)
            

            Call CopyMemory(ByVal sData, ByVal lpData, nClipSize)
            
            Dim nStartFrag As Long, nEndFrag As Long
            Dim nIndx As Long
            
            'If StartFragment appears in the data's description,
            'then retrieve the offset specified in the description
            'for the start of the fragment. Likewise, if EndFragment
            'appears in the description, then retrieve the
            'corresponding offset.
            nIndx = InStr(sData, "StartFragment:")
            If nIndx Then
               nStartFrag = CLng(Mid(sData, _
                                 nIndx + Len("StartFragment:"), 10))

            End If
            nIndx = InStr(sData, "EndFragment:")
            If nIndx Then
               nEndFrag = CLng(Mid(sData, nIndx + Len("EndFragment:"), 10))
            End If
            
            'Return the fragment given the starting and ending
            'offsets
            If (nStartFrag > 0 And nEndFrag > 0) Then
               GetHTMLClipboard = Mid(sData, nStartFrag + 1, _
                                 (nEndFrag - nStartFrag))
            End If
                        
         End If
      
      End If

   
      Call CloseClipboard
   End If


End Function
