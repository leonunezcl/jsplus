VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************
' HTML Data Tree Class v1.00
' Copyright ©2002 by Adam Ruggles, All Rights Reserved.
'*********************************************************************
' If you like this code, please vote for it at Planet-Source-Code.com
' Thank you
'*********************************************************************
' You are free to use this code within your own applications, but you
' are expressly forbidden from selling or otherwise distributing the
' source code without prior written consent.
'*********************************************************************
' This Class Uses the cAppendString for quicker string concatenation,
' Please see that code for any Credits and Copyrights it may have
'*********************************************************************

Public Enum NodeType
  ntText = 1
  ntTag = 2
  ntEndTag = 3
End Enum

Private Type TreeNode
  ID As String
  ParentID As String
  tagname As String
  Text As String
  Child() As Long
  Pos As Long
  Line As Long
  Col As Long
  NType As NodeType
  Linked As Boolean
End Type

Private AppendStr As cAppendString
Private LinkedList() As TreeNode
Private RootList() As Long
Private SavePos As Long
Private SaveID As Long
Private bInA As Boolean
Private bUsedCrLf As Boolean
'These three Variables are for the MakeLink Function
'There where declared here so they would not be
'Reproduced in the recursion
Private LinkChild() As String
Private MatchTag As String
Private LinkCnt As Long
Option Explicit
Event TreeError(Pos As Long, Text As String, Tag As String, ID As Long)


Public Sub InitializeHTML()
  'This makes the data tree for use in all of the functions
  'This needs to be called first before using any of the
  'other procedures or functions
  ReDim LinkedList(0)
  'Produces a list of tag, end tags, and text from the HTML string
  Parse
  'Links the list into a Multiple Trees
  'and puts the Roots into the RootLists
  DoLink
End Sub

'****************************************************************
'*** The next two functions add the HTML tree into a TreeView ***
'****************************************************************
Public Sub ProduceTree(ByRef inTreeView As TreeView)
'*** Produces a list from the HTML string then links
'*** the list into a tree structure then puts it into the TreeView
  On Error Resume Next
  Dim lCnt, lCnt2 As Long
  Dim lUBound As Long
  inTreeView.Nodes.Clear
  'If CallInit = True Then InitializeHTML HTML
  For lCnt = 1 To UBound(RootList)
    With LinkedList(RootList(lCnt))
      inTreeView.Nodes.Add , , .ID, "<" & .tagname & ">", 1, 1
      If .Child(0) <> -1 Then
        'A recusive procedure that walks the tree
        'and adds the nodes to the TreeView
        AddRecTreeView RootList(lCnt), .ID, inTreeView
      End If
    End With
  Next lCnt
End Sub

Private Sub AddRecTreeView(CurIndex As Long, CurID As String, ByRef inTreeView As TreeView)
  'Recursivly Walks the tree adding each element of the
  'Tree Data structure into the TreeView Control
  Dim lCnt As Long
  For lCnt = UBound(LinkedList(CurIndex).Child) To 0 Step -1
    With LinkedList(LinkedList(CurIndex).Child(lCnt))
      If .NType = ntTag Then
        inTreeView.Nodes.Add CurID, tvwChild, .ID, "<" & .tagname & ">", 1, 1
        If .Child(0) <> -1 Then
          AddRecTreeView LinkedList(CurIndex).Child(lCnt), .ID, inTreeView
        End If
      ElseIf .NType = ntEndTag Then
        inTreeView.Nodes.Add CurID, tvwChild, .ID, "</" & .tagname & ">", 1, 1
      End If
    End With
  Next lCnt
End Sub

Private Sub Parse()
  '*** This produces a List but doesn't link the list into a tree
  Dim lBegin, lEnd, lUBound, lLenOfStr, lCnt As Long
  Dim sTemp, sStrip() As String
  Dim Line As Long
  
  Dim strHTML As New cStringBuilder
  Dim nFreeFile As Long
  
  nFreeFile = FreeFile
  
  Open StripPath(App.Path) & "htmltemp.txt" For Input As #nFreeFile
   strHTML.Append Input(LOF(nFreeFile), nFreeFile)
  Close #nFreeFile
    
  HTML = strHTML.ToString
  
  lLenOfStr = Len(HTML)
  lBegin = 1
  lEnd = 1
  Line = 1
  
  'Go through each "<" & ">" to parse the file into pieces
  Do
    lBegin = InStr(lEnd, HTML, "<", vbBinaryCompare)
    If lBegin = 0 Then Exit Do
    '*** Everything here is not a Tag (ie. Not "<" * ">")
    If lBegin <> lEnd Then
      sTemp = Replace(Trim(Mid$(HTML, lEnd + 1, lBegin - lEnd - 1)), _
        vbTab, "", , , vbBinaryCompare)
      
      If Replace(sTemp, vbCrLf, "", , , vbBinaryCompare) <> "" And _
         Replace(sTemp, vbLf, "", , , vbBinaryCompare) <> "" Then
        'This breaks up each piece of the text by crlf
        sStrip = Split(sTemp, vbLf, , vbBinaryCompare)
        For lCnt = 0 To UBound(sStrip)
          If Replace(sStrip(lCnt), vbCr, "") <> "" Then
            lUBound = UBound(LinkedList) + 1
            ReDim Preserve LinkedList(lUBound)
            With LinkedList(lUBound)
              ReDim .Child(0)
              .Child(0) = -1
              .Linked = False
              '.Pos = 0
              '.Line = Line
              .Text = Trim(Replace(sStrip(lCnt), vbCr, ""))
              .NType = ntText
              .ID = lUBound & "ID"
             End With
          End If
        Next lCnt
      End If
    End If
    lEnd = InStr(lBegin, HTML, ">", vbBinaryCompare)
    If lEnd = 0 Then Exit Do
    '*** Everything here is a Tag (ie. "<" * ">")
    If lBegin <> lEnd Then
      sTemp = Trim(Mid$(HTML, lBegin, lEnd - lBegin + 1))
      lUBound = UBound(LinkedList) + 1
      ReDim Preserve LinkedList(lUBound)
      With LinkedList(lUBound)
        ReDim .Child(0)
        .ID = lUBound & "ID"
        .Child(0) = -1
        '.Pos = lBegin
        '.Line = Line
        .Text = sTemp
        .Linked = False
        'Now we want to categorize each tag
        If Mid$(sTemp, 2, 1) = "/" Then
          .NType = ntEndTag
          .tagname = UCase(Mid$(sTemp, 3, Len(sTemp) - 3))
        ElseIf Mid$(sTemp, 2, 1) = "!" Then
          .tagname = UCase(Mid$(sTemp, 2, Len(sTemp) - 2))
          .NType = ntTag
        Else
          .NType = ntTag
          .tagname = Mid$(sTemp, 2, Len(sTemp) - 2)
          sStrip = Split(.tagname, " ", , vbBinaryCompare)
          .tagname = UCase(sStrip(0))
        End If
      End With
      'Check For a Script and Handle it differently
      If (LinkedList(lUBound).tagname = "SCRIPT" Or _
         LinkedList(lUBound).tagname = "STYLE") _
         And LinkedList(lUBound).NType = ntTag Then
        'I used lCnt as a temp variable in the following lines
        lCnt = lEnd
        If LinkedList(lUBound).tagname = "STYLE" Then
          lEnd = InStr(lEnd + 1, HTML, "</style>", vbTextCompare)
        Else
          lEnd = InStr(lEnd + 1, HTML, "</script>", vbTextCompare)
        End If
        sTemp = Trim(Mid$(HTML, lCnt + 1, lEnd - lCnt - 1))
        'I assigned the script the same as any text except it is not
        'parsed into smaller text fields on the tree.  This is for
        'the format function. We don't want to mess with the script
        'formatting
        sStrip = Split(sTemp, vbLf, , vbBinaryCompare)
        For lCnt = 0 To UBound(sStrip)
          If Replace(sStrip(lCnt), vbCr, "") <> "" Then
            lUBound = UBound(LinkedList) + 1
            ReDim Preserve LinkedList(lUBound)
            With LinkedList(lUBound)
              ReDim .Child(0)
              .Child(0) = -1
              .Linked = False
              '.Pos = 0
              '.Line = Line
              .Text = Trim(Replace(sStrip(lCnt), vbCr, ""))
              .NType = ntText
              .ID = "ID" & lUBound
             End With
          End If
        Next lCnt
      End If
    End If
    Line = Line + 1
  Loop Until lEnd > lLenOfStr Or lEnd = 0
  
  Set strHTML = Nothing
  
End Sub

Private Sub DoLink()
  'This link fills the Child Property
  'so the tree after this function is created
  ReDim RootList(0)
  Dim lCnt As Long
  Dim lCnt2 As Long
  For lCnt = 1 To UBound(LinkedList)
    With LinkedList(lCnt)
      'First we will search for an end tag
      If .NType = ntEndTag Then
        MatchTag = .tagname
        ReDim LinkChild(0)
        LinkChild(0) = lCnt
        'SavePos = .Pos
        SaveID = lCnt
        'Now we recusivly move up the tree to find
        'The parent
        .ParentID = MakeLink(lCnt - 1)
        If .ParentID <> "" Then .Linked = True
      End If
    End With
  Next lCnt
  'Now we gather all of the nodes with out parents
  'and put a pointer to them in the RootList array
  For lCnt = 1 To UBound(LinkedList)
    With LinkedList(lCnt)
      If .Linked = False Then
        lCnt2 = UBound(RootList) + 1
        ReDim Preserve RootList(lCnt2)
        RootList(lCnt2) = lCnt
      End If
    End With
  Next lCnt
End Sub

Private Function MakeLink(CurIndex As Long) As String
  'This function makes the links
  If CurIndex = 0 Then
    'If we've hit the top of the tree then we couldn't
    'find the begin tag for an end tag so we generate
    'an error
    RaiseEvent TreeError(SavePos, "End tag without matching begin tag", MatchTag, SaveID)
    MakeLink = ""
    Exit Function
  End If
  With LinkedList(CurIndex)
    'We are looking for the top tag
    If .tagname = MatchTag And .Child(0) = -1 And .NType = ntTag Then
      'If we have found the top tag then we link it's Children
      ReDim .Child(UBound(LinkChild))
      'This can be done faster but it works
      'We are copying one array to another
      For LinkCnt = 0 To UBound(LinkChild)
        .Child(LinkCnt) = LinkChild(LinkCnt)
      Next LinkCnt
      'Now return the top tag ID
      MakeLink = .ID
    Else
      If .Linked = False Then
        'we havn't found the top tag and we have no parent yet (unlinked)
        LinkCnt = UBound(LinkChild) + 1
        'Now add this tag to the child tag array
        ReDim Preserve LinkChild(LinkCnt)
        LinkChild(LinkCnt) = CurIndex
        'Call MakeLink again and find the parent
        .ParentID = MakeLink(CurIndex - 1)
        'Return the top tag ID
        MakeLink = .ParentID
        If .ParentID <> "" Then .Linked = True
      Else
        'If we have already been linked we still are looking
        'for the top tag so we call MakeLink again
        MakeLink = MakeLink(CurIndex - 1)
      End If
    End If
  End With
End Function

Private Sub Class_Initialize()
  ReDim LinkedList(0)
  ReDim RootList(0)
  ReDim LinkChild(0)
  MatchTag = ""
  LinkCnt = 0
End Sub
