VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Private mCurrKey As String
Private mXMLIsLoaded As Boolean
Private mCurrFileName As String
Private oDoc As DOMDocument
Private mTreeView As TreeView
Private mMaxID As Integer
Private mKeyCont As Integer
Public Sub ChangeText(NewValue As String)
    Set currXMLNode = oDoc.nodeFromID(mCurrKey)
    currXMLNode.Text = NewValue
End Sub


Public Sub SetTreeView(ByRef oTreeView As TreeView)
'sets an internal reference to the treeview object of the form
    Set mTreeView = oTreeView
End Sub

Public Property Get CurrFileName() As String
'allows to catch the current file name
    CurrFileName = mCurrFileName
End Property
Public Property Get XMLIsLoaded() As Boolean
'is there an XML loaded in the object
    XMLIsLoaded = mXMLIsLoaded
End Property

Public Property Let CurrKey(ByVal vData As String)
'when the user clicks on a node, the instance has to be updated via this property
    mCurrKey = vData
End Property


Public Property Get CurrKey() As String
' to fetch the current key
    CurrKey = mCurrKey
End Property



Public Sub DeleteNode()
    'mTreeView.Nodes.Remove (mCurrKey)
End Sub

Public Function SaveAsXMLFile(ByVal FilePath As String) As Boolean
    'save as function
    oDoc.Save FilePath
End Function

Public Function SaveXMLFile() As Boolean
    ' save the XML tree to ... an XML file
    oDoc.Save mCurrFileName
End Function

Public Function CloseXMLFile() As Boolean
    'set the internal content to nothing
    Class_Initialize
End Function

Public Function OpenXMLFile(ByVal FilePath As String) As Boolean

    'open an existing xml file and start filling the tree on the form
    Set oDoc = New DOMDocument
    oDoc.async = False
    'oDoc.validateOnParse = False 'If validation is not important, skip it
    oDoc.Load FilePath
    If oDoc.parseError.errorCode = 0 Then
         mTreeView.Nodes.Clear
         'this starts the filling process
        mKeyCont = 0
        AddTreeNode mTreeView, oDoc.documentElement
        mCurrFileName = FilePath
        mXMLIsLoaded = True
        OpenXMLFile = True
    Else
        mXMLIsLoaded = False
        OpenXMLFile = False
    End If
End Function
Private Sub AddTreeNode(ByRef objTreeView As TreeView, ByRef oElem As IXMLDOMNode, Optional ByRef oTreeNode As Node)
    
    Dim oNewNode As Node
    Dim oNodeList As IXMLDOMNodeList
    Dim i As Long
    Dim k As Long
    
    If oTreeNode Is Nothing Then
        Set oNewNode = objTreeView.Nodes.Add
    Else
        Set oNewNode = objTreeView.Nodes.Add(oTreeNode, tvwChild)
    End If
    
    oNewNode.Expanded = False
    nodeName = oElem.nodeName
    
    If oElem.hasChildNodes Then
        oNewNode.Image = 1
    Else
        oNewNode.Image = 3
    End If
    
    If oElem.nodeType = 1 Then
        use = oElem.nodeName
    Else
        use = oElem.Text
    End If
    
    'agregar el texto
    oNewNode.Text = use
    
    If oNewNode.Image = 1 Then
        If mKeyCont = 0 Then
            oNewNode.Tag = "root"
        Else
            oNewNode.Tag = "Folder"
        End If
    Else
        oNewNode.Tag = "Node"
    End If
    
    oNewNode.key = "k" & mKeyCont

    If oNewNode.Image = 1 Then
        oNewNode.Expanded = True
        oNewNode.Image = 2
    End If
    
    'agregar los atributos
    If Not oElem.Attributes Is Nothing Then
        If oElem.Attributes.Length > 0 Then
            For k = 0 To oElem.Attributes.Length - 1
                mKeyCont = mKeyCont + 1
                mTreeView.Nodes.Add oNewNode.key, tvwChild, "k" & mKeyCont, oElem.Attributes(k).nodeName & "=" & oElem.Attributes(k).Text, 4, 4
                mTreeView.Nodes("k" & mKeyCont).Tag = "Attribute"
            Next k
        End If
    End If
    
    mKeyCont = mKeyCont + 1
    
    Set oNodeList = oElem.childNodes
        
    For i = 0 To oNodeList.Length - 1
        AddTreeNode mTreeView, oNodeList.Item(i), oNewNode
    Next

End Sub
Private Sub UpdateMaxID(ByVal sID As String)
'the maximal ID is not a strict incrementing number but changes when you delete or add nodes
inID = CInt(Right(sID, Len(sID) - 2))
    If inID > mMaxID Then
        mMaxID = inID
    End If
End Sub


Private Sub Class_Initialize()
'standard class ini
    Set oDoc = Nothing
    mCurrKey = 0
    mXMLIsLoaded = False
    mCurrFileName = ""
    mcurrFilepath = ""
    mMaxID = 1
End Sub
