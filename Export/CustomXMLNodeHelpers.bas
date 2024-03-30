Attribute VB_Name = "CustomXMLNodeHelpers"
'@Folder "PersistentStorage.XMLSettings"
Option Explicit

Private Const ITEM_NODE As String = "Item"

Public Sub UpsertText(ByVal CustomXMLPart As CustomXMLPart, ByVal XPath As String, ByVal vNewValue As String)
    Dim Result As CustomXMLNode
    Set Result = CustomXMLPart.SelectSingleNode(XPath)
    If Result Is Nothing Then
        Set Result = GetOrCreateXPath(CustomXMLPart, XPath)
    End If
    Result.Text = vNewValue
End Sub

Public Sub UpsertCollection(ByVal CustomXMLPart As CustomXMLPart, ByVal XPath As String, ByVal vNewCollection As Collection)
    Dim Result As CustomXMLNode
    Set Result = CustomXMLPart.SelectSingleNode(XPath)
    If Not Result Is Nothing Then
        Result.Delete
    End If
    
    '@Ignore AssignmentNotUsed
    Set Result = GetOrCreateXPath(CustomXMLPart, XPath)
    
    Dim Item As Variant
    For Each Item In vNewCollection
        Result.AppendChildNode Name:=ITEM_NODE, NodeValue:=Item
    Next Item
End Sub

Private Function GetOrCreateXPath(ByVal CustomXMLPart As CustomXMLPart, ByVal XPath As String) As CustomXMLNode
    Dim Tokens() As String
    Tokens = Split(XPath, "/")
    
    Dim parent As CustomXMLNode
    Set parent = CustomXMLPart.SelectSingleNode("/" & Tokens(1))
    Debug.Assert Not parent Is Nothing
    
    Dim Result As CustomXMLNode
    Dim i As Long
    For i = 2 To UBound(Tokens)
        Set Result = parent.SelectSingleNode(Tokens(i))
        If Result Is Nothing Then
            Set Result = AppendXPathToken(parent, Tokens(i))
        End If
        
        Set parent = Result
    Next i
    
    Set GetOrCreateXPath = Result
End Function

Private Function AppendXPathToken(ByVal CustomXMLNode As CustomXMLNode, ByVal XPathToken As String) As CustomXMLNode
    Dim Result As CustomXMLNode
    
    Dim NodeName As String
    NodeName = Split(XPathToken, "[")(0)
    
    CustomXMLNode.AppendChildNode Name:=NodeName
    Set Result = CustomXMLNode.LastChild
    
    Dim Delimeters(1 To 3) As Long
    Delimeters(1) = InStr(XPathToken, "[@")
    Delimeters(2) = InStr(XPathToken, "='")
    Delimeters(3) = InStr(XPathToken, "']")
    
    If Delimeters(1) > 0 Then
        Dim AttrName As String
        Dim AttrValue As String
        AttrName = Mid$(XPathToken, Delimeters(1) + 2, Delimeters(2) - Delimeters(1) - 2)
        AttrValue = Mid$(XPathToken, Delimeters(2) + 2, Delimeters(3) - Delimeters(2) - 2)
        Result.AppendChildNode Name:=AttrName, NodeType:=msoCustomXMLNodeAttribute, NodeValue:=AttrValue
    End If
    
    Set AppendXPathToken = Result
End Function

