Attribute VB_Name = "CustomXMLNodeHelpers"
'@Folder("TestCustomXMLPart")
Option Explicit

Public Function InsertMissingNode(ByVal CustomXMLPart As CustomXMLPart, ByVal XPath As String) As CustomXMLNode
    Dim Tokens() As String
    Tokens = Split(XPath, "/")
    
    Dim Parent As CustomXMLNode
    Set Parent = CustomXMLPart.SelectSingleNode("/" & Tokens(1))
    'Stop
    
    Dim Result As CustomXMLNode
    'Dim ThisToken As String
    'Dim ThisAttr As String
    Dim i As Long
    For i = 2 To UBound(Tokens)
        Set Result = Parent.SelectSingleNode(Tokens(i)) ' Full token including attribute check
        If Result Is Nothing Then
            Dim NodeName As String
            NodeName = Split(Tokens(i), "[")(0)
            Parent.AppendChildNode Name:=NodeName
            Set Result = Parent.LastChild
            
            Dim HasAttr As Long
            HasAttr = InStr(Tokens(i), "[@")
            If HasAttr > 0 Then
                Dim AttrName As String
                Dim AttrValue As String
                AttrName = Mid$(Tokens(i), HasAttr + 2)
                AttrName = Left$(AttrName, Len(AttrName) - 2)
                AttrValue = Mid$(AttrName, InStr(AttrName, "'") + 1)
                AttrName = Left$(AttrName, InStr(AttrName, "=") - 1)
                Result.AppendChildNode Name:=AttrName, NodeType:=msoCustomXMLNodeAttribute, NodeValue:=AttrValue
            End If
        End If
        Set Parent = Result
    Next i
    
    Set InsertMissingNode = Result
End Function

Public Sub UpsertText(ByVal CustomXMLPart As CustomXMLPart, ByVal XPath As String, ByVal vNewValue As String)
    Dim Result As CustomXMLNode
    Set Result = CustomXMLPart.SelectSingleNode(XPath)
    If Result Is Nothing Then
        Set Result = InsertMissingNode(CustomXMLPart, XPath)
    End If
    Result.Text = vNewValue
End Sub

Public Sub UpsertCollection(ByVal CustomXMLPart As CustomXMLPart, ByVal XPath As String, ByVal vNewCollection As Collection)
    Dim Result As CustomXMLNode
    Set Result = CustomXMLPart.SelectSingleNode(XPath)
    If Not Result Is Nothing Then
        Result.Delete
    End If
    
    Set Result = InsertMissingNode(CustomXMLPart, XPath)
    
    Dim Item As Variant
    For Each Item In vNewCollection
        Result.AppendChildNode Name:="Item", NodeValue:=Item
    Next Item
End Sub
