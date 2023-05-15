VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLSettingsPerTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Private Const FIRST_CHILD As String = "TableTransferTool"
Private Const ATTR_VALUE As String = "Value"
Private Const DEFAULT_XML As String = "<Table Name='Table1'><PreferredDirection Value='Source'/><KeyColumn Value='MyID'/><StarredColumns><StarredColumn>Column A</StarredColumn><StarredColumn>Column C</StarredColumn><StarredColumn>Column D</StarredColumn></StarredColumns></Table>"
    
Private Type TState
    CustomXMLParts As CustomXMLParts
    CustomXMLPart As CustomXMLPart
    XMLDoc As MSXML2.DOMDocument60
    TableName As String
End Type
Private This As TState

Public Sub Load(ByVal Workbook As Workbook, ByVal TableName As String)
    This.TableName = TableName
    
    Set This.CustomXMLParts = Workbook.CustomXMLParts
    
    If Not AlreadyExistsInWorkbook Then
        Debug.Assert False
    End If
    
    GetMyXMLDoc
    
    If HasTableNode = False Then
        AddTableNode
    End If
End Sub

Private Function AlreadyExistsInWorkbook() As Boolean
   Dim CustomXMLPart As CustomXMLPart
   For Each CustomXMLPart In This.CustomXMLParts
        If IsMyXMLDoc(CustomXMLPart) Then
            AlreadyExistsInWorkbook = True
            Exit Function
        End If
    Next CustomXMLPart
End Function

Private Function IsMyXMLDoc(ByVal CustomXMLPart As CustomXMLPart) As Boolean
    IsMyXMLDoc = (StringToXML(CustomXMLPart.XML).FirstChild.BaseName = FIRST_CHILD)
End Function

Private Function StringToXML(ByVal XMLString As String) As MSXML2.DOMDocument60
    Set StringToXML = New MSXML2.DOMDocument60
    StringToXML.LoadXML XMLString
End Function

Private Function GetMyXMLDoc() As Boolean
    Dim CustomXMLPart As CustomXMLPart
    For Each CustomXMLPart In This.CustomXMLParts
        If IsMyXMLDoc(CustomXMLPart) Then
            Set This.CustomXMLPart = CustomXMLPart
            Set This.XMLDoc = StringToXML(CustomXMLPart.XML)
            GetMyXMLDoc = True
            Exit Function
        End If
    Next CustomXMLPart
End Function

Public Sub Reset()
    ResetTableNode
End Sub

Private Property Get XPathPrefix() As String
    XPathPrefix = "/TableTransferTool/Tables/Table[@Name='" & This.TableName & "']"
End Property

Public Function GetTextValue(ByVal NodeName As String) As String
    Dim Result As CustomXMLNode
    Set Result = This.CustomXMLPart.SelectSingleNode(XPathPrefix() & "/" & NodeName & "/@" & ATTR_VALUE)
    If Result Is Nothing Then Exit Function
    GetTextValue = Result.Text
End Function

Public Sub SetTextValue(ByVal NodeName As String, ByVal vNewValue As String)
    Dim Result As CustomXMLNode
    Set Result = This.CustomXMLPart.SelectSingleNode(XPathPrefix() & "/" & NodeName)
    
    If Result Is Nothing Then
        Set Result = This.CustomXMLPart.SelectSingleNode(XPathPrefix())
        Result.AppendChildNode Name:=NodeName
        Set Result = Result.LastChild
        Result.AppendChildNode Name:=ATTR_VALUE, NodeType:=msoCustomXMLNodeAttribute, NodeValue:=vNewValue
    Else
        ' We are not properly checking if this is the correct attribute
        Debug.Assert Result.Attributes.Item(1).BaseName = ATTR_VALUE
        Result.Attributes.Item(1).Text = vNewValue
    End If
End Sub

Private Sub ResetTableNode()
    On Error Resume Next
    Dim Result As CustomXMLNode
    Set Result = This.CustomXMLPart.SelectSingleNode("/TableTransferTool/Tables")
    Result.Delete
    This.CustomXMLPart.SelectSingleNode("/TableTransferTool").AppendChildNode Name:="Tables"
    On Error GoTo 0
    
    AddTableNode
End Sub

Private Sub AddTableNode()
    If HasTableNode Then Exit Sub
    
    Dim Result As CustomXMLNode
    Set Result = This.CustomXMLPart.SelectSingleNode("//TableTransferTool/Tables")
    
    Result.AppendChildSubtree DEFAULT_XML
End Sub

Private Function HasTableNode() As Boolean
    Dim Result As CustomXMLNode
    Set Result = GetTableNode
    HasTableNode = Not Result Is Nothing
End Function

Private Function GetTableNode() As CustomXMLNode
    Dim Results As CustomXMLNodes
    Set Results = This.CustomXMLPart.SelectNodes("//TableTransferTool/Tables/Table")
    
    Dim Result As CustomXMLNode
    Dim Attr As CustomXMLNode
    For Each Result In Results
        For Each Attr In Result.Attributes
            If Attr.BaseName = "Name" And Attr.Text = This.TableName Then
                Set GetTableNode = Result
                Exit Function
            End If
        Next Attr
    Next Result
End Function

Public Function GetList(ByVal XPathFragment As String) As Collection
    Set GetList = New Collection
    Dim Results As CustomXMLNodes
    Set Results = This.CustomXMLPart.SelectNodes("/TableTransferTool/Tables/Table[@Name='Table1']/" & XPathFragment)
    
    Dim Result As CustomXMLNode
    For Each Result In Results
        GetList.Add Item:=Result.Text
    Next Result
End Function

Public Sub SetList(ByVal XPathFragment As String, ByVal Coll As Collection)
    ' TODO Need to split fragment into parent and child node names
    Dim Parent As String
    Dim Child As String
    Parent = Split(XPathFragment, "/")(0)
    Child = Split(XPathFragment, "/")(1)
    
    Dim Result As CustomXMLNode
    Set Result = This.CustomXMLPart.SelectSingleNode("/TableTransferTool/Tables/Table[@Name='Table1']/" & Parent)
    If Not Result Is Nothing Then
        Result.Delete
    End If
    
    Set Result = This.CustomXMLPart.SelectSingleNode("/TableTransferTool/Tables/Table[@Name='Table1']")
    Result.AppendChildNode Name:=Parent
    Set Result = Result.LastChild
    
    Dim ThisItem As Variant
    For Each ThisItem In Coll
        Result.AppendChildNode Name:=Child, NodeValue:=CStr(ThisItem)
    Next ThisItem
End Sub
