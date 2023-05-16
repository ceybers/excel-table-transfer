VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SomeSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Models")
Option Explicit
Implements ISomeSettings

Private Const FLAGS_NODE As String = "Flags"
Private Const SETTINGS_NODE As String = "Settings"
Private Const COLLECTIONS_NODE As String = "Collections"
Private Const COLLECTION_NODE As String = "Collection"

Private Type TState
    CustomXMLPart As CustomXMLPart
    XPathPrefix As String
End Type
Private This As TState

Public Property Get CustomXMLPart() As CustomXMLPart
    Set CustomXMLPart = This.CustomXMLPart
End Property

Public Property Set CustomXMLPart(ByVal vNewValue As CustomXMLPart)
    Set This.CustomXMLPart = vNewValue
End Property

Public Property Get XPathPrefix() As String
    XPathPrefix = This.XPathPrefix
End Property

Public Property Let XPathPrefix(ByVal vNewValue As String)
    This.XPathPrefix = vNewValue
End Property

Private Function ISomeSettings_GetFlag(ByVal FlagName As String) As Boolean
    Dim XPath As String
    XPath = XPathPrefix & "/" & FLAGS_NODE & "/" & FlagName
    
    Dim Result As CustomXMLNode
    Set Result = This.CustomXMLPart.SelectSingleNode(XPath)
    
    If Not Result Is Nothing Then
        ISomeSettings_GetFlag = (Result.Text = "TRUE")
    End If
End Function

Private Sub ISomeSettings_SetFlag(ByVal FlagName As String, ByVal Value As Boolean)
    Dim XPath As String
    XPath = XPathPrefix & "/" & FLAGS_NODE & "/" & FlagName
    
    UpsertText This.CustomXMLPart, XPath, IIf(Value, "TRUE", "FALSE")
End Sub

Private Function ISomeSettings_GetSetting(ByVal SettingName As String) As String
    Dim XPath As String
    XPath = XPathPrefix & "/" & SETTINGS_NODE & "/" & SettingName
    
    Dim Result As CustomXMLNode
    Set Result = This.CustomXMLPart.SelectSingleNode(XPath)
    
    If Not Result Is Nothing Then
        ISomeSettings_GetSetting = Result.Text
    End If
End Function

Private Sub ISomeSettings_SetSetting(ByVal SettingName As String, ByVal Value As String)
    Dim XPath As String
    XPath = XPathPrefix & "/" & SETTINGS_NODE & "/" & SettingName
    
    UpsertText This.CustomXMLPart, XPath, Value
End Sub

Private Function ISomeSettings_GetCollection(ByVal CollectionName As String) As Collection
    Set ISomeSettings_GetCollection = New Collection
    
    Dim XPath As String
    XPath = XPathPrefix & "/" & COLLECTIONS_NODE & "/" & COLLECTION_NODE & "[@Name='" & CollectionName & "']/Item"
    
    Dim Results As CustomXMLNodes
    Set Results = This.CustomXMLPart.SelectNodes(XPath)
    
    If Not Results Is Nothing Then
        Dim Item As CustomXMLNode
        For Each Item In Results
            ISomeSettings_GetCollection.Add Item:=Item.Text
        Next
    End If
End Function

Private Sub ISomeSettings_SetCollection(ByVal CollectionName As String, ByVal Collection As Collection)
    Dim XPath As String
    XPath = XPathPrefix & "/" & COLLECTIONS_NODE & "/" & COLLECTION_NODE & "[@Name='" & CollectionName & "']"
    
    UpsertCollection This.CustomXMLPart, XPath, Collection
End Sub



