VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PersistentStorage.XMLSettings"
Option Explicit
Implements ISettings

Private Const DO_DEBUG As Boolean = False

Private Const FLAGS_NODE As String = "Flags"
Private Const SETTINGS_NODE As String = "Settings"
Private Const COLLECTIONS_NODE As String = "Collections"
Private Const COLLECTION_NODE As String = "Collection"
Private Const ITEM_NODE As String = "Item"
Private Const TRUE_SYMBOL As String = "TRUE"
Private Const FALSE_SYMBOL As String = "FALSE"

Private Type TState
    Parent As XMLSettingsModel
    XPathPrefix As String
End Type
Private This As TState

Private Property Get CustomXMLPart() As CustomXMLPart
    Debug.Assert Not This.Parent Is Nothing
    Debug.Assert Not This.Parent.CustomXMLPart Is Nothing
    Set CustomXMLPart = This.Parent.CustomXMLPart
End Property

Public Property Get XPathPrefix() As String
    XPathPrefix = This.XPathPrefix
End Property

Public Property Let XPathPrefix(ByVal vNewValue As String)
    This.XPathPrefix = vNewValue
End Property

Public Sub SetParent(ByVal Parent As XMLSettingsModel)
    Set This.Parent = Parent
End Sub

Private Function ISettings_GetFlag(ByVal FlagName As String) As Boolean
    Dim XPath As String
    XPath = XPathPrefix & "/" & FLAGS_NODE & "/" & FlagName
    
    Dim Result As CustomXMLNode
    Set Result = CustomXMLPart.SelectSingleNode(XPath)
    
    If Not Result Is Nothing Then
        Select Case UCase$(Result.Text)
            Case TRUE_SYMBOL:
                ISettings_GetFlag = True
            Case FALSE_SYMBOL:
                ISettings_GetFlag = False
            Case Else:
                Debug.Assert False
        End Select
    End If
    If DO_DEBUG Then Debug.Print "ISettings_GetFlag("; FlagName; ") = "; ISettings_GetFlag
End Function

Private Sub ISettings_SetFlag(ByVal FlagName As String, ByVal Value As Boolean)
    If DO_DEBUG Then Debug.Print "ISettings_SetFlag("; FlagName; ") = "; Value
    Dim XPath As String
    XPath = XPathPrefix & "/" & FLAGS_NODE & "/" & FlagName
    
    UpsertText CustomXMLPart, XPath, IIf(Value, TRUE_SYMBOL, FALSE_SYMBOL)
End Sub

Private Function ISettings_GetSetting(ByVal SettingName As String) As String
    Dim XPath As String
    XPath = XPathPrefix & "/" & SETTINGS_NODE & "/" & SettingName
    
    Dim Result As CustomXMLNode
    Set Result = CustomXMLPart.SelectSingleNode(XPath)
    
    If Not Result Is Nothing Then
        ISettings_GetSetting = Result.Text
    End If
End Function

Private Sub ISettings_SetSetting(ByVal SettingName As String, ByVal Value As String)
    Dim XPath As String
    XPath = XPathPrefix & "/" & SETTINGS_NODE & "/" & SettingName
    
    UpsertText CustomXMLPart, XPath, Value
End Sub

Private Function ISettings_GetCollection(ByVal CollectionName As String) As Collection
    Set ISettings_GetCollection = New Collection
    
    Dim XPath As String
    XPath = XPathPrefix & "/" & COLLECTIONS_NODE & "/" & COLLECTION_NODE & "[@Name='" & CollectionName & "']/" & ITEM_NODE
    
    Dim Results As CustomXMLNodes
    Set Results = CustomXMLPart.SelectNodes(XPath)
    
    If Not Results Is Nothing Then
        Dim Item As CustomXMLNode
        For Each Item In Results
            ISettings_GetCollection.Add Item:=Item.Text
        Next
    End If
End Function

Private Sub ISettings_SetCollection(ByVal CollectionName As String, ByVal Collection As Collection)
    Dim XPath As String
    XPath = XPathPrefix & "/" & COLLECTIONS_NODE & "/" & COLLECTION_NODE & "[@Name='" & CollectionName & "']"
    
    UpsertCollection CustomXMLPart, XPath, Collection
End Sub
