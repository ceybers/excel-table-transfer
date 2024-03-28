VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule EmptyMethod, ProcedureNotUsed
'@PredeclaredId
'@Folder "PersistentStorage.XMLSettings"
Option Explicit
Implements ISettings

Private Const FLAGS_NODE As String = "Flags"
Private Const SETTINGS_NODE As String = "Settings"
Private Const COLLECTIONS_NODE As String = "Collections"
Private Const COLLECTION_NODE As String = "Collection"
Private Const ITEM_NODE As String = "Item"
Private Const TRUE_SYMBOL As String = "TRUE"
Private Const FALSE_SYMBOL As String = "FALSE"

Private Type TState
    Workbook As Workbook
    XPathPrefix As String
End Type
Private This As TState

' ---
Public Property Get XPathPrefix() As String
    XPathPrefix = This.XPathPrefix
End Property

Public Property Let XPathPrefix(ByVal vNewValue As String)
    This.XPathPrefix = vNewValue
End Property

Public Property Get Workbook() As Workbook
    Set Workbook = This.Workbook
End Property

Public Property Set Workbook(ByVal vNewValue As Workbook)
    Set This.Workbook = vNewValue
End Property

' ---
Private Function ISettings_GetFlag(ByVal FlagName As String) As Boolean
    Dim XPath As String
    XPath = XPathPrefix & "/" & FLAGS_NODE & "/" & FlagName
    
    Dim Result As CustomXMLNode
    Set Result = GetCustomXMLPart.SelectSingleNode(XPath)
    
    If Not Result Is Nothing Then
        Select Case UCase$(Result.text)
            Case TRUE_SYMBOL:
                ISettings_GetFlag = True
            Case FALSE_SYMBOL:
                ISettings_GetFlag = False
            Case Else:
                Debug.Assert False
        End Select
    End If
End Function

Private Sub ISettings_SetFlag(ByVal FlagName As String, ByVal Value As Boolean)
    Dim XPath As String
    XPath = XPathPrefix & "/" & FLAGS_NODE & "/" & FlagName
    
    UpsertText GetCustomXMLPart, XPath, IIf(Value, TRUE_SYMBOL, FALSE_SYMBOL)
End Sub

Private Function ISettings_GetSetting(ByVal SettingName As String) As String
    Dim XPath As String
    XPath = XPathPrefix & "/" & SETTINGS_NODE & "/" & SettingName
    
    Dim Result As CustomXMLNode
    Set Result = GetCustomXMLPart.SelectSingleNode(XPath)
    
    If Not Result Is Nothing Then
        ISettings_GetSetting = Result.text
    End If
End Function

Private Sub ISettings_SetSetting(ByVal SettingName As String, ByVal Value As String)
    Dim XPath As String
    XPath = XPathPrefix & "/" & SETTINGS_NODE & "/" & SettingName
    
    UpsertText GetCustomXMLPart, XPath, Value
End Sub

Private Function ISettings_GetCollection(ByVal CollectionName As String) As Collection
    Set ISettings_GetCollection = New Collection
    
    Dim XPath As String
    XPath = XPathPrefix & "/" & COLLECTIONS_NODE & "/" & COLLECTION_NODE & "[@Name='" & CollectionName & "']/" & ITEM_NODE
    
    Dim Results As CustomXMLNodes
    Set Results = GetCustomXMLPart.SelectNodes(XPath)
    
    If Not Results Is Nothing Then
        Dim Item As CustomXMLNode
        For Each Item In Results
            ISettings_GetCollection.Add Item:=Item.text
        Next
    End If
End Function

Private Function ISettings_HasCollection(ByVal CollectionName As String) As Boolean
    ISettings_HasCollection = Not ISettings_GetCollection(CollectionName) Is Nothing
End Function

Private Sub ISettings_SetCollection(ByVal CollectionName As String, ByVal Collection As Collection)
    Dim XPath As String
    XPath = XPathPrefix & "/" & COLLECTIONS_NODE & "/" & COLLECTION_NODE & "[@Name='" & CollectionName & "']"
    
    UpsertCollection GetCustomXMLPart, XPath, Collection
End Sub

Private Sub ISettings_Reset()
    GetCustomXMLPart.Delete
End Sub

' ---
Public Function Create(ByVal Workbook As Workbook, ByVal XPathPrefix As String) As ISettings
    Dim Result As XMLSettings
    Set Result = New XMLSettings
    
    With Result
        Set .Workbook = Workbook
        .XPathPrefix = XPathPrefix
        .Load
    End With
    
    Set Create = Result
End Function

Public Sub Load()
End Sub


Private Property Get RootNode() As String
    RootNode = Mid$(This.XPathPrefix, 2, InStr(2, This.XPathPrefix, "/") - 2)
End Property

Public Sub DebugPrint()
    Debug.Print GetCustomXMLPart.XML
End Sub

Private Function GetCustomXMLPart() As CustomXMLPart
    Dim Result As CustomXMLNode
    Dim CustomXMLPart As CustomXMLPart
    For Each CustomXMLPart In Workbook.CustomXMLParts
        Set Result = CustomXMLPart.SelectSingleNode("/" & RootNode)
        If Not Result Is Nothing Then
            Set GetCustomXMLPart = CustomXMLPart
            Exit Function
        End If
    Next CustomXMLPart

    Workbook.CustomXMLParts.Add "<" & RootNode & "/>"
    Set GetCustomXMLPart = Workbook.CustomXMLParts.Item(Workbook.CustomXMLParts.Count)
End Function
