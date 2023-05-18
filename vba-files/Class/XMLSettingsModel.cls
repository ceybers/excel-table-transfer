VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMLSettingsModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "PersistentStorage.XMLSettings"
Option Explicit
Implements ISettingsModel

Private Const WORKBOOK_NODE As String = "Workbook"
Private Const TABLES_NODE As String = "Tables"
Private Const TABLE_NODE As String = "Table"

Private Type TState
    RootNode As String
    Workbook As Workbook
    WorkbookSettings As ISettings
    TableSettings As Collection ' of type SomeSettings
End Type
Private This As TState

Public Property Get RootNode() As String
    RootNode = This.RootNode
End Property

Public Property Let RootNode(ByVal vNewValue As String)
    This.RootNode = vNewValue
End Property

Public Property Get Workbook() As Workbook
    Set Workbook = This.Workbook
End Property

Public Property Get CustomXMLPart() As CustomXMLPart
    Set CustomXMLPart = GetCustomXMLPart()
End Property

Private Sub Class_Initialize()
    Set This.TableSettings = New Collection
End Sub

'@Description "Creates a new XMLSettingsModel for the given Workbook and RootNode."
Public Function Create(ByVal Workbook As Workbook, ByVal RootNode As String) As XMLSettingsModel
Attribute Create.VB_Description = "Creates a new XMLSettingsModel for the given Workbook and RootNode."
    Set Create = New XMLSettingsModel
    Create.RootNode = RootNode
    Create.Load Workbook
End Function

'@Description "Loads the Workbook settings and any Table settings from the Workbook's CustomXMLPart."
Public Sub Load(ByVal Workbook As Workbook)
Attribute Load.VB_Description = "Loads the Workbook settings and any Table settings from the Workbook's CustomXMLPart."
    Set This.Workbook = Workbook
    
    AddWorkbookSettings
        
    Dim TableName As Variant
    For Each TableName In GetTableNames
        AddTableSettings TableName
    Next TableName
End Sub

Private Sub AddWorkbookSettings()
    Dim WorkbookSettings As XMLSettings
    Set WorkbookSettings = New XMLSettings
    
    WorkbookSettings.XPathPrefix = "/" & RootNode & "/" & WORKBOOK_NODE
    WorkbookSettings.SetParent Me
    
    Set This.WorkbookSettings = WorkbookSettings
End Sub

Private Function AddTableSettings(ByVal TableName As String) As XMLSettings
    Dim TableSettings As XMLSettings
    Set TableSettings = New XMLSettings
    
    TableSettings.XPathPrefix = "/" & RootNode & "/" & TABLES_NODE & "/" & TABLE_NODE & "[@Name='" & TableName & "']"
    TableSettings.SetParent Me
    
    This.TableSettings.Add Item:=TableSettings, Key:=TableName
    
    ' If a table is not found, we can add it on demand
    Set AddTableSettings = TableSettings
End Function

Private Function GetTableNames() As Collection
    Dim CustomXMLPart As CustomXMLPart
    Set CustomXMLPart = GetCustomXMLPart
    
    Dim Result As Collection
    Set Result = New Collection
    
    Dim XPath As String
    XPath = "/" & RootNode & "/" & TABLES_NODE & " / " & TABLE_NODE & "[@Name]/@Name"
    
    Dim TableNodes As CustomXMLNodes
    Set TableNodes = CustomXMLPart.SelectNodes(XPath)
    
    Dim TableNode As CustomXMLNode
    For Each TableNode In TableNodes
        Result.Add Item:=TableNode.Text
    Next TableNode
    
    Set GetTableNames = Result
End Function

'@Description "Gets the CustomXMLPart in this Workbook that has the same root node as our settings. If it doesn't exist, we create an empty one."
Private Function GetCustomXMLPart() As CustomXMLPart
Attribute GetCustomXMLPart.VB_Description = "Gets the CustomXMLPart in this Workbook that has the same root node as our settings. If it doesn't exist, we create an empty one."
    Dim Result As CustomXMLNode
    Dim CustomXMLPart As CustomXMLPart
    For Each CustomXMLPart In Workbook.CustomXMLParts
        Set Result = CustomXMLPart.SelectSingleNode("/" & RootNode)
        If Not Result Is Nothing Then
            Set GetCustomXMLPart = CustomXMLPart
            Exit Function
        End If
    Next CustomXMLPart
    
    AddEmptyCustomXMLPart
    Set GetCustomXMLPart = Workbook.CustomXMLParts.Item(Workbook.CustomXMLParts.Count)
End Function

'@Description "Removes any existing CustomXMLPart, and replaces it with an empty root node."
Public Sub Reset()
Attribute Reset.VB_Description = "Removes any existing CustomXMLPart, and replaces it with an empty root node."
    ' If there was no existing CustomXMLPart, we would have created a new empty one in GetCustomXMLPart().
    ' It is easier to just do this, delete it, then create an empty one again.
    GetCustomXMLPart.Delete
    
    AddEmptyCustomXMLPart
End Sub

'@Description "Removes the CustomXMLPart from the Workbook."
Public Sub Delete()
Attribute Delete.VB_Description = "Removes the CustomXMLPart from the Workbook."
    If Not CustomXMLPart Is Nothing Then
        GetCustomXMLPart().Delete
    End If
End Sub

Private Sub AddEmptyCustomXMLPart()
    Workbook.CustomXMLParts.Add "<" & RootNode & "/>"
End Sub

'@Description "Debug.Print's the CustomXMLPart's XML as a string."
Public Sub DebugPrint()
Attribute DebugPrint.VB_Description = "Debug.Print's the CustomXMLPart's XML as a string."
    Debug.Print CustomXMLPart.XML
End Sub

Private Function ISettingsModel_Workbook() As ISettings
    Set ISettingsModel_Workbook = This.WorkbookSettings
End Function

Private Function ISettingsModel_Table(ByVal TableName As String) As ISettings
    On Error Resume Next
    Set ISettingsModel_Table = This.TableSettings.Item(TableName)
    On Error GoTo 0
    
    If ISettingsModel_Table Is Nothing Then
        Set ISettingsModel_Table = AddTableSettings(TableName)
    End If
End Function
