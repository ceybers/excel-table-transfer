VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AvailableTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Model2.AvailableTables"
Option Explicit

Private Type TState
    AllTables As Collection
    Nodes As Collection
    RootNodeKey As String
End Type
Private This As TState

Public Property Get Nodes() As Collection
    Set Nodes = This.Nodes
End Property

Public Sub Load()
    Set This.AllTables = ListObjectHelpers.GetAllTablesInApplication
    Set This.Nodes = New Collection
    GenerateNodes
End Sub

Private Sub GenerateNodes()
    GenerateRootNode
    GenerateTableNodes
End Sub

Private Sub GenerateRootNode()
    Dim Node As AvailableTableNode
    Set Node = New AvailableTableNode
    With Node
        .Caption = "Tables"
        .NodeType = ttApplication
        .Key = GetKey
    End With
    This.Nodes.Add Item:=Node, Key:=Node.Key
    This.RootNodeKey = Node.Key
End Sub

Private Sub GenerateTableNodes()
    Dim ListObject As ListObject
    For Each ListObject In This.AllTables
        Dim Node As AvailableTableNode
        Set Node = New AvailableTableNode
        With Node
            .Caption = ListObject.Name & " (" & ListObject.Parent.Parent.Name & ")"
            .NodeType = ttListObject
            .Key = GetKey
            .ParentKey = This.RootNodeKey
            Set .Object = ListObject
        End With
        This.Nodes.Add Item:=Node, Key:=Node.Key
    Next ListObject
End Sub

Private Function GetKey(Optional ByVal Reset As Boolean) As String
    Static LastKey As Long
    If Reset Then
        LastKey = 1
    End If
    
    GetKey = "KEY#" & Format$(LastKey & "000")
    
    LastKey = LastKey + 1
End Function
