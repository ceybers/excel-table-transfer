VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablePickerViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM2.ViewModels")
Option Explicit
Implements IViewModel

Private Type TState
    SelectedTable As ListObject
    SourceTable As ListObject
    DestinationTable As ListObject
    SourceNode As AvailableTableNode
    DestinationNode As AvailableTableNode
    BaseItems As AvailableTables
    Items As Collection
End Type

Private This As TState

Public Property Get Foobar() As String
    Foobar = "foobar"
End Property

Public Property Get SelectedTable() As ListObject
    Set SelectedTable = This.SelectedTable
End Property

Public Property Set SelectedTable(ByVal RHS As ListObject)
    Set This.SelectedTable = RHS
End Property

Public Property Get SourceTable() As ListObject
    Set SourceTable = This.SourceTable
End Property

Public Property Set SourceTable(ByVal RHS As ListObject)
    Set This.SourceTable = RHS
End Property

Public Property Get DestinationTable() As ListObject
    Set DestinationTable = This.DestinationTable
End Property

Public Property Set DestinationTable(ByVal RHS As ListObject)
    Set This.DestinationTable = RHS
End Property

Public Property Get Items() As Collection
    Set Items = This.Items
End Property

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = True
End Function

Public Sub Load()
    Set This.BaseItems = New AvailableTables
    Set This.Items = New Collection
    This.BaseItems.Load
    UpdateNodes
End Sub

Private Sub UpdateNodes()
    CollectionClear This.Items
    
    Dim Node As AvailableTableNode
    For Each Node In This.BaseItems.Nodes
        This.Items.Add Node
    Next Node
End Sub

Public Function TryClick(ByVal Key As String) As Boolean
    Set This.SelectedTable = Nothing
    Dim Node As AvailableTableNode
    For Each Node In This.Items
        If Node.Key = Key Then
            If Node.NodeType = ttListObject Then
                Set This.SelectedTable = Node.Object
                TryClick = True
                Exit Function
            End If
        End If
    Next Node
End Function

Public Function CanSelect() As Boolean
    CanSelect = Not This.SelectedTable Is Nothing
End Function

Public Function CanNext() As Boolean
    If This.SourceTable Is Nothing Then Exit Function
    If This.DestinationTable Is Nothing Then Exit Function
    CanNext = True
End Function

Public Sub DoSelect(ByVal Direction As TransferDirection)
    If Direction = tdSource Then
        If Not This.SourceNode Is Nothing Then
            This.SourceNode.IsSelected = False
        End If
        Set This.SourceTable = This.SelectedTable
        Set This.SourceNode = GetNodeByObject(This.SelectedTable)
        This.SourceNode.IsSelected = True
    ElseIf Direction = tdDestination Then
        If Not This.DestinationNode Is Nothing Then
            This.DestinationNode.IsSelected = False
        End If
        Set This.DestinationTable = This.SelectedTable
        Set This.DestinationNode = GetNodeByObject(This.SelectedTable)
        This.DestinationNode.IsSelected = True
    End If
End Sub

Private Function GetNodeByObject(ByVal Object As Object) As AvailableTableNode
    Dim Item As AvailableTableNode
    For Each Item In This.Items
        If Item.Object Is Object Then
            Set GetNodeByObject = Item
            Exit Function
        End If
    Next Item
End Function
