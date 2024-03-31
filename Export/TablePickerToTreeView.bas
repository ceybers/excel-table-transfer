Attribute VB_Name = "TablePickerToTreeView"
'@Folder("MVVM2.ValueConverters")
Option Explicit

Private Const MSO_APP As String = "Excel"
Private Const MSO_LO As String = "lo"
Private Const MSO_LO_SEL As String = "activeLo"

Public Sub Initialize(ByVal TreeView As TreeView)
    With TreeView
        .Nodes.Clear ' TODO change to remove index 0 because it is faster
        .Indentation = 16
        .Style = tvwTreelinesPictureText
        Set .ImageList = StandardImageList.GetMSOImageList
    End With
End Sub

Public Sub Load(ByVal TreeView As TreeView, ByVal ViewModel As TablePickerViewModel)
    Dim SelectedKey As String
    If Not TreeView.SelectedItem Is Nothing Then
        SelectedKey = TreeView.SelectedItem.Key
    End If
    
    TreeView.Nodes.Clear
    
    'If TransferDeltasViewModel Is Nothing Then Exit Sub
    Debug.Assert Not ViewModel Is Nothing
    
    Dim NodeCache As Object
    Set NodeCache = CreateObject("Scripting.Dictionary")
    
    Dim Node As AvailableTableNode
    For Each Node In ViewModel.Items
        AddNode TreeView, Node, NodeCache
    Next Node
    
    If SelectedKey <> vbNullString Then
        If NodeCache.Exists(SelectedKey) Then
            Dim TreeViewNode As Node
            Set TreeViewNode = NodeCache.Item(SelectedKey)
            TreeViewNode.Selected = True
        End If
    End If
End Sub

Private Sub AddNode(ByVal TreeView As TreeView, ByVal AvailableTable As AvailableTableNode, ByVal NodeCache As Object)
    Dim Node As Node
    If AvailableTable.NodeType = ttApplication Then
        Set Node = TreeView.Nodes.Add(Key:=AvailableTable.Key, Text:=AvailableTable.Caption)
        Node.Expanded = True
        Node.image = MSO_APP
    Else
        Set Node = TreeView.Nodes.Add( _
            Relative:=NodeCache.Item(AvailableTable.ParentKey), Relationship:=tvwChild, _
            Key:=AvailableTable.Key, Text:=AvailableTable.Caption)
        Node.image = MSO_LO
        If AvailableTable.IsSelected Then
            Node.image = MSO_LO_SEL
        End If
    End If
    NodeCache.Add Key:=Node.Key, Item:=Node
End Sub
