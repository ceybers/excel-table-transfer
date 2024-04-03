Attribute VB_Name = "TablePickerToTreeView"
'@Folder "MVVM.ValueConverters"
Option Explicit

Public Sub Initialize(ByVal TreeView As MSComctlLib.TreeView)
    With TreeView
        .Nodes.Clear ' TODO change to remove index 0 because it is faster
        .Indentation = 16
        .Style = tvwTreelinesPictureText
        Set .ImageList = StandardImageList.GetMSOImageList
    End With
End Sub

Public Sub Load(ByVal TreeView As MSComctlLib.TreeView, ByVal ViewModel As TablePickerViewModel)
    Debug.Assert Not ViewModel Is Nothing
    
    ' Preserve current selected node by its key
    Dim SelectedKey As String
    If Not TreeView.SelectedItem Is Nothing Then
        SelectedKey = TreeView.SelectedItem.Key
    End If
    
    TreeView.Nodes.Clear
    
    Dim NodeCache As Object
    Set NodeCache = CreateObject("Scripting.Dictionary")
    
    Dim Node As AvailableTableNode
    For Each Node In ViewModel.Items
        AddNode TreeView, Node, NodeCache
    Next Node

    ' Restore previously selected node by its key
    If SelectedKey <> vbNullString Then
        If NodeCache.Exists(SelectedKey) Then
            Dim TreeViewNode As Node
            Set TreeViewNode = NodeCache.Item(SelectedKey)
            TreeViewNode.Selected = True
        End If
    End If
    
    CheckNoTablesFound TreeView
End Sub

Private Sub AddNode(ByVal TreeView As MSComctlLib.TreeView, ByVal AvailableTable As AvailableTableNode, ByVal NodeCache As Object)
    Dim Node As MSComctlLib.Node
    If AvailableTable.NodeType = ttApplication Then
        Set Node = TreeView.Nodes.Add(Key:=AvailableTable.Key, Text:=AvailableTable.Caption)
        Node.Expanded = True
    Else
        Set Node = TreeView.Nodes.Add( _
            Relative:=NodeCache.Item(AvailableTable.ParentKey), relationship:=tvwChild, _
            Key:=AvailableTable.Key, Text:=AvailableTable.Caption)
    End If
    
    UpdateNodeIcon AvailableTable, Node
    NodeCache.Add Key:=Node.Key, Item:=Node
End Sub

Private Sub UpdateNodeIcon(ByVal AvailableTable As AvailableTableNode, ByVal Node As Node)
    If AvailableTable.NodeType = ttApplication Then
        Node.image = MSO_EXCEL_APP
        Exit Sub
    End If
    
    With AvailableTable
        Select Case True
            Case .IsSelected
                Node.image = MSO_TABLE_TICK
                Node.Bold = True
            Case .IsProtected
                Node.image = MSO_SHEET_PROTECTED
            Case .IsHidden
                Node.image = MSO_SHEET_HIDDEN
            Case Else
                Node.image = MSO_TABLE
        End Select
    End With
End Sub

Private Sub CheckNoTablesFound(ByVal TreeView As MSComctlLib.TreeView)
    'NO_TABLES_FOUND
    If TreeView.Nodes.Count = 1 Then
        Dim Node As MSComctlLib.Node
        Set Node = TreeView.Nodes.Add(Relative:=TreeView.Nodes.Item(1), relationship:=tvwChild, _
            Key:=vbNullString, Text:=NO_TABLES_FOUND)
        Node.ForeColor = StringConstants.COLOR_NO_TABLES_AVAILABLE
        Node.image = MSO_NOT_FOUND
    End If
End Sub


