Attribute VB_Name = "KeyColumnsToListView"
'@Folder "MVVM.ValueConverters"
Option Explicit

Public Sub Initialize(ByVal ListView As MScomctllib.ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Column Name", Width:=ListView.Width - 16
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = False
        .LabelEdit = lvwManual
        .BorderStyle = ccNone
        .HideSelection = False
        Set .SmallIcons = StandardImageList.GetMSOImageList(16)
    End With
End Sub

Public Sub Load(ByVal ListView As MScomctllib.ListView, ByVal KeyColumns As KeyColumns)
    Debug.Assert Not KeyColumns Is Nothing
    
    ListView.ListItems.Clear
    
    Dim KeyColumn As KeyColumn
    For Each KeyColumn In KeyColumns.KeyColumns
        AddItem ListView, KeyColumn
    Next KeyColumn
    
    If Not KeyColumns.Selected Is Nothing Then
        With ListView.ListItems.Item(KeyColumns.Selected.Name)
            .SmallIcon = MSO_KEY
            .Bold = True
            .Selected = True
        End With
    End If
End Sub

Private Sub AddItem(ByVal ListView As MScomctllib.ListView, ByVal KeyColumn As KeyColumn)
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Key:=KeyColumn.Name, Text:=KeyColumn.Name)
    ListItem.SmallIcon = MSO_CELL
End Sub

