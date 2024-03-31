Attribute VB_Name = "KeyColumnsToListView"
'@Folder("MVVM2.ValueConverters")
Option Explicit

Private Const MSO_ITEM As String = "lblSelCol"
Private Const MSO_SELECTED As String = "lblKey2"

Public Sub Initialize(ByVal ListView As ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Column Name"
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = False
        .LabelEdit = lvwManual
        .BorderStyle = ccNone
        .HideSelection = False
        Set .SmallIcons = StandardImageList.GetMSOImageList(16)
    End With
End Sub

Public Sub Load(ByVal ListView As ListView, ByVal KeyColumns As KeyColumns)
    Debug.Assert Not KeyColumns Is Nothing
    
    ListView.ListItems.Clear
    
    Dim KeyColumn As KeyColumn
    For Each KeyColumn In KeyColumns.KeyColumns
        AddItem ListView, KeyColumn
    Next KeyColumn
    
    If Not KeyColumns.Selected Is Nothing Then
        With ListView.ListItems.Item(KeyColumns.Selected.Name)
            .SmallIcon = MSO_SELECTED
            .Bold = True
            .Selected = True
        End With
    End If
End Sub

Private Sub AddItem(ByVal ListView As ListView, ByVal KeyColumn As KeyColumn)
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Key:=KeyColumn.Name, Text:=KeyColumn.Name)
    ListItem.SmallIcon = MSO_ITEM
End Sub


