Attribute VB_Name = "KeyColumnToListView"
'@Folder "MVVM.ValueConverters"
Option Explicit

Public Sub UpdateListView(ByVal lv As ListView)
    With lv.ListItems.Item(lv.ListItems.Count)
        .Bold = True
        .ListSubItems.Item(1).Bold = True
    End With
End Sub

Private Sub AddListViewItem(ByVal lv As ListView, ByVal Caption As String, ByVal Value As Long, ByVal icon As String)
    With lv.ListItems.Add(Text:=Caption, icon:=icon, SmallIcon:=icon)
        .ListSubItems.Add Text:=Value
    End With
End Sub

Public Sub InitializeListView(ByVal lv As ListView4)
    SetListViewImageList lv
    SetListViewProperties lv
End Sub

Private Sub SetListViewProperties(ByVal lv As ListView)
    With lv
        .View = lvwReport
        .HideSelection = False
        .CheckBoxes = False
        .LabelEdit = lvwManual
        .Gridlines = True
        .BorderStyle = ccNone
    End With
    
    lv.ListItems.Clear
    lv.ColumnHeaders.Clear
    
    lv.ColumnHeaders.Add Text:="Description"
    lv.ColumnHeaders.Add Text:="Value"
    lv.ColumnHeaders.Item(2).Alignment = lvwColumnRight
    lv.ColumnHeaders.Item(2).Width = (72 / 2)
    lv.ColumnHeaders.Item(1).Width = lv.Width - (72 / 2) - 5
End Sub

Private Sub SetListViewImageList(ByVal lv As ListView)
    'Dim il As ImageList
    'If lv.Icons Is Nothing Then
    'If True Then
        'Set il = GetMSOImageList
        Set lv.Icons = GetMSOImageList(32)
        Set lv.SmallIcons = GetMSOImageList(16)
    'End If
End Sub

