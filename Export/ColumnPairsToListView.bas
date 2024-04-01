Attribute VB_Name = "ColumnPairsToListView"
'@Folder "MVVM.ValueConverters"
Option Explicit

Public Sub Initialize(ByVal ListView As MSComctlLib.ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Source", Width:=(ListView.Width - 16) / 2
        .ColumnHeaders.Add Text:="Destination", Width:=(ListView.Width - 16) / 2
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = False
        .LabelEdit = lvwManual
        .BorderStyle = ccNone
        .HideSelection = False
        Set .SmallIcons = StandardImageList.GetMSOImageList(16)
    End With
End Sub

Public Sub Load(ByVal ListView As MSComctlLib.ListView, ByVal ColumnPairs As ColumnPairs)
    Debug.Assert Not ColumnPairs Is Nothing
    
    ListView.ListItems.Clear
    
    Dim ColumnPair As ColumnPair
    For Each ColumnPair In ColumnPairs
        AddItem ListView, ColumnPair
    Next ColumnPair
End Sub

Private Sub AddItem(ByVal ListView As MSComctlLib.ListView, ByVal ColumnPair As ColumnPair)
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Text:=ColumnPair.Source)
    ListItem.ListSubItems.Add Text:=ColumnPair.Destination
    ListItem.SmallIcon = MSO_LINK
End Sub

