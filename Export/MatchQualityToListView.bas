Attribute VB_Name = "MatchQualityToListView"
'@Folder("MVVM.ValueConverters")
Option Explicit

Public Sub Initialize(ByVal ListView As MScomctllib.ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwList
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .BorderStyle = ccNone
    End With
End Sub

Public Sub Load(ByVal ListView As MScomctllib.ListView, ByVal Collection As Collection)
    Debug.Assert Not Collection Is Nothing
    
    ListView.ListItems.Clear
    
    Dim Item As Variant
    For Each Item In Collection
        ListView.ListItems.Add Text:=CStr(Item)
    Next Item
End Sub
