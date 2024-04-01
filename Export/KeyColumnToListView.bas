Attribute VB_Name = "KeyColumnToListView"
'@Folder("MVVM.ValueConverters")
Option Explicit

Public Sub Initialize(ByVal ListView As MSComctlLib.ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Measurement", Width:=64
        .ColumnHeaders.Add Text:="Value", Width:=48
        .ColumnHeaders.Item(2).Alignment = lvwColumnRight
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = False
        .LabelEdit = lvwManual
        .BorderStyle = ccNone
        .HideSelection = False
        Set .SmallIcons = StandardImageList.GetMSOImageList(16)
    End With
End Sub

Public Sub Load(ByVal ListView As MSComctlLib.ListView, ByVal KeyColumn As KeyColumn)
    Debug.Assert Not KeyColumn Is Nothing
    
    ListView.ListItems.Clear
    
    Dim Total As Long
    Total = KeyColumn.Range.Count
    
    With KeyColumn
        AddItem ListView, "Distinct", .Keys.Count, IIf(.Keys.Count = Total, MSO_ACCEPT, IIf(.Keys.Count = 0, MSO_ERROR, MSO_WARNING))
        AddItem ListView, "Unique", .UniqueKeys.Count, IIf(.UniqueKeys.Count = Total, MSO_ACCEPT, IIf(.Keys.Count = 0, MSO_ERROR, MSO_WARNING))
        AddItem ListView, "Non-text", .NonTextCount, IIf(.NonTextCount > 0, MSO_WARNING, MSO_ACCEPT)
        AddItem ListView, "Blanks", .BlankCount, IIf(.BlankCount > 0, MSO_WARNING, MSO_ACCEPT)
        AddItem ListView, "Errors", .ErrorCount, IIf(.ErrorCount > 0, MSO_WARNING, MSO_ACCEPT)
        AddItem ListView, "Total", Total, MSO_AUTO_SUM
    End With
    
    With ListView.ListItems
        .Item(.Count).Bold = True
        .Item(.Count).ListSubItems.Item(1).Bold = True
    End With
End Sub

Private Sub AddItem(ByVal ListView As MSComctlLib.ListView, ByVal Text As String, ByVal Value As Long, _
    ByVal Icon As String)
    Dim ListItem As MSComctlLib.ListItem
    Set ListItem = ListView.ListItems.Add(Text:=Text, SmallIcon:=Icon)
    ListItem.ListSubItems.Add Text:=CStr(Value)
End Sub

