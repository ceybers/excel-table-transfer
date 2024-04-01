VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyColumnToListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("MVVM.ValueConverters")
Option Explicit

Implements IValueConverter

Private Sub IValueConverter_Convert(ByVal Source As Object, ByVal Target As Object)
    Load Source, Target
End Sub

Private Sub IValueConverter_ConvertBack(ByVal PropertyBinding As IPropertyBinding)
End Sub

Private Sub Load(ByVal KeyColumn As KeyColumn, ByVal ListView As MScomctllib.ListView)
    ListView.ListItems.Clear
    
    Dim Total As Long
    Total = KeyColumn.Range.Count
    
    With KeyColumn
        ' TODO Move list item text to Consts
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

Private Sub AddItem(ByVal ListView As MScomctllib.ListView, ByVal Text As String, ByVal Value As Long, _
    ByVal Icon As String)
    Dim ListItem As MScomctllib.ListItem
    Set ListItem = ListView.ListItems.Add(Text:=Text, SmallIcon:=Icon)
    ListItem.ListSubItems.Add Text:=CStr(Value)
End Sub

Public Sub Initialize(ByVal ListView As MScomctllib.ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Measurement", Width:=64 ' TODO Header Consts
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
