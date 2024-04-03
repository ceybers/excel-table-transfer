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

Private Sub Load(ByVal KeyColumn As KeyColumn, ByVal ListView As MSComctlLib.ListView)
    ListView.ListItems.Clear
    
    Dim Total As Long
    Total = KeyColumn.Range.Count
    
    With KeyColumn
        AddItem ListView, COL_QUALITY_DISTINCT, .Keys.Count, IIf(.Keys.Count = Total, MSO_ACCEPT, IIf(.Keys.Count = 0, MSO_ERROR, MSO_WARNING))
        AddItem ListView, COL_QUALITY_UNIQUE, .UniqueKeys.Count, IIf(.UniqueKeys.Count = Total, MSO_ACCEPT, IIf(.Keys.Count = 0, MSO_ERROR, MSO_WARNING))
        AddItem ListView, COL_QUALITY_NONTEXT, .NonTextCount, IIf(.NonTextCount > 0, MSO_WARNING, MSO_ACCEPT)
        AddItem ListView, COL_QUALITY_BLANKS, .BlankCount, IIf(.BlankCount > 0, MSO_WARNING, MSO_ACCEPT)
        AddItem ListView, COL_QUALITY_ERRORS, .ErrorCount, IIf(.ErrorCount > 0, MSO_WARNING, MSO_ACCEPT)
        AddItem ListView, COL_QUALITY_TOTAL, Total, MSO_AUTO_SUM
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

Private Sub IValueConverter_Initialize(ByVal Target As Object)
    Dim ListView As MSComctlLib.ListView
    Set ListView = Target
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:=COL_HDR_MEASURE, Width:=64
        .ColumnHeaders.Add Text:=COL_HDR_VALUE, Width:=48
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
