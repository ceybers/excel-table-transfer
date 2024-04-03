VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatchQualityToListView"
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

Private Sub Load(ByVal KeyColumnComparerResults As Collection, ByVal ListView As MSComctlLib.ListView)
    ListView.ListItems.Clear
    
    Dim Item As Variant
    For Each Item In KeyColumnComparerResults
        ListView.ListItems.Add Text:=CStr(Item)
    Next Item
    
    ListView.ColumnHeaders.Item(1).Text = COL_HDR_RESULTS & " (" & KeyColumnComparerResults.Count & ")"
End Sub

Private Sub IValueConverter_Initialize(ByVal Target As Object)
    Dim ListView As MSComctlLib.ListView
    Set ListView = Target
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:=COL_HDR_RESULTS, Width:=ListView.Width - 16
        .BorderStyle = ccNone
        .FullRowSelect = True
        .Gridlines = False
        .HideSelection = False
        .LabelEdit = lvwManual
        .View = lvwReport
    End With
End Sub



