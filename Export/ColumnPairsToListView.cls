VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnPairsToListView"
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
    SelectItem PropertyBinding.Target, PropertyBinding.Source, PropertyBinding.SourcePropertyPath
End Sub

Private Sub SelectItem(ByVal ListView As MScomctllib.ListView, ByVal VM As ValueMapperViewModel, _
    ByVal SourcePropertyPath As String)
    Dim ValueColumns As ValueColumns
    Set ValueColumns = CallByName(VM, SourcePropertyPath, VbGet)
    
    ValueColumns.TrySelect ListView.SelectedItem.Key
End Sub

Private Sub Load(ByVal ColumnPairs As ColumnPairs, ByVal ListView As MScomctllib.ListView)
    ListView.ListItems.Clear
    
    Dim ColumnPair As ColumnPair
    For Each ColumnPair In ColumnPairs
        AddItem ListView, ColumnPair
    Next ColumnPair
End Sub

Private Sub AddItem(ByVal ListView As MScomctllib.ListView, ByVal ColumnPair As ColumnPair)
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Text:=ColumnPair.Source)
    ListItem.ListSubItems.Add Text:=ColumnPair.Destination
    ListItem.SmallIcon = MSO_LINK
End Sub

Public Sub Initialize(ByVal ListView As MScomctllib.ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Source", Width:=(ListView.Width - 16) / 2 ' TODO Const column header text
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
