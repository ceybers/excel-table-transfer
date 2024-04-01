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

Private Sub IValueConverter_Convert(ByVal Source As Object, ByVal Target As Object) ' read
    Load Source, Target
End Sub

Private Sub IValueConverter_ConvertBack(ByVal PropertyBinding As IPropertyBinding) ' write
    Dim ListView As MScomctllib.ListView
    Set ListView = PropertyBinding.Target
    
    Dim VM As KeyMapperViewModel
    Set VM = PropertyBinding.Source
    
    Dim ValueColumns As ValueColumns
    Set ValueColumns = CallByName(VM, PropertyBinding.SourcePropertyPath, VbGet)
    
    ValueColumns.TrySelect ListView.SelectedItem.Key
End Sub

Public Sub Load(ByVal ColumnPairs As ColumnPairs, ByVal ListView As MScomctllib.ListView)
    Debug.Assert Not ColumnPairs Is Nothing
    
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