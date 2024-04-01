VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyColumnsToListView"
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

Private Sub SelectItem(ByVal ListView As MScomctllib.ListView, ByVal VM As KeyMapperViewModel, _
    ByVal SourcePropertyPath As String)
    Dim KeyColumns As KeyColumns
    Set KeyColumns = CallByName(VM, SourcePropertyPath, VbGet)
    
    KeyColumns.TrySelect ListView.SelectedItem.Key
End Sub

Private Sub Load(ByVal KeyColumns As KeyColumns, ByVal ListView As MScomctllib.ListView)
    ListView.ListItems.Clear
    
    Dim KeyColumn As KeyColumn
    For Each KeyColumn In KeyColumns.KeyColumns
        AddItem ListView, KeyColumn
    Next KeyColumn
    
    If Not KeyColumns.Selected Is Nothing Then
        With ListView.ListItems.Item(KeyColumns.Selected.Name)
            .SmallIcon = MSO_KEY
            .Bold = True
            .Selected = True
        End With
    End If
End Sub

Private Sub AddItem(ByVal ListView As MScomctllib.ListView, ByVal KeyColumn As KeyColumn)
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Key:=KeyColumn.Name, Text:=KeyColumn.Name)
    ListItem.SmallIcon = MSO_CELL
End Sub

Public Sub Initialize(ByVal ListView As MScomctllib.ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        '.ColumnHeaders.Add Text:="Column Name", Width:=ListView.Width - 16 ' TODO Header consts
        '.View = lvwReport
        .View = lvwList
        .FullRowSelect = True
        .Gridlines = False
        .LabelEdit = lvwManual
        .BorderStyle = ccNone
        .HideSelection = False
        Set .SmallIcons = StandardImageList.GetMSOImageList(16)
    End With
End Sub
