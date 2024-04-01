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

Private Sub IValueConverter_Convert(ByVal Source As Object, ByVal Target As Object) ' read
    Load Source, Target
End Sub

Private Sub IValueConverter_ConvertBack(ByVal PropertyBinding As IPropertyBinding) ' write
    Dim ListView As MScomctllib.ListView
    Set ListView = PropertyBinding.Target
    
    Dim VM As KeyMapperViewModel
    Set VM = PropertyBinding.Source
    
    Dim KeyColumns As KeyColumns
    Set KeyColumns = CallByName(VM, PropertyBinding.SourcePropertyPath, VbGet)
    
    KeyColumns.TrySelect ListView.SelectedItem.Key
End Sub

Public Sub Load(ByVal KeyColumns As KeyColumns, ByVal ListView As MScomctllib.ListView)
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
        .ColumnHeaders.Add Text:="Column Name", Width:=ListView.Width - 16
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = False
        .LabelEdit = lvwManual
        .BorderStyle = ccNone
        .HideSelection = False
        Set .SmallIcons = StandardImageList.GetMSOImageList(16)
    End With
End Sub
