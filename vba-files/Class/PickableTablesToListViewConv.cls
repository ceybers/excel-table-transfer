VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PickableTablesToListViewConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.TablePicker.Converters"
Option Explicit
Implements IValueConverter

Private Sub IValueConverter_Convert(ByVal Source As Object, ByVal Target As Object)
    ConvertListItems Source, Target
End Sub

Private Sub IValueConverter_ConvertBack(ByVal PropertyBinding As IPropertyBinding)
    Dim ListView As ListView
    Set ListView = PropertyBinding.Target
    
    Dim VM As TablePickerViewModel
    Set VM = PropertyBinding.Source
    Dim PickableTables As PickableTables
    Set PickableTables = CallByName(VM, PropertyBinding.SourcePropertyPath, VbGet)
    
    Dim ListItem As ListItem
    For Each ListItem In ListView.ListItems
        VM.TrySetColumnChecked PropertyBinding.SourcePropertyPath, ListItem.Key, ListItem.Checked
        VM.TrySetColumnSelected PropertyBinding.SourcePropertyPath, ListItem.Key, ListItem.Selected
    Next
End Sub

Private Sub ConvertSelectedItem(ByVal SelectedItemKey As String, ByVal ListView As ListView)
    Stop
End Sub

Private Sub ConvertListItems(ByVal PickableTables As PickableTables, ByVal ListView As ListView)
    Dim PickableTable As PickableTable
    For Each PickableTable In PickableTables.Item
        LoadListItem PickableTable, ListView
    Next PickableTable
End Sub

Private Sub LoadListItem(ByVal PickableTable As PickableTable, ByVal ListView As ListView)
    Dim ListItem As ListItem

    For Each ListItem In ListView.ListItems
        If ListItem.Key = PickableTable.Key Then
            'ListItem.Text = ColumnProperty.Name
            ListItem.ListSubItems.Clear
            Exit For
        End If
    Next ListItem
    
    If ListItem Is Nothing Then
        Set ListItem = ListView.ListItems.Add(Key:=PickableTable.Key, Text:=PickableTable.Name)
    End If
    
    ListItem.ListSubItems.Add Text:=PickableTable.WorksheetName
    ListItem.ListSubItems.Add Text:=PickableTable.WorkbookName
    ListItem.ListSubItems.Add Text:="Local"
    ListItem.ListSubItems.Add Text:="No"
    ListItem.ListSubItems.Add Text:="No"
    
    ListItem.ForeColor = IIf(PickableTable.Enabled, vbBlack, vbGrayText)
    ListItem.Bold = PickableTable.Checked
    ListItem.Checked = PickableTable.Checked
    ListItem.Selected = PickableTable.Selected
End Sub

Public Sub InitializeListView(ByVal ListView As ListView)
    Debug.Assert Not ListView Is Nothing
     
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add Text:="Table", Width:=48
        .ColumnHeaders.Add Text:="Worksheet", Width:=48
        .ColumnHeaders.Add Text:="Workbook", Width:=48
        .ColumnHeaders.Add Text:="Location", Width:=32
        .ColumnHeaders.Add Text:="Protected", Width:=24
        .ColumnHeaders.Add Text:="Read-only", Width:=24
        
        .Appearance = cc3D
        .BorderStyle = ccNone
        .CheckBoxes = True
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
    End With
End Sub

