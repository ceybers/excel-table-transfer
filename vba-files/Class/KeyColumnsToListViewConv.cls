VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyColumnsToListViewConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.KeyMapper.Converters"
Option Explicit
Implements IValueConverter

Private Sub IValueConverter_Convert(ByVal Source As Object, ByVal Target As Object)
    ConvertListItems Source, Target
End Sub

Private Sub IValueConverter_ConvertBack(ByVal PropertyBinding As IPropertyBinding)
    Dim ListView As ListView
    Set ListView = PropertyBinding.Target
    
    Dim VM As KeyPickerViewModel
    Set VM = PropertyBinding.Source
    
    Dim ListItem As ListItem
    For Each ListItem In ListView.ListItems
        VM.TrySetColumnChecked PropertyBinding.SourcePropertyPath, ListItem.Key, ListItem.Checked
    Next ListItem
    
    Dim SelListItem As ListItem
    Set SelListItem = ListView.SelectedItem
    If Not SelListItem Is Nothing Then
        VM.TrySetColumnChecked PropertyBinding.SourcePropertyPath, SelListItem.Key, True
    End If
End Sub

Private Sub ConvertSelectedItem(ByVal SelectedItemKey As String, ByVal ListView As ListView)
    Stop
End Sub

Private Sub ConvertListItems(ByVal Dictionary As Scripting.Dictionary, ByVal ListView As ListView)
    Dim Key As Variant
    For Each Key In Dictionary.Keys
        Dim KeyColumn As KeyColumn
        Set KeyColumn = Dictionary.Item(Key)
        LoadListItem KeyColumn, ListView
    Next Key
End Sub

Private Sub LoadListItem(ByVal KeyColumn As KeyColumn, ByVal ListView As ListView)
    Dim ListItem As ListItem

    For Each ListItem In ListView.ListItems
        If ListItem.Key = KeyColumn.Key Then
            'ListItem.Text = ColumnProperty.Name
            ListItem.ListSubItems.Clear
            Exit For
        End If
    Next ListItem
    
    If ListItem Is Nothing Then
        Set ListItem = ListView.ListItems.Add(Key:=KeyColumn.Key, Text:=KeyColumn.Name)
    End If
    
    Dim ForeColor As Long
    ForeColor = IIf(KeyColumn.Enabled, vbBlack, vbGrayText)
    
    With ListItem.ListSubItems
        .Add Text:=KeyColumn.DataType
        .Add Text:=IIf(KeyColumn.IsText, "Y", "N")
        .Add Text:=IIf(KeyColumn.IsUnique, "Y", "N")
    End With
    
    ListItem.ForeColor = ForeColor
    Dim ListSubItem As ListSubItem
    For Each ListSubItem In ListItem.ListSubItems
        ListSubItem.ForeColor = ForeColor
    Next ListSubItem
    
    ListItem.Bold = KeyColumn.Checked
    ListItem.Checked = KeyColumn.Checked
    'If KeyColumn.Checked Then
        'If Not ListView.DropHighlight Is Nothing Then
            'ListView.DropHighlight = -1
        'End If
        'Set ListView.DropHighlight = ListItem
    'End If
End Sub

Public Sub InitializeListView(ByVal ListView As ListView)
    Debug.Assert Not ListView Is Nothing
     
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add Text:="Name", Width:=72
        .ColumnHeaders.Add Text:="Type", Width:=48
        .ColumnHeaders.Add Text:="IsText", Width:=48
        .ColumnHeaders.Add Text:="IsUnique", Width:=32
        .ColumnHeaders.Add Text:="D", Width:=24
        .ColumnHeaders.Add Text:="E", Width:=24
        
        .Appearance = cc3D
        .BorderStyle = ccNone
        .CheckBoxes = True
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
    End With
End Sub

