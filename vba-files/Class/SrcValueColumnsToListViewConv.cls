VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SrcValueColumnsToListViewConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.ValueMapper.Converters"
Option Explicit
Implements IValueConverter

Private Sub IValueConverter_Convert(ByVal Source As Object, ByVal Target As Object)
    ConvertListItems Source, Target
End Sub

Private Sub IValueConverter_ConvertBack(ByVal PropertyBinding As IPropertyBinding)
    'DebugPrintPropertyBinding PropertyBinding

    Dim ListView As ListView
    Set ListView = PropertyBinding.Target
    
    Dim VM As ValuePickerViewModel
    Set VM = PropertyBinding.Source
    
    VM.TryUpdateSelected ListView.SelectedItem
    
    'Dim ListItem As ListItem
    'For Each ListItem In ListView.ListItems
        ' Order matters, checked must be last
        'VM.TrySetColumnSelected PropertyBinding.SourcePropertyPath, ListItem.Key, ListItem.Selected
        'VM.TrySetColumnChecked PropertyBinding.SourcePropertyPath, ListItem.Key, ListItem.Checked
    'Next ListItem
End Sub

Private Sub ConvertSelectedItem(ByVal SelectedItemKey As String, ByVal ListView As ListView)
    Stop
End Sub

Private Sub ConvertListItems(ByVal Dictionary As Scripting.Dictionary, ByVal ListView As ListView)
    'Set ListView.SelectedItem = Nothing
    Set ListView.DropHighlight = Nothing
    
    Dim Key As Variant
    For Each Key In Dictionary.Keys
        Dim ThisValueColumn As ValueColumn
        Set ThisValueColumn = Dictionary.Item(Key)
        LoadListItem ThisValueColumn, ListView
    Next Key
End Sub

Private Sub LoadListItem(ByVal ValueColumn As ValueColumn, ByVal ListView As ListView)
    Dim ListItem As ListItem

    For Each ListItem In ListView.ListItems
        If ListItem.Key = ValueColumn.Key Then
            'ListItem.Text = ColumnProperty.Name
            ListItem.ListSubItems.Clear
            Exit For
        End If
    Next ListItem
    
    If ListItem Is Nothing Then
        Set ListItem = ListView.ListItems.Add(Key:=ValueColumn.Key, Text:=ValueColumn.Name)
    End If
    
    Dim ForeColor As Long
    ForeColor = IIf(ValueColumn.Enabled, vbBlack, vbGrayText)
    
    With ListItem.ListSubItems
        .Add Text:="(" & ValueColumn.ColumnLetter & ")"
        .Add Text:=ValueColumn.DataType
        .Add Text:=IIf(ValueColumn.IsProtected, "Y", "N")
        .Add Text:=IIf(ValueColumn.IsHidden, "Y", "N")
        .Add Text:=IIf(ValueColumn.IsKey, "Y", "N")
        .Add Text:=IIf(ValueColumn.IsStarred, "Y", "N")
    End With
    
    ListItem.ForeColor = ForeColor
    Dim ListSubItem As ListSubItem
    For Each ListSubItem In ListItem.ListSubItems
        ListSubItem.ForeColor = ForeColor
    Next ListSubItem
    
    'Debug.Print ValueColumn.Name; " "; ValueColumn.Selected
    'ListItem.Bold = ValueColumn.Checked
    If ValueColumn.Selected Then
        ListItem.Selected = True
        'Set ListView.SelectedItem = ListItem
        Set ListView.DropHighlight = ListItem
    End If
End Sub

Public Sub InitializeListView(ByVal ListView As ListView)
    Debug.Assert Not ListView Is Nothing
     
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add Text:="Name", Width:=72
        .ColumnHeaders.Add Text:="Address", Width:=24
        .ColumnHeaders.Add Text:="Type", Width:=24
        .ColumnHeaders.Add Text:="IsProtected", Width:=32
        .ColumnHeaders.Add Text:="IsHidden", Width:=24
        .ColumnHeaders.Add Text:="IsKey", Width:=24
        .ColumnHeaders.Add Text:="IsStarred", Width:=24
        
        .Appearance = cc3D
        .BorderStyle = ccNone
        .CheckBoxes = False
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
        .MultiSelect = False
    End With
End Sub


