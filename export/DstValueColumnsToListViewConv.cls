VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DstValueColumnsToListViewConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.ValueMapper.Converters"
Option Explicit
Implements IValueConverter

Private Const NOT_MAPPED As String = " (not mapped)"

Private Sub IValueConverter_Convert(ByVal Source As Object, ByVal Target As Object)
    ConvertListItems Source, Target
End Sub

Private Sub IValueConverter_ConvertBack(ByVal PropertyBinding As IPropertyBinding)
    Dim ListView As ListView
    Set ListView = PropertyBinding.Target
    
    Dim VM As ValuePickerViewModel
    Set VM = PropertyBinding.Source
    
    VM.TryUpdateSelected ListView.SelectedItem
End Sub

Private Sub ConvertSelectedItem(ByVal SelectedItemKey As String, ByVal ListView As ListView)
    Stop
End Sub

Private Sub ConvertListItems(ByVal Dictionary As Scripting.Dictionary, ByVal ListView As ListView)
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
            ListItem.ListSubItems.Clear
            Exit For
        End If
    Next ListItem
    
    If ListItem Is Nothing Then
        Set ListItem = ListView.ListItems.Add(Key:=ValueColumn.Key, Text:=ValueColumn.Name)
    End If
    
    ListItem.Bold = ValueColumn.MappedTo <> vbNullString
    ListItem.SmallIcon = IIf(ValueColumn.IsKey, "Key", Empty)
    
    Dim ForeColor As Long
    ForeColor = IIf(ValueColumn.Enabled, vbBlack, vbGrayText)
    
    With ListItem.ListSubItems
        .Add Text:="(" & ValueColumn.ColumnLetter & ")"
        .Add Text:=IIf(ValueColumn.MappedTo = vbNullString, NOT_MAPPED, ValueColumn.MappedTo)
        .Item(2).ReportIcon = IIf(ValueColumn.MappedTo <> vbNullString, "MappedTrue", "MappedFalse")
        .Add Text:=ValueColumn.DataType
        .Add Text:=vbNullString 'IIf(ValueColumn.IsProtected, "Y", "N")
        .Item(4).ReportIcon = IIf(ValueColumn.IsProtected, "Protected", Empty)
        .Add Text:=vbNullString 'IIf(ValueColumn.IsHidden, "Y", "N")
        .Item(5).ReportIcon = IIf(ValueColumn.IsHidden, "Hidden", Empty)
        .Add Text:=IIf(ValueColumn.IsKey, "Y", "N")
        .Add Text:=vbNullString 'IIf(ValueColumn.IsStarred, "Y", "N")
        .Item(7).ReportIcon = IIf(ValueColumn.IsStarred, "StarTrue", "StarFalse")
    End With
    
    ListItem.ForeColor = ForeColor
    Dim ListSubItem As ListSubItem
    For Each ListSubItem In ListItem.ListSubItems
        ListSubItem.ForeColor = ForeColor
    Next ListSubItem
    
    With ListItem.ListSubItems.Item(1)
    If .Text = NOT_MAPPED Then
        .ForeColor = vbGrayText
    End If
    End With
    
    If ValueColumn.Selected Then
        ListItem.Selected = True
        Set ListView.DropHighlight = ListItem
    End If
End Sub

Public Sub InitializeListView(ByVal Context As IAppContext, ByVal ListView As ListView)
    Debug.Assert Not ListView Is Nothing
     
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add Text:="Name", Width:=72
        .ColumnHeaders.Add Text:="Address", Width:=24
        .ColumnHeaders.Add Text:="Mapped to", Width:=64
        .ColumnHeaders.Add Text:="Type", Width:=24
        .ColumnHeaders.Add Key:="IsProtected", Text:="", Width:=18
        .ColumnHeaders.Add Key:="IsHidden", Text:="", Width:=18
        .ColumnHeaders.Add Key:="IsKey", Text:="", Width:=0
        .ColumnHeaders.Add Key:="IsStarred", Text:="", Width:=18
        
        .Appearance = cc3D
        .BorderStyle = ccNone
        .CheckBoxes = False
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
        .MultiSelect = False
        
        Set .Icons = Context.ImageLists("32")
        Set .SmallIcons = Context.ImageLists("16")
    End With
End Sub
