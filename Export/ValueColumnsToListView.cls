VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueColumnsToListView"
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
    SelectItem PropertyBinding.Target, PropertyBinding.Source, PropertyBinding.SourcePropertyPath
End Sub

Private Sub SelectItem(ByVal ListView As MSComctlLib.ListView, _
    ByVal VM As ValueMapperViewModel, _
    ByVal SourcePropertyPath As String)
    Dim ValueColumns As ValueColumns
    Set ValueColumns = CallByName(VM, SourcePropertyPath, VbGet)
    
    ValueColumns.TrySelect ListView.SelectedItem.Key
End Sub

Private Sub Load(ByVal ValueColumns As ValueColumns, ByVal ListView As MSComctlLib.ListView)
    If ListView.ListItems.Count <> ValueColumns.ValueColumns.Count Then
        LoadItems ValueColumns, ListView
    End If
        
    UpdateItems ValueColumns, ListView
End Sub

Private Sub LoadItems(ByVal ValueColumns As ValueColumns, ByVal ListView As MSComctlLib.ListView)
    ListView.ListItems.Clear
    
    Dim ValueColumn As ValueColumn
    For Each ValueColumn In ValueColumns.ValueColumns
        AddItem ListView, ValueColumn
    Next ValueColumn
End Sub

Private Sub UpdateItems(ByVal ValueColumns As ValueColumns, ByVal ListView As MSComctlLib.ListView)
    Dim ListItem As ListItem
    For Each ListItem In ListView.ListItems
        UpdateListItem ListItem, ValueColumns
    Next ListItem
    
    If Not ListView.SelectedItem Is Nothing Then
        If ValueColumns.ValueColumns.Item(ListView.SelectedItem.Key).IsSelectable = False Then
            Set ListView.SelectedItem = Nothing
        End If
    End If
    
    If Not ValueColumns.Selected Is Nothing Then
        ListView.ListItems.Item(ValueColumns.Selected.Name).Selected = True
    End If
End Sub

Private Sub AddItem(ByVal ListView As MSComctlLib.ListView, ByVal ValueColumn As ValueColumn)
    ListView.ListItems.Add Key:=ValueColumn.Name, Text:=ValueColumn.Name
End Sub

Private Sub UpdateListItem(ByVal ListItem As MSComctlLib.ListItem, ByVal ValueColumns As ValueColumns)
    Dim ValueColumn As ValueColumn
    Set ValueColumn = ValueColumns.ValueColumns.Item(ListItem.Key)
    
    If ValueColumn.IsSelectable = False Then
        ListItem.ForeColor = COLOR_DISABLED
    End If
    
    ListItem.Bold = False
            
    Select Case True
        '@Ignore UnassignedVariableUsage
        Case ValueColumn.IsKeyColumn
            ListItem.SmallIcon = MSO_KEY
        Case ValueColumn.IsMapped
            ListItem.SmallIcon = MSO_LINK
            ListItem.Bold = True
        Case ValueColumn.IsFormula
            ListItem.SmallIcon = MSO_FORMULA
        Case ValueColumn.IsEmpty
            ListItem.SmallIcon = MSO_EMPTY
        Case ValueColumn.HasErrors
            ListItem.SmallIcon = MSO_ERROR
        Case ValueColumn.HasNumbers
            If ValueColumn.DataType = vbCurrency Then
                ListItem.SmallIcon = MSO_TYPE_CURRENCY
            ElseIf ValueColumn.DataType = vbDate Then
                ListItem.SmallIcon = MSO_TYPE_DATE
            Else
                ListItem.SmallIcon = MSO_TYPE_LONG
            End If
        Case Else
            ListItem.SmallIcon = MSO_TYPE_STRING
    End Select
End Sub

Private Sub IValueConverter_Initialize(ByVal Target As Object)
    Dim ListView As MSComctlLib.ListView
    Set ListView = Target
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:=COL_HDR_COLUMN_NAME, Width:=ListView.Width - 16
        Set .SmallIcons = StandardImageList.GetMSOImageList(16)
        .BorderStyle = ccNone
        .FullRowSelect = True
        .Gridlines = False
        .HideColumnHeaders = True
        .HideSelection = False
        .LabelEdit = lvwManual
        .View = lvwReport
    End With
End Sub
