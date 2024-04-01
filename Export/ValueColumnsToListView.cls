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
    Dim ListView As MScomctllib.ListView
    Set ListView = PropertyBinding.Target
    
    Dim VM As ValueMapperViewModel
    Set VM = PropertyBinding.Source
    
    Dim ValueColumns As ValueColumns
    Set ValueColumns = CallByName(VM, PropertyBinding.SourcePropertyPath, VbGet)
    
    ValueColumns.TrySelect ListView.SelectedItem.Key
End Sub


Public Sub Load(ByVal ValueColumns As ValueColumns, ByVal ListView As MScomctllib.ListView)
    Debug.Assert Not ValueColumns Is Nothing
    
    If ListView.ListItems.Count <> ValueColumns.ValueColumns.Count Then
        LoadItems ListView, ByVal ValueColumns
    End If
        
    UpdateItems ListView, ByVal ValueColumns
End Sub

Private Sub LoadItems(ByVal ListView As MScomctllib.ListView, ByVal ValueColumns As ValueColumns)
    ListView.ListItems.Clear
    Dim ValueColumn As ValueColumn
    For Each ValueColumn In ValueColumns.ValueColumns
        AddItem ListView, ValueColumn
    Next ValueColumn
End Sub

Private Sub UpdateItems(ByVal ListView As MScomctllib.ListView, ByVal ValueColumns As ValueColumns)
    Dim ListItem As ListItem
    For Each ListItem In ListView.ListItems
        UpdateListItem ListItem, ValueColumns
    Next ListItem
    
    If Not ListView.SelectedItem Is Nothing Then
        If ValueColumns.ValueColumns.Item(ListView.SelectedItem.Key).IsSelectable = False Then
            Set ListView.SelectedItem = Nothing
        End If
    End If
End Sub

Private Sub AddItem(ByVal ListView As MScomctllib.ListView, ByVal ValueColumn As ValueColumn)
    ListView.ListItems.Add Key:=ValueColumn.Name, Text:=ValueColumn.Name
End Sub

Private Sub UpdateListItem(ByVal ListItem As MScomctllib.ListItem, ByVal ValueColumns As ValueColumns)
    Dim ValueColumn As ValueColumn
    Set ValueColumn = ValueColumns.ValueColumns.Item(ListItem.Key)
    
    If ValueColumn.IsSelectable = False Then
        ListItem.ForeColor = RGB(128, 128, 128)
    End If
    
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