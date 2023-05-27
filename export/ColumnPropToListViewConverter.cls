VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnPropToListViewConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule AssignmentNotUsed
'@PredeclaredId
'@Folder "MVVM.TableProps.Converters"
Option Explicit
Implements IValueConverter

Private Sub IValueConverter_Convert(ByVal Source As Object, ByVal Target As Object)
    
    'If Source Is Nothing Then
    '    ConvertSelectedItem vbNullString, Target
    'ElseIf TypeOf Source Is Scripting.Dictionary Then
        ConvertListItems Source, Target
    'Else
    '    ConvertSelectedItem Source, Target
    'End If
End Sub

Private Sub IValueConverter_ConvertBack(ByVal PropertyBinding As IPropertyBinding)
    Dim ListView As ListView
    Set ListView = PropertyBinding.Target
    
    Dim VM As TableStarColumnsVM
    Set VM = PropertyBinding.Source
    Dim Columns As Scripting.Dictionary
    Set Columns = CallByName(VM, PropertyBinding.SourcePropertyPath, VbGet)
    
    Dim ListItem As ListItem
    For Each ListItem In ListView.ListItems
        VM.TrySetColumnChecked ListItem.Key, ListItem.Checked
        VM.TrySetColumnSelected ListItem.Key, ListItem.Selected
    Next
End Sub

Private Sub ConvertSelectedItem(ByVal SelectedItemKey As String, ByVal ListView As ListView)
    'If SelectedItemKey = vbNullString Then
    '    Set ListView.SelectedItem = Nothing
    '    Exit Sub
    'End If
   '
   ' Dim ListItem As ListItem
   ' For Each ListItem In ListView.ListItems
   '     If ListItem.Key = SelectedItemKey Then
   '         Set ListView.SelectedItem = ListItem
   '         Exit Sub
   '     End If
   ' Next ListItem
End Sub

Private Sub ConvertListItems(ByVal Columns As Scripting.Dictionary, ByVal ListView As ListView)
    Dim Key As Variant
    For Each Key In Columns
        Dim ThisColumn As ColumnProperty
        Set ThisColumn = Columns.Item(Key)
        Debug.Assert TypeOf ThisColumn Is ColumnProperty
        LoadListItem ThisColumn, ListView
    Next Key
End Sub

Private Sub LoadListItem(ByVal ColumnProperty As ColumnProperty, ByVal ListView As ListView)
    Dim ListItem As ListItem

    For Each ListItem In ListView.ListItems
        If ListItem.Key = ColumnProperty.Name Then
            ListItem.Text = ColumnProperty.Name
            ListItem.ListSubItems.Clear
            Exit For
        End If
    Next ListItem
    
    If ListItem Is Nothing Then
        Set ListItem = ListView.ListItems.Add(Key:=ColumnProperty.Name, Text:=ColumnProperty.Name)
    End If
    
    ListItem.ListSubItems.Add Text:=ColumnProperty.ColumnType
    ListItem.ListSubItems.Add Text:=ColumnProperty.R1C1
    ListItem.ListSubItems.Add Text:=ListColumnPropHelpers.EnumToString(ColumnProperty.Blanks)
    ListItem.ListSubItems.Add Text:=ListColumnPropHelpers.EnumToString(ColumnProperty.Errors)
    ListItem.ListSubItems.Add Text:=ListColumnPropHelpers.EnumToString(ColumnProperty.Locked)
    ListItem.ListSubItems.Add Text:=ListColumnPropHelpers.EnumToString(ColumnProperty.Formula)
    ListItem.ListSubItems.Add Text:=ListColumnPropHelpers.EnumToString(ColumnProperty.Validation)
    ListItem.ListSubItems.Add Text:=ListColumnPropHelpers.EnumToString(ColumnProperty.Unique)
    
    ListItem.Checked = ColumnProperty.Checked
    ListItem.Selected = ColumnProperty.Selected
End Sub

Public Sub InitializeListView(ByVal ListView As ListView)
    Debug.Assert Not ListView Is Nothing
     
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add Text:="Name"
        .ColumnHeaders.Add Text:="Type", Width:=48
        .ColumnHeaders.Add Text:="R1C1", Width:=25
        .ColumnHeaders.Add Text:="Blanks", Width:=32
        .ColumnHeaders.Add Text:="Errors", Width:=32
        .ColumnHeaders.Add Text:="Locked", Width:=32
        .ColumnHeaders.Add Text:="Formula", Width:=32
        .ColumnHeaders.Add Text:="Validation", Width:=32
        .ColumnHeaders.Add Text:="Unique", Width:=32
        
        .Appearance = cc3D
        .BorderStyle = ccNone
        .CheckBoxes = True
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
    End With
End Sub
