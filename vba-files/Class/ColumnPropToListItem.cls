VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnPropToListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Converters"
Option Explicit
Implements IValueConverter

Private Sub IValueConverter_Convert(ByVal Control As Object, ByVal Object As Object)
    Debug.Assert TypeOf Control Is ListView
    Debug.Assert TypeOf Object Is ColumnProperty
    
    Dim ListView As ListView
    Set ListView = Control
    
    Dim ColumnProperty As ColumnProperty
    Set ColumnProperty = Object
    
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(Text:=ColumnProperty.Name)
    ListItem.ListSubItems.Add Text:=ColumnProperty.ColumnType
    ListItem.ListSubItems.Add Text:=ColumnProperty.R1C1
    ListItem.ListSubItems.Add Text:=ListColumnPropHelpers.EnumToString(ColumnProperty.Blanks)
    ListItem.ListSubItems.Add Text:=ListColumnPropHelpers.EnumToString(ColumnProperty.Errors)
    ListItem.ListSubItems.Add Text:=ListColumnPropHelpers.EnumToString(ColumnProperty.Locked)
    ListItem.ListSubItems.Add Text:=ListColumnPropHelpers.EnumToString(ColumnProperty.Formula)
    ListItem.ListSubItems.Add Text:=ListColumnPropHelpers.EnumToString(ColumnProperty.Validation)
    ListItem.ListSubItems.Add Text:=ListColumnPropHelpers.EnumToString(ColumnProperty.Unique)
End Sub

Private Sub IValueConverter_Initialize(ByVal Control As Object)
    Debug.Assert TypeOf Control Is ListView
    Dim ListView As ListView
    Set ListView = Control
    
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
        .view = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
    End With
End Sub

