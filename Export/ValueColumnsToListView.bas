Attribute VB_Name = "ValueColumnsToListView"
'@Folder("MVVM2.ValueConverters")
Option Explicit

Private Const MSO_ITEM As String = "lblSelCol"
Private Const MSO_KEY_COL As String = "lblKey2"
Private Const MSO_FORMULA As String = "lblFunction"
Private Const MSO_LINK As String = "lblLink"
Private Const MSO_ERRORS As String = "lblError"

Private Const MSO_TYPE_STRING As String = "lblDataString"
Private Const MSO_TYPE_LONG As String = "lblDataLong"
Private Const MSO_TYPE_BOOLEAN As String = "lblDataBoolean"
Private Const MSO_TYPE_DATE As String = "lblDataDate"
Private Const MSO_TYPE_CURRENCY As String = "lblDataCurrency"
Private Const MSO_EMPTY As String = "lblDataEmpty"

Public Sub Initialize(ByVal ListView As MSComctlLib.ListView)
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

Public Sub Load(ByVal ListView As MSComctlLib.ListView, ByVal ValueColumns As ValueColumns)
    Debug.Assert Not ValueColumns Is Nothing
    
    If ListView.ListItems.Count <> ValueColumns.ValueColumns.Count Then
        LoadItems ListView, ByVal ValueColumns
    End If
        
    UpdateItems ListView, ByVal ValueColumns
End Sub

Private Sub LoadItems(ByVal ListView As MSComctlLib.ListView, ByVal ValueColumns As ValueColumns)
    ListView.ListItems.Clear
    Dim ValueColumn As ValueColumn
    For Each ValueColumn In ValueColumns.ValueColumns
        AddItem ListView, ValueColumn
    Next ValueColumn
End Sub

Private Sub UpdateItems(ByVal ListView As MSComctlLib.ListView, ByVal ValueColumns As ValueColumns)
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

Private Sub AddItem(ByVal ListView As MSComctlLib.ListView, ByVal ValueColumn As ValueColumn)
    ListView.ListItems.Add Key:=ValueColumn.Name, Text:=ValueColumn.Name
End Sub


Private Sub UpdateListItem(ByVal ListItem As MSComctlLib.ListItem, ByVal ValueColumns As ValueColumns)
    Dim ValueColumn As ValueColumn
    Set ValueColumn = ValueColumns.ValueColumns.Item(ListItem.Key)
    
    If ValueColumn.IsSelectable = False Then
        ListItem.ForeColor = RGB(128, 128, 128)
    End If
    
    Select Case True
        '@Ignore UnassignedVariableUsage
        Case ValueColumn.IsKeyColumn
            ListItem.SmallIcon = MSO_KEY_COL
        Case ValueColumn.IsMapped
            ListItem.SmallIcon = MSO_LINK
            ListItem.Bold = True
        Case ValueColumn.IsFormula
            ListItem.SmallIcon = MSO_FORMULA
        Case ValueColumn.IsEmpty
            ListItem.SmallIcon = MSO_EMPTY
        Case ValueColumn.HasErrors
            ListItem.SmallIcon = MSO_ERRORS
        Case ValueColumn.HasNumbers
            ListItem.SmallIcon = MSO_TYPE_LONG
        Case Else
            ListItem.SmallIcon = MSO_TYPE_STRING
    End Select
End Sub

