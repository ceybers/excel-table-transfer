VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CountryToListViewConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.Example2.Converters"
Option Explicit
Implements IValueConverter

Private Sub IValueConverter_Convert(ByVal Source As Object, ByVal Target As Object)
    If TypeOf Source Is Scripting.Dictionary Then
        ConvertListItems Source, Target
    Else
        ConvertSelectedItem Source, Target
    End If
End Sub

Private Sub IValueConverter_ConvertBack(ByVal PropertyBinding As IPropertyBinding) 'ByVal Source As Object, ByVal Target As Object)
    'noop
End Sub

Private Sub ConvertSelectedItem(ByVal SelectedItemKey As String, ByVal ListView As ListView)
    Dim ListItem As ListItem
    For Each ListItem In ListView.ListItems
        If ListItem.Key = SelectedItemKey Then
            Set ListView.SelectedItem = ListItem
            Exit Sub
        End If
    Next ListItem
End Sub

Private Sub ConvertListItems(ByVal Countries As Scripting.Dictionary, ByVal ListView As ListView)
    Dim Key As Variant
    For Each Key In Countries
        Dim ThisCountry As Country
        Set ThisCountry = Countries.Item(Key)
        Debug.Assert TypeOf ThisCountry Is Country
        
        Dim ListItem As ListItem
        Set ListItem = ListView.ListItems.Add(Key:=ThisCountry.Name, Text:=ThisCountry.Name)
    Next Key
End Sub

Public Sub InitializeListView(ByVal ListView As ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Country Name"
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .FullRowSelect = True
    End With
End Sub
