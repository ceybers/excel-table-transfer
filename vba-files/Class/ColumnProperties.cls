VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TState
    ListObject As ListObject
    Item As Collection
End Type
Private This As TState

Public Property Get Name() As String
    Name = This.ListObject.Name
End Property

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    LoadCollection
End Sub

Private Sub LoadCollection()
    Dim ListColumn As ListColumn
    Dim ColumnProperty As ColumnProperty
    For Each ListColumn In This.ListObject.ListColumns
        Set ColumnProperty = New ColumnProperty
        ColumnProperty.Load ListColumn
        This.Item.Add ColumnProperty, Key:=ListColumn.Name
    Next ListColumn
End Sub

Private Sub Class_Initialize()
    Set This.Item = New Collection
End Sub

Public Sub LoadListView(ByVal ListView As MSComctlLib.ListView)
    Dim Converter As IValueConverter
    Set Converter = New ColumnPropToListItem
    
    Converter.Initialize ListView 'CHK
    
    Dim ColumnProperty As ColumnProperty
    For Each ColumnProperty In This.Item
        Converter.Convert ListView, ColumnProperty
    Next ColumnProperty
End Sub

Public Function GetSelected() As Collection
    Set GetSelected = New Collection
    Dim ColumnProperty As ColumnProperty
    For Each ColumnProperty In This.Item
        If ColumnProperty.Selected Then
            GetSelected.Add Item:=ColumnProperty, Key:=ColumnProperty.Name
        End If
    Next ColumnProperty
End Function
