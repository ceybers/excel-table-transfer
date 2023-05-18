VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.TableProps.Model"
Option Explicit

Event SelectedChanged()
Event CollectionChanged()

Private Type TState
    ListObject As ListObject
    Item As Collection
    Selected As ColumnProperty
    ActivateOnWorksheet As Boolean
End Type
Private This As TState

Public Property Get ActivateOnWorksheet() As Boolean
    ActivateOnWorksheet = This.ActivateOnWorksheet
End Property

Public Property Let ActivateOnWorksheet(ByVal vNewValue As Boolean)
    This.ActivateOnWorksheet = vNewValue
End Property

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

Private Sub class_initialize()
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
    RaiseEvent CollectionChanged
End Sub

' TODO Rename this. Selected is the one that has the focus. This should get GetAllChecked or something to that effect.
Public Function GetSelected() As Collection
    Set GetSelected = New Collection
    Dim ColumnProperty As ColumnProperty
    For Each ColumnProperty In This.Item
        If ColumnProperty.Selected Then
            GetSelected.Add Item:=ColumnProperty, Key:=ColumnProperty.Name
        End If
    Next ColumnProperty
End Function

Public Function TrySelectByName(ByVal Name As String) As Boolean
    Dim ColumnProperty As ColumnProperty
    For Each ColumnProperty In This.Item
        If ColumnProperty.Name = Name Then
            Set This.Selected = ColumnProperty
            If This.ActivateOnWorksheet Then ColumnProperty.Activate
            TrySelectByName = True
            Exit Function
        End If
    Next ColumnProperty
End Function

Public Sub Reset()
    For Each ColumnProperty In This.Item
        If ColumnProperty.Name = Name Then
            Set This.Selected = False
        End If
    Next ColumnProperty
    Set This.Selected = Nothing
    RaiseEvent SelectedChanged
End Sub

Public Function CanReset() As Boolean
    For Each ColumnProperty In This.Item
        If ColumnProperty.Selected = True Then
            CanReset = True
            Exit Function
        End If
    Next ColumnProperty
End Function
