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
    Item As Scripting.Dictionary
End Type
Private This As TState

Public Property Get Name() As String
    Name = This.ListObject.Name
End Property

Private Sub Class_Initialize()
    Set This.Item = New Scripting.Dictionary
End Sub

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    LoadColumns
End Sub

Private Sub LoadColumns()
    Dim ListColumn As ListColumn
    Dim ColumnProperty As ColumnProperty
    For Each ListColumn In This.ListObject.ListColumns
        Set ColumnProperty = New ColumnProperty
        ColumnProperty.Load ListColumn
        This.Item.Add Key:=ColumnProperty.Name, Item:=ColumnProperty
    Next ListColumn
End Sub

Public Function GetDictionary() As Scripting.Dictionary
    Set GetDictionary = This.Item
End Function

Public Function GetChecked() As Collection
    Dim Result As Collection
    Set Result = New Collection
    
    Dim Key As Variant
    For Each Key In This.Item.Keys
        Dim ColumnProperty As ColumnProperty
        Set ColumnProperty = This.Item.Item(Key)
        If ColumnProperty.Checked Then
            Result.Add Item:=ColumnProperty.Name, Key:=ColumnProperty.Name
        End If
    Next Key
    Set GetChecked = Result
End Function
