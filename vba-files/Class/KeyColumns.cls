VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyColumns"
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
End Type
Private This As TState

Public Property Get Name() As String
    Name = This.ListObject.Name
End Property

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    LoadCollection
End Sub

Public Property Get Selected() As ColumnProperty
    Set Selected = This.Selected
End Property

Public Property Set Selected(ByVal vNewValue As ColumnProperty)
    If Not This.Selected Is Nothing Then
        If This.Selected.Name = vNewValue.Name Then
            Exit Property
        End If
    End If
    
    Set This.Selected = vNewValue
    RaiseEvent SelectedChanged
End Property


Private Sub LoadCollection()
    Dim ListColumn As ListColumn
    Dim ColumnProperty As ColumnProperty
    For Each ListColumn In This.ListObject.ListColumns
        Set ColumnProperty = New ColumnProperty
        ColumnProperty.Load ListColumn
        This.Item.Add ColumnProperty, Key:=ListColumn.Name
    Next ListColumn
    RaiseEvent CollectionChanged
End Sub

Private Sub class_initialize()
    Set This.Item = New Collection
End Sub

Public Sub LoadComboBox(ByVal ComboBox As ComboBox)
    Dim Converter As IValueConverter
    Set Converter = New ColumnPropToComboBox
    
    Converter.Initialize ComboBox 'CHK
    
    Dim ColumnProperty As ColumnProperty
    For Each ColumnProperty In This.Item
        Converter.Convert ComboBox, ColumnProperty
    Next ColumnProperty
    
    ' This will always be > 0 because we add an "(undefined)" default item
    'If This.Item.Count > 0 Then
    '    TrySelectByName This.Item.Item(1).Name
    'End If
    
    If Not This.Selected Is Nothing Then
        ComboBox.Value = This.Selected.Name
        TrySelectByName This.Selected.Name
        ' TODO Should not the combobox be conditionally executed if TrySelect works?
    End If
End Sub

Public Function TrySelectByName(ByVal Name As String) As Boolean
    If Not This.Selected Is Nothing Then
        If This.Selected.Name = Name Then
            Exit Function
        End If
    End If
    
    Dim ColumnProperty As ColumnProperty
    For Each ColumnProperty In This.Item
        If ColumnProperty.Name = Name Then ' TODO InStr this because of (non-text) etc. suffixes
            Set Selected = ColumnProperty
            TrySelectByName = True
            Exit Function
        End If
    Next ColumnProperty
End Function

