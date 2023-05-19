VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableStarColumnsVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.TableProps.ViewModel")
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged
Implements IChangeTracking

Private Const STAR_COLUMN_COLLECTION As String = "STAR_COLUMN_COLLECTION"

Private Type TState
    Notifier As INotifyPropertyChanged
    Settings As ISettings
    IsChanged As Boolean
    
    SelectedColumn As String
    ColumnProperties As ColumnProperties
    Columns As Scripting.Dictionary ' of type ColumnProperty
End Type
Private This As TState

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = True
End Function

Private Function IViewModel_Self() As IViewModel
    Set IViewModel_Self = Me
End Function

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    This.Notifier.RegisterHandler Handler
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    This.Notifier.OnPropertyChanged Source, PropertyName
End Sub

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    INotifyPropertyChanged_OnPropertyChanged Me, PropertyName
    This.IsChanged = True
End Sub

Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifier
    Set This.Columns = New Scripting.Dictionary
End Sub

'---
Public Property Get SelectedColumn() As String
    SelectedColumn = This.SelectedColumn
End Property

Public Property Let SelectedColumn(ByVal vNewValue As String)
    This.SelectedColumn = vNewValue
End Property

Public Property Get Columns() As Scripting.Dictionary
    Set Columns = This.Columns
End Property

' ---
Public Sub Load(ByVal ListObject As ListObject, ByVal Settings As ISettings)
    Set This.Settings = Settings
    
    Set This.ColumnProperties = New ColumnProperties
    This.ColumnProperties.Load ListObject
    
    Set This.Columns = This.ColumnProperties.GetDictionary
    
    LoadSettings
End Sub

Private Sub LoadSettings()
    Dim CheckedColumns As Collection
    Set CheckedColumns = This.Settings.GetCollection(STAR_COLUMN_COLLECTION)
    
    Dim Key As Variant
    For Each Key In CheckedColumns
        Dim ColumnProperty As ColumnProperty
        Set ColumnProperty = This.Columns.Item(Key)
        ColumnProperty.Checked = True
    Next Key
End Sub

Public Sub Commit()
    If This.IsChanged = False Then Exit Sub
    Dim CheckedColumns As Collection
    Set CheckedColumns = This.ColumnProperties.GetChecked
    This.Settings.SetCollection STAR_COLUMN_COLLECTION, CheckedColumns
End Sub

Private Property Get IChangeTracking_IsChanged() As Boolean
   IChangeTracking_IsChanged = This.IsChanged
End Property

Private Sub IChangeTracking_AcceptChanges()
    Commit
    This.IsChanged = False
End Sub

Public Sub TrySetColumnChecked(ByVal ColumnKey As String, ByVal vNewValue As Boolean)
    Dim ColumnProperty As ColumnProperty
    Set ColumnProperty = This.Columns.Item(ColumnKey)
    If ColumnProperty.Checked = vNewValue Then Exit Sub
    ColumnProperty.Checked = vNewValue
    OnPropertyChanged "ColumnCheckedChanged"
End Sub

Public Sub TrySetColumnSelected(ByVal ColumnKey As String, ByVal vNewValue As Boolean)
    Dim ColumnProperty As ColumnProperty
    Set ColumnProperty = This.Columns.Item(ColumnKey)
    If ColumnProperty.Selected = vNewValue Then Exit Sub
    ColumnProperty.Selected = vNewValue
    OnPropertyChanged "ColumnSelectedChanged"
End Sub

Public Function CanReset() As Boolean
    Dim Key As Variant
    For Each Key In This.Columns.Keys
        Dim ColumnProperty As ColumnProperty
        Set ColumnProperty = This.Columns.Item(Key)
        If ColumnProperty.Checked Then
            CanReset = True
            Exit Function
        End If
    Next Key
End Function

Public Sub DoReset()
    Dim Key As Variant
    For Each Key In This.Columns.Keys
        TrySetColumnChecked Key, False
    Next Key
End Sub
