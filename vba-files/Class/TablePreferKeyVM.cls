VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablePreferKeyVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.TableProps.ViewModel")
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged
Implements IChangeTracking

Private Const PREFER_KEY_COLUMN As String = "PREFER_KEY_COLUMN"

Private Type TState
    Notifier As INotifyPropertyChanged
    Settings As ISettings
    IsChanged As Boolean
    
    ListObject As ListObject
    
    Selected As ListColumn
    SelectedColumn As String
    Columns As Scripting.Dictionary
    KeyColumns As KeyColumnProps
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
    'Set This.Columns = New Scripting.Dictionary
    Set This.KeyColumns = New KeyColumnProps
End Sub

'---
Public Property Get Selected() As ListColumn
    Set Selected = This.Selected
End Property

Public Property Get SelectedColumn() As String
    SelectedColumn = This.SelectedColumn
End Property

Public Property Let SelectedColumn(ByVal vNewValue As String)
    If This.SelectedColumn = vNewValue Then Exit Property
    This.SelectedColumn = vNewValue
    TrySetSelectedFromSuffixedName
    OnPropertyChanged "SelectedColumn"
End Property

Public Property Get Columns() As Scripting.Dictionary
    Set Columns = This.Columns
End Property

' ---
Public Sub Load(ByVal ListObject As ListObject, ByVal Settings As ISettings)
    Set This.Settings = Settings
    Set This.ListObject = ListObject
    
    This.KeyColumns.Load ListObject
    
    LoadColumns
    LoadSelectedColumn
End Sub

Private Sub LoadColumns()
    Set This.Columns = This.KeyColumns.GetDictionary
    
    SelectedColumn = This.Columns.Keys(0)
End Sub

Private Sub LoadSelectedColumn()
    Dim SavedColumn As String
    SavedColumn = This.Settings.GetSetting(PREFER_KEY_COLUMN)
    
    Dim Key As Variant
    For Each Key In This.Columns.Keys
        If Key = SavedColumn Then
            SelectedColumn = This.Columns.Item(Key)
            Set This.Selected = This.ListObject.ListColumns(Key)
        End If
    Next Key
End Sub

Public Sub Commit()
    If This.IsChanged = False Then Exit Sub
    If Not This.Selected Is Nothing Then
        This.Settings.SetSetting PREFER_KEY_COLUMN, This.Selected.Name
    Else
        This.Settings.SetSetting PREFER_KEY_COLUMN, vbNullString
    End If
End Sub

Private Property Get IChangeTracking_IsChanged() As Boolean
   IChangeTracking_IsChanged = This.IsChanged
End Property

Private Sub IChangeTracking_AcceptChanges()
    If This.IsChanged = False Then Exit Sub
    Commit
    This.IsChanged = False
End Sub

Private Sub TrySetSelectedFromSuffixedName()
    Debug.Print "TrySetSelectedFromSuffixedName: "; This.SelectedColumn
    
    Dim Dict As Scripting.Dictionary
    Set Dict = This.KeyColumns.GetDictionary
    
    Dim ActualColumnName As String
    Dim Key As Variant
    For Each Key In Dict.Keys
        If Dict.Item(Key) = This.SelectedColumn Then
            ActualColumnName = Key
            Exit For
        End If
    Next Key
    
    Dim ListColumn As ListColumn
    For Each ListColumn In This.ListObject.ListColumns
        If ListColumn.Name = ActualColumnName Then
            Set This.Selected = ListColumn
            Exit Sub
        End If
    Next ListColumn
    
    ' ListColumn doesn't exist in ListObject
    Set This.Selected = Nothing
End Sub
