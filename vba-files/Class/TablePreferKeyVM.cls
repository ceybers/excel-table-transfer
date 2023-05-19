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
Public Property Get SelectedColumn() As String
    SelectedColumn = This.SelectedColumn
End Property

Public Property Let SelectedColumn(ByVal vNewValue As String)
    If This.SelectedColumn = vNewValue Then Exit Property
    This.SelectedColumn = vNewValue
    OnPropertyChanged "SelectedColumn"
End Property

Public Property Get Columns() As Scripting.Dictionary
    Set Columns = This.Columns
End Property

' ---
Public Sub Load(ByVal ListObject As ListObject, ByVal Settings As ISettings)
    Set This.Settings = Settings
    
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
        If This.Columns.Item(Key) = SavedColumn Then
                SelectedColumn = This.Columns.Item(Key)
            Exit Sub
        End If
    Next Key
End Sub

Public Sub Commit()
    If This.IsChanged = False Then Exit Sub
    This.Settings.SetSetting PREFER_KEY_COLUMN, SelectedColumn
End Sub

Private Property Get IChangeTracking_IsChanged() As Boolean
   IChangeTracking_IsChanged = This.IsChanged
End Property

Private Sub IChangeTracking_AcceptChanges()
    If This.IsChanged = False Then Exit Sub
    Commit
    This.IsChanged = False
End Sub
