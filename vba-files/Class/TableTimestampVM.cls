VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableTimestampVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.TableProps.ViewModel")
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged
Implements IChangeTracking

Private Const TIMESTAMP_ENABLED As String = "TIMESTAMP_ENABLED"
Private Const TIMESTAMP_ADDRESS As String = "TIMESTAMP_ADDRESS"

Private Type TState
    Notifier As INotifyPropertyChanged
    Settings As ISettings
    IsChanged As Boolean
    
    IsEnabled As Boolean
    Address As String
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
End Sub

'---
Public Property Get IsEnabled() As Boolean
    IsEnabled = This.IsEnabled
End Property

Public Property Let IsEnabled(ByVal vNewValue As Boolean)
    If This.IsEnabled = vNewValue Then Exit Property
    This.IsEnabled = vNewValue
    OnPropertyChanged "IsEnabled"
End Property

Public Property Get Address() As String
    Address = This.Address
End Property

Public Property Let Address(ByVal vNewValue As String)
    If This.Address = vNewValue Then Exit Property
    This.Address = vNewValue
    OnPropertyChanged "Address"
End Property

' ---
Public Sub Load(ByVal ListObject As ListObject, ByVal Settings As ISettings)
    Set This.Settings = Settings
    
    LoadSaved
End Sub

Private Sub LoadSaved()
    IsEnabled = This.Settings.GetFlag(TIMESTAMP_ENABLED)
    Address = This.Settings.GetSetting(TIMESTAMP_ADDRESS)
End Sub

Public Sub Commit()
    If This.IsChanged = False Then Exit Sub
    This.Settings.SetFlag TIMESTAMP_ENABLED, IsEnabled
    This.Settings.SetSetting TIMESTAMP_ADDRESS, Address
End Sub

Private Property Get IChangeTracking_IsChanged() As Boolean
    IChangeTracking_IsChanged = This.IsChanged
End Property

Private Sub IChangeTracking_AcceptChanges()
    Commit
    This.IsChanged = False
End Sub

