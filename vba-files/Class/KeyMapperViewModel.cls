VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyMapperViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.KeyMapper.ViewModel"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged

    SrcKeyColumnVM As KeyPickerViewModel
    DstKeyColumnVM As KeyPickerViewModel
    
    EnableNontext As Boolean
    EnableNonunique As Boolean
End Type
Private This As TState

Private Function IViewModel_IsValid() As Boolean
    If This.SrcKeyColumnVM.Selected Is Nothing Then Exit Function
    If This.DstKeyColumnVM.Selected Is Nothing Then Exit Function
    
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
End Sub

Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifier
End Sub

'---
Public Property Get SrcKeyColumnVM() As KeyPickerViewModel
    Set SrcKeyColumnVM = This.SrcKeyColumnVM
End Property

Public Property Get DstKeyColumnVM() As KeyPickerViewModel
    Set DstKeyColumnVM = This.DstKeyColumnVM
End Property

Public Property Get EnableNontext() As Boolean
    EnableNontext = This.EnableNontext
End Property

Public Property Let EnableNontext(ByVal vNewValue As Boolean)
    If This.EnableNontext = vNewValue Then Exit Property
    This.EnableNontext = vNewValue
    OnPropertyChanged "EnableNontext"
End Property

Public Property Get EnableNonunique() As Boolean
    EnableNonunique = This.EnableNonunique
End Property

Public Property Let EnableNonunique(ByVal vNewValue As Boolean)
    If This.EnableNonunique = vNewValue Then Exit Property
    This.EnableNonunique = vNewValue
    OnPropertyChanged "EnableNonunique"
End Property

'---
Public Sub Load(ByVal SrcTable As ListObject, ByVal DstTable As ListObject)
    Set This.SrcKeyColumnVM = KeyPickerViewModel.Create(SrcTable, "Key Column for Source Table", Source, This.Notifier)
    Set This.DstKeyColumnVM = KeyPickerViewModel.Create(DstTable, "Key Column for Destination Table", Destination, This.Notifier)
End Sub

