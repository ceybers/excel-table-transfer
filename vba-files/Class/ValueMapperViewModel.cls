VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueMapperViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.ValueMapper.ViewModel"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged

    SrcValueColumnVM As ValuePickerViewModel
    DstValueColumnVM As ValuePickerViewModel
    
    EnableNontext As Boolean
    EnableNonunique As Boolean
End Type
Private This As TState

Private Function IViewModel_IsValid() As Boolean
    If This.SrcValueColumnVM.Selected Is Nothing Then Exit Function
    If This.DstValueColumnVM.Selected Is Nothing Then Exit Function
    
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
Public Property Get SrcValueColumnVM() As ValuePickerViewModel
    Set SrcValueColumnVM = This.SrcValueColumnVM
End Property

Public Property Get DstValueColumnVM() As ValuePickerViewModel
    Set DstValueColumnVM = This.DstValueColumnVM
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
    Set This.SrcValueColumnVM = ValuePickerViewModel.Create(SrcTable, "Value Columns for Source Table", Source, This.Notifier)
    Set This.DstValueColumnVM = ValuePickerViewModel.Create(DstTable, "Value Columns for Destination Table", Destination, This.Notifier)
End Sub


