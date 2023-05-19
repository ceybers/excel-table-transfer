VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableMapperViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.TableMapper.ViewModel"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged

    SrcTableVM As TablePickerViewModel
    DstTableVM As TablePickerViewModel
End Type
Private This As TState

Private Function IViewModel_IsValid() As Boolean
    If This.SrcTableVM.Selected Is Nothing Then Exit Function
    If This.DstTableVM.Selected Is Nothing Then Exit Function
    If This.SrcTableVM.Selected.Key = This.DstTableVM.Selected.Key Then Exit Function
    
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
    Set This.SrcTableVM = New TablePickerViewModel
    Set This.DstTableVM = New TablePickerViewModel
End Sub

'---
Public Property Get SrcTableVM() As TablePickerViewModel
    Set SrcTableVM = This.SrcTableVM
End Property

Public Property Get DstTableVM() As TablePickerViewModel
    Set DstTableVM = This.DstTableVM
End Property

Public Property Get DummyBoolean() As Boolean
    DummyBoolean = True
End Property

'---
Public Sub Load(ByVal ListObject As ListObject)
    Dim Worksheet As Worksheet
    Dim Workbook As Workbook
    Set Worksheet = ListObject.Parent
    Set Workbook = Worksheet.Parent
    
    This.SrcTableVM.Load ListObject
    SrcTableVM.Caption = "Source Table"
    
    This.DstTableVM.Load ListObject
    DstTableVM.Caption = "Destination Table"
End Sub
