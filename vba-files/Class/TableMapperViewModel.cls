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
Implements IHandlePropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged

    SrcTableVM As TablePickerViewModel
    DstTableVM As TablePickerViewModel
    
    EnableProtected As Boolean
    EnableReadonly As Boolean
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
End Sub

'---
Public Property Get SrcTableVM() As TablePickerViewModel
    Set SrcTableVM = This.SrcTableVM
End Property

Public Property Get DstTableVM() As TablePickerViewModel
    Set DstTableVM = This.DstTableVM
End Property

Public Property Get EnableProtected() As Boolean
    EnableProtected = This.EnableProtected
End Property

Public Property Let EnableProtected(ByVal vNewValue As Boolean)
    If This.EnableProtected = vNewValue Then Exit Property
    This.EnableProtected = vNewValue
    OnPropertyChanged "EnableProtected"
End Property

Public Property Get EnableReadonly() As Boolean
    EnableReadonly = This.EnableReadonly
End Property

Public Property Let EnableReadonly(ByVal vNewValue As Boolean)
    If This.EnableReadonly = vNewValue Then Exit Property
    This.EnableReadonly = vNewValue
    OnPropertyChanged "EnableReadonly"
End Property

'---
Public Sub Load(ByVal SrcTable As ListObject, ByVal DstTable As ListObject)
    Set This.SrcTableVM = TablePickerViewModel.Create(SrcTable, "Source Table", Source, This.Notifier, Me)
    Set This.DstTableVM = TablePickerViewModel.Create(DstTable, "Destination Table", Destination, This.Notifier, Me)
    
    TrySelectOtherIfOnlyOne
End Sub

'@Description "If there are only two tables to choose from, and one is already selected on load, automatically select the other table."
Private Sub TrySelectOtherIfOnlyOne()
Attribute TrySelectOtherIfOnlyOne.VB_Description = "If there are only two tables to choose from, and one is already selected on load, automatically select the other table."
    ' TODO NYI
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    Select Case PropertyName
        Case "Selected"
            Dim SourceVM As TablePickerViewModel
            Set SourceVM = Source
            InformOtherVM SourceVM.Direction
    End Select
End Sub

Private Sub InformOtherVM(ByVal Direction As TransferDirection)
    Select Case Direction
        Case Source
            This.DstTableVM.ExcludeTable This.SrcTableVM.SelectedAsText
        Case Destination
            This.SrcTableVM.ExcludeTable This.DstTableVM.SelectedAsText
    End Select
End Sub
