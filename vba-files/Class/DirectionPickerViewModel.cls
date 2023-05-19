VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DirectionPickerViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.DirectionPicker.ViewModel"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged

    TableDetailsVM As TableDetailsVM
    TableLocationVM As TableLocationVM
    
    Result As TransferDirection
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
End Sub

Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifier

    Set This.TableDetailsVM = New TableDetailsVM
    Set This.TableLocationVM = New TableLocationVM
End Sub

'---
Public Property Get TableDetailsVM() As TableDetailsVM
    Set TableDetailsVM = This.TableDetailsVM
End Property

Public Property Get TableLocationVM() As TableLocationVM
    Set TableLocationVM = This.TableLocationVM
End Property

Public Property Get Result() As TransferDirection
    Result = This.Result
End Property

Public Property Let Result(ByVal vNewValue As TransferDirection)
    This.Result = vNewValue
End Property

'---
Public Sub Load(ByVal ListObject As ListObject)
    Dim Worksheet As Worksheet
    Dim Workbook As Workbook
    Set Worksheet = ListObject.Parent
    Set Workbook = Worksheet.Parent
    
    This.TableDetailsVM.Load ListObject
    This.TableLocationVM.Load Workbook
    
    This.Result = TransferDirection.Undefined
End Sub
