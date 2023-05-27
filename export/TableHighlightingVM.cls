VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableHighlightingVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.TableProps.ViewModel")
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged

    ListObject As ListObject
    
    RemoveHighlightingCommand As ICommand
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
End Sub

Public Property Get HasHighlighting() As Boolean
    HasHighlighting = CheckListObjectHasHighlighting
End Property

Public Property Get RemoveHighlightingCommand() As ICommand
    Set RemoveHighlightingCommand = This.RemoveHighlightingCommand
End Property

Public Property Set HasHighlighting(ByVal vNewValue As ICommand)
    Set This.RemoveHighlightingCommand = vNewValue
End Property
' ---
Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
End Sub

Public Sub DoRemoveHighlighting()
    MsgBox "DoRemoveHighlighting NYI"
End Sub

Private Function CheckListObjectHasHighlighting() As Boolean
    CheckListObjectHasHighlighting = True
End Function
