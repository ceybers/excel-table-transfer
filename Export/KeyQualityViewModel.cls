VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyQualityViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.ViewModels")
Option Explicit
Implements INotifyPropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged
    KeyColumn As KeyColumn
End Type
Private This As TState

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

Public Property Get KeyColumn() As KeyColumn
    Set KeyColumn = This.KeyColumn
End Property

Public Sub Load(ByVal KeyColumn As KeyColumn)
    Set This.KeyColumn = KeyColumn
End Sub
