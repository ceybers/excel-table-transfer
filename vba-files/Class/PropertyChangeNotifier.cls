VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyChangeNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule AssignmentNotUsed
'@Folder "MVVM.Infrastructure.Bindings"
Option Explicit
Implements INotifyPropertyChanged

Private Type TState
    Handlers As Collection
End Type
Private This As TState

Private Sub Class_Initialize()
    Set This.Handlers = New Collection
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    Dim Handler As IHandlePropertyChanged
    For Each Handler In This.Handlers
        Handler.HandlePropertyChanged Source, PropertyName
    Next
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    Dim ThisHandler As IHandlePropertyChanged
    For Each ThisHandler In This.Handlers
        If ThisHandler Is Handler Then
            Exit Sub
        End If
    Next ThisHandler
    
    This.Handlers.Add Handler
End Sub
