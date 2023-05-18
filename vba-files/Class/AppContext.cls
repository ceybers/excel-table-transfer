VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Infrastructure"
Option Explicit
Implements IAppContext

Private Type TAppContext
    BindingManager As BindingManager
    CommandManager As CommandManager
End Type
Private This As TAppContext

Private Property Get IAppContext_BindingManager() As BindingManager
    Set IAppContext_BindingManager = This.BindingManager
End Property

Private Property Get IAppContext_CommandManager() As CommandManager
    Set IAppContext_CommandManager = This.CommandManager
End Property

Private Sub Class_Initialize()
    Set This.BindingManager = New BindingManager
    Set This.BindingManager.Context = Me
    Set This.CommandManager = New CommandManager
End Sub

Public Sub RegisterViewModel(ByVal ViewModel As INotifyPropertyChanged)
    ViewModel.RegisterHandler This.BindingManager
End Sub