VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder MVVM.Infrastructure.Bindings.CommandBindings
Option Explicit
Implements IHandlePropertyChanged

Private Type TBindingManager
    Commands As Collection
End Type
Private This As TBindingManager

Private Sub Class_Initialize()
    Set This.Commands = New Collection
End Sub

Public Function BindCommand(ByVal Context As IAppContext, ByVal ViewModel As IViewModel, _
    ByVal Command As ICommand, ByVal Control As Object) As ICommandBinding
    
    Dim NewBinding As ICommandBinding
        
    If TypeOf Control Is MSForms.CommandButton Then
        Set NewBinding = CommandButtonCommandBinding.Create(Context, ViewModel, Command, Control)
    ElseIf TypeOf Control Is MSComctlLib.ListView Then
        Set NewBinding = ListViewCommandBinding.Create(Context, ViewModel, Command, Control)
    Else
        Debug.Assert False
    End If
    
    This.Commands.Add Item:=NewBinding
    Set BindCommand = NewBinding
End Function

Public Sub EvaluateCanExecute(ByVal Context As Object)
    Dim Binding As ICommandBinding
    For Each Binding In This.Commands
        Binding.EvaluateCanExecute Context
    Next
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    EvaluateCanExecute Source
End Sub

