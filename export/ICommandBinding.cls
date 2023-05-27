VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommandBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder MVVM.Infrastructure.Abstract
Option Explicit

Public Property Get Context() As IAppContext
End Property

'@Description "The ViewModel associated with this Command Binding."
Public Property Get Source() As IViewModel
Attribute Source.VB_Description = "The ViewModel associated with this Command Binding."
End Property

Public Property Get Command() As ICommand
End Property

'@Description "The UserForm Control that will be bound to the ViewModel's Property."
Public Property Get Target() As Object
Attribute Target.VB_Description = "The UserForm Control that will be bound to the ViewModel's Property."
End Property

'@Description "Evaluates whether the command can execute given the binding context."
Public Sub EvaluateCanExecute(ByVal Context As Object)
Attribute EvaluateCanExecute.VB_Description = "Evaluates whether the command can execute given the binding context."
End Sub
