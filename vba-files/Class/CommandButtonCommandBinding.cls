VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandButtonCommandBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder MVVM.Infrastructure.Bindings.CommandBindings
Option Explicit
Implements ICommandBinding

Private Const DO_DEBUG As Boolean = False

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents mTarget As MSForms.CommandButton
Attribute mTarget.VB_VarHelpID = -1

Private Type TCommandBinding
    AppContext As IAppContext
    Source As IViewModel
    Command As ICommand
End Type
Private This As TCommandBinding

Private Property Get ICommandBinding_Context() As IAppContext
    Set ICommandBinding_Context = This.AppContext
End Property

Private Property Get ICommandBinding_Source() As IViewModel
    Set ICommandBinding_Source = This.Source
End Property

Private Property Get ICommandBinding_Command() As ICommand
    Set ICommandBinding_Command = This.Command
End Property

Private Property Get ICommandBinding_Target() As Object
    Set ICommandBinding_Target = mTarget
End Property

Public Property Get AppContext() As IAppContext
    Set AppContext = This.AppContext
End Property

Public Property Set AppContext(ByVal vNewValue As IAppContext)
    Set This.AppContext = vNewValue
End Property

Public Property Get Source() As IViewModel
    Set Source = This.Source
End Property

Public Property Set Source(ByVal vNewValue As IViewModel)
    Set This.Source = vNewValue
End Property

Public Property Get Command() As ICommand
    Set Command = This.Command
End Property

Public Property Set Command(ByVal vNewValue As ICommand)
    Set This.Command = vNewValue
End Property

Public Sub SetTarget(ByVal Target As Object)
    Set mTarget = Target
    mTarget.Caption = This.Command.Caption
    mTarget.Enabled = This.Command.CanExecute
End Sub

Public Function Create(ByVal AppContext As IAppContext, ByVal Source As IViewModel, ByVal Command As ICommand, ByVal Target As MSForms.CommandButton) As ICommandBinding
    Dim NewBinding As CommandButtonCommandBinding
    Set NewBinding = New CommandButtonCommandBinding
    With NewBinding
        Set .AppContext = AppContext
        Set .Source = Source
        Set .Command = Command
        .SetTarget Target
    End With
    Set Create = NewBinding
End Function

Private Sub mTarget_Click()
    If DO_DEBUG Then Debug.Print "CommandBinding mTarget_Click()"
    This.Command.Execute
End Sub

' Cannot RaiseEvents at ICommand interface from the concrete class TestMsgBoxCommand etc.
'Private Sub mCommand_CanExecuteChanged()
'    If DO_DEBUG Then Debug.Print "CommandBinding cmd_CanExecuteChanged()"
'    mTarget.Enabled = mCommand.CanExecute
'End Sub

'Private Sub mCommand_Executed()
'    If DO_DEBUG Then Debug.Print "CommandBinding cmd_Executed()"
'End Sub

Private Sub ICommandBinding_EvaluateCanExecute(ByVal Context As Object)
    EvaluateCanExecute Context
End Sub

Private Sub EvaluateCanExecute(ByVal Context As Object)
    mTarget.Enabled = This.Command.CanExecute
End Sub
