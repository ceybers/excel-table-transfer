VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CancelViewCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.Example2.Commands"
Option Explicit
Implements ICommand

Private mView As ICancellable
Private mAppContext As IAppContext
Private mViewModel As IViewModel

Public Property Get Context() As IAppContext
    Set Context = mAppContext
End Property

Public Property Set Context(ByVal vNewValue As IAppContext)
    Set mAppContext = vNewValue
End Property

Public Property Get view() As IView
    Set view = mView
End Property

Public Property Set view(ByVal vNewValue As IView)
    Set mView = vNewValue
End Property

Public Property Get ViewModel() As IViewModel
    Set ViewModel = mViewModel
End Property

Public Property Set ViewModel(ByVal vNewValue As IViewModel)
    Set mViewModel = vNewValue
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Cancel"
End Property

Public Function Create(ByVal Context As IAppContext, ByVal view As IView, ByVal ViewModel As IViewModel) _
    As CancelViewCommand
    Dim Result As CancelViewCommand
    Set Result = New CancelViewCommand
    With Result
        Set .Context = Context
        Set .view = view
        Set .ViewModel = ViewModel
    End With
    Set Create = Result
End Function

Private Function ICommand_CanExecute() As Boolean
    ICommand_CanExecute = True
End Function

Private Sub ICommand_Execute()
    Debug.Print "CancelViewCommand executing"
    mView.OnCancel
End Sub