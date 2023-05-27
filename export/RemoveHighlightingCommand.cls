VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoveHighlightingCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.TableProps.Commands"
Option Explicit
Implements ICommand

Private mView As IView
Private mAppContext As IAppContext
Private mViewModel As TableHighlightingVM

Public Property Get Context() As IAppContext
    Set Context = mAppContext
End Property

Public Property Set Context(ByVal vNewValue As IAppContext)
    Set mAppContext = vNewValue
End Property

Public Property Get View() As IView
    Set View = mView
End Property

Public Property Set View(ByVal vNewValue As IView)
    Set mView = vNewValue
End Property

Public Property Get ViewModel() As IViewModel
    Set ViewModel = mViewModel
End Property

Public Property Set ViewModel(ByVal vNewValue As IViewModel)
    Set mViewModel = vNewValue
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Remove Highlighting"
End Property

Public Function Create(ByVal Context As IAppContext, ByVal View As IView, ByVal ViewModel As IViewModel) _
    As RemoveHighlightingCommand
    Dim Result As RemoveHighlightingCommand
    Set Result = New RemoveHighlightingCommand
    With Result
        Set .Context = Context
        Set .View = View
        Set .ViewModel = ViewModel
    End With
    Set Create = Result
End Function

Private Function ICommand_CanExecute() As Boolean
    ICommand_CanExecute = False 'True
End Function

Private Sub ICommand_Execute()
    mViewModel.DoRemoveHighlighting
End Sub
