VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoMapColumnsCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.ValueMapper.Commands"
Option Explicit
Implements ICommand

Private mView As IView
Private mAppContext As IAppContext
Private mViewModel As ValueMapperViewModel

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
    ICommand_Caption = "Auto Map"
End Property

Public Function Create(ByVal Context As IAppContext, ByVal View As IView, ByVal ViewModel As IViewModel) _
    As AutoMapColumnsCommand
    Dim Result As AutoMapColumnsCommand
    Set Result = New AutoMapColumnsCommand
    With Result
        Set .Context = Context
        Set .View = View
        Set .ViewModel = ViewModel
    End With
    Set Create = Result
End Function

Private Function ICommand_CanExecute() As Boolean
    ICommand_CanExecute = mViewModel.CanAutoMap
End Function

Private Sub ICommand_Execute()
    mViewModel.DoAutoMap
End Sub




