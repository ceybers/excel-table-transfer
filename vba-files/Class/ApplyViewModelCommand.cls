VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplyViewModelCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TestTableProperties")
Option Explicit

Private Type TState
    ViewModel As TablePropViewModel
    CanExecute As Boolean
End Type
Private This As TState

Public Function CanExecute() As Boolean
    CanExecute = This.ViewModel.IsDirty
End Function

Public Sub Execute()
    This.ViewModel.Apply
End Sub

Public Sub Create(ByVal ViewModel As TablePropViewModel)
    Set This.ViewModel = ViewModel
End Sub

Public Sub UpdateCommandButton(ByVal CommandButton As CommandButton)
    CommandButton.Enabled = CanExecute
End Sub

