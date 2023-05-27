VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder MVVM.Infrastructure.Abstract
Option Explicit

Public Property Get Caption() As String
End Property

Public Function CanExecute() As Boolean
End Function

Public Sub Execute()
End Sub
