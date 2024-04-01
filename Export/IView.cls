VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder "MVVM.Infrastructure.Abstract"
Option Explicit

'@Description "Gets the ViewModel / binding Source."
Public Property Get ViewModel() As Object
End Property

'@Description "Displays the View modally and returns a value indicating confirmation (True) or cancellation (False)."
Public Function ShowDialog() As TtViewResult
End Function

'@Description "Displays the View."
Public Sub Show()
End Sub

'@Description "Hides the View."
Public Sub Hide()
End Sub
