VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyQualityViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.ViewModels")
Option Explicit

Private Type TState
    KeyColumn As KeyColumn
End Type
Private This As TState

Public Property Get KeyColumn() As KeyColumn
    Set KeyColumn = This.KeyColumn
End Property

Public Sub Load(ByVal KeyColumn As KeyColumn)
    Set This.KeyColumn = KeyColumn
End Sub
