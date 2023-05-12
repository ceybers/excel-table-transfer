VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablePropViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event PropertyChanged(ByVal PropertyName As String)

Private Type TState
    ListObject As ListObject
    ColumnProperties As ColumnProperties
End Type
Private This As TState

Public Property Get ColumnProperties() As ColumnProperties
    Set ColumnProperties = This.ColumnProperties
End Property

Public Property Get TableName() As String
    TableName = This.ListObject.Name
End Property

Public Property Get WorkSheetName() As String
    WorkSheetName = This.ListObject.Parent.Name
End Property

Public Property Get WorkBookName() As String
    WorkBookName = This.ListObject.Parent.Parent.Name
End Property

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    Set This.ColumnProperties = New ColumnProperties
    This.ColumnProperties.Load ListObject
End Sub

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    Debug.Print "PropertyChanged('"; PropertyName; "')"
    RaiseEvent PropertyChanged(PropertyName)
End Sub

Public Sub DoActiveListObject()
    This.ListObject.Parent.Activate
End Sub
