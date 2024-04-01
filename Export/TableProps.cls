VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Model.TableProps"
Option Explicit

Private Type TState
    ListObject As Object
    AvailableTableNode As AvailableTableNode
End Type
Private This As TState

Public Property Get IsValid() As Boolean
    IsValid = Not ListObject Is Nothing
End Property

Public Property Get TableName() As String
    TableName = This.ListObject.Name
End Property

Public Property Get WorkbookName() As String
    WorkbookName = This.ListObject.Parent.Parent.Name
End Property

Public Property Get ListObject() As ListObject
    Set ListObject = This.ListObject
End Property

Public Property Get Node() As AvailableTableNode
    Set Node = This.AvailableTableNode
End Property

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
End Sub
