VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Model.KeyColumn"
Option Explicit

Private Type TState
    ListObject As ListObject
    KeyColumns As Collection
    Selected As KeyColumn
End Type
Private This As TState

Public Property Get ListObject() As ListObject
    Set ListObject = This.ListObject
End Property

Public Property Get KeyColumns() As Collection
    Set KeyColumns = This.KeyColumns
End Property

Public Property Get Selected() As KeyColumn
    Set Selected = This.Selected
End Property

Private Sub Class_Initialize()
    Set This.KeyColumns = New Collection
End Sub

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    GenerateKeyColumns
End Sub

Private Sub GenerateKeyColumns()
    Dim ListColumn As ListColumn
    For Each ListColumn In This.ListObject.ListColumns
        Dim NewKeyColumn As KeyColumn
        Set NewKeyColumn = KeyColumn.FromColumn(ListColumn)
        This.KeyColumns.Add Key:=ListColumn.Name, Item:=NewKeyColumn
    Next ListColumn
End Sub

Public Sub TrySelect(ByVal ColumnName As String)
    Set This.Selected = This.KeyColumns.Item(ColumnName)
End Sub
