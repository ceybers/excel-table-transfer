VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyColumnProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.TableProps.Model"
Option Explicit

Private Const NOT_TEXT_SUFFIX As String = " (non-text)"
Private Const NOT_UNIQUE_SUFFIX As String = " (not unique)"
Private Const UNDEFINED_COLUMN As String = "(No key column set)"

Private Type TState
    ListObject As ListObject
    Item As Scripting.Dictionary
End Type
Private This As TState

Public Property Get Name() As String
    Name = This.ListObject.Name
End Property

Private Sub Class_Initialize()
    Set This.Item = New Scripting.Dictionary
End Sub

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    LoadCollection
End Sub

Private Sub LoadCollection()
    This.Item.Add Key:=UNDEFINED_COLUMN, Item:=UNDEFINED_COLUMN
    
    Dim ListColumn As ListColumn
    For Each ListColumn In This.ListObject.ListColumns
        Dim ColumnProperty As ColumnProperty
        Set ColumnProperty = New ColumnProperty
        ColumnProperty.Load ListColumn
        This.Item.Add Key:=ListColumn.Name, Item:=GetNameWithSuffix(ColumnProperty)
    Next ListColumn
End Sub

Private Function GetNameWithSuffix(ByVal ColumnProperty As ColumnProperty) As String
    GetNameWithSuffix = ColumnProperty.Name
    
    If ColumnProperty.ColumnType <> "Text" Then
        GetNameWithSuffix = ColumnProperty.Name & NOT_TEXT_SUFFIX
        Exit Function
    End If
    
    If ColumnProperty.Unique <> All Then
        GetNameWithSuffix = ColumnProperty.Name & NOT_UNIQUE_SUFFIX
        Exit Function
    End If
End Function

Public Function GetDictionary() As Scripting.Dictionary
    Set GetDictionary = This.Item
End Function
