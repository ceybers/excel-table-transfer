VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Model.TransferInstruction2"
Option Explicit

Private Type TState
    Table As ListObject
    KeyColumnName As String
    ValueColumnNames As Variant ' String()
    ValueColumns As Variant
End Type

Private This As TState

'Public Property Get Table() As ListObject
'    Set Table = This.Table
'End Property

'Public Property Set Table(ByVal RHS As ListObject)
'    Set This.Table = RHS
'End Property

'Public Property Get KeyColumn() As ListColumn
'    Set KeyColumn = This.Table.ListColumns.Item(This.KeyColumnName)
'End Property

'Public Property Get KeyColumnName() As String
'    KeyColumnName = This.KeyColumnName
'End Property

'Public Property Let KeyColumnName(ByVal RHS As String)
'    This.KeyColumnName = RHS
'End Property

Public Property Get ValueColumns() As Variant
    ValueColumns = This.ValueColumns
End Property

'Public Property Get ValueColumnNames() As Variant
'    ValueColumnNames = This.ValueColumnNames
'End Property

'Public Property Let ValueColumnNames(ByVal RHS As Variant)
'    This.ValueColumnNames = RHS
'End Property

Public Property Get ValueColumnCount() As Long
    ValueColumnCount = UBound(This.ValueColumns) - LBound(This.ValueColumns) + 1
End Property

Public Sub Load(ByVal Table As ListObject, ByVal KeyColumnName As String, ByVal ValueColumnNames As Variant)
    Set This.Table = Table
    This.KeyColumnName = KeyColumnName
    This.ValueColumnNames = ValueColumnNames
    
    '@Ignore FunctionReturnValueDiscarded
    TryGetValueListColumns This.ValueColumns
End Sub

Public Function IsValid() As Boolean
    If This.Table Is Nothing Then Exit Function
    Debug.Assert Not This.Table.ListColumns.Item(This.KeyColumnName) Is Nothing
    
    IsValid = True
End Function

Public Function TryGetKeyListColumn(ByRef OutListColumn As ListColumn) As Boolean
    If IsValid Then
        Set OutListColumn = This.Table.ListColumns.Item(This.KeyColumnName)
        TryGetKeyListColumn = True
    End If
End Function

'@Ignore FunctionReturnValueAlwaysDiscarded
Public Function TryGetValueListColumns(ByRef OutListColumns As Variant) As Boolean
    If Not IsValid Then
        Exit Function
    End If
    
    Dim Result() As Variant
    ReDim Result(0 To UBound(This.ValueColumnNames))
    
    Dim i As Long
    For i = 0 To UBound(Result)
        Set Result(i) = This.Table.ListColumns.Item(This.ValueColumnNames(i))
    Next i
    
    OutListColumns = Result
    
    TryGetValueListColumns = True
End Function

