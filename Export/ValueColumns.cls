VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Model2.ValueColumns"
Option Explicit

Public AllowFormula As Boolean
Public AllowEmpty As Boolean
Public AllowError As Boolean

Private Type TState
    ListObject As ListObject
    ValueColumns As Collection
    Selected As ValueColumn
End Type
Private This As TState

Public Property Get ListObject() As ListObject
    Set ListObject = This.ListObject
End Property

Public Property Get ValueColumns() As Collection
    Set ValueColumns = This.ValueColumns
End Property

Public Property Get Selected() As ValueColumn
    Set Selected = This.Selected
End Property

Private Sub Class_Initialize()
    Set This.ValueColumns = New Collection
End Sub

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    GenerateValueColumns
End Sub

Public Sub Reset()
    Dim ValueColumn As ValueColumn
    For Each ValueColumn In This.ValueColumns
        ValueColumn.IsMapped = False
    Next ValueColumn
End Sub

Public Function ToArray() As Variant
    Dim Result() As Variant
    ReDim Result(0 To This.ValueColumns.Count - 1)
    
    Dim i As Long
    For i = 0 To UBound(Result)
        Result(i) = This.ValueColumns.Item(i + 1).Name
    Next i
    
    ToArray = Result
End Function

Private Sub GenerateValueColumns()
    Dim ListColumn As ListColumn
    For Each ListColumn In This.ListObject.ListColumns
        Dim ValueColumn As ValueColumn
        Set ValueColumn = New ValueColumn
        ValueColumn.Load ListColumn
        'With ValueColumn
        '    .Name = ListColumn.Name
        'End With
        ValueColumn.IsSelectable = IsValueColumnSelectable(ValueColumn)
        This.ValueColumns.Add Key:=ListColumn.Name, Item:=ValueColumn
    Next ListColumn
End Sub

Private Function IsValueColumnSelectable(ByVal ValueColumn As ValueColumn) As Boolean
    Dim Result As Boolean
    Result = True
    
    '@Ignore UnassignedVariableUsage
    If ValueColumn.IsKeyColumn Then Result = False
    If AllowError = False And ValueColumn.HasErrors Then Result = False
    If AllowEmpty = False And ValueColumn.IsEmpty Then Result = False
    If AllowFormula = False And ValueColumn.IsFormula Then Result = False
    
    IsValueColumnSelectable = Result
End Function

Public Sub TrySelect(ByVal ColumnName As String)
    Set This.Selected = This.ValueColumns.Item(ColumnName)
End Sub

Public Sub SetKeyColumn(ByVal ColumnName As String)
    This.ValueColumns.Item(ColumnName).IsKeyColumn = True
End Sub
