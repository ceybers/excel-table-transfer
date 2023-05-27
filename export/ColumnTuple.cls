VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnTuple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("MVVM.ValueMapper.Model")
Option Explicit

Private Type TColumnTuple
    SrcListColumn As ListColumn
    DstListColumn As ListColumn
End Type
Private This As TColumnTuple

Public Property Get SourceListColumn() As ListColumn
    Debug.Assert Not Me Is ColumnTuple
    Set SourceListColumn = This.SrcListColumn
End Property

Public Property Set SourceListColumn(ByVal vNewValue As ListColumn)
    Debug.Assert Not Me Is ColumnTuple
    Set This.SrcListColumn = vNewValue
End Property

Public Property Get DestinationListColumn() As ListColumn
    Debug.Assert Not Me Is ColumnTuple
    Set DestinationListColumn = This.DstListColumn
End Property

Public Property Set DestinationListColumn(ByVal vNewValue As ListColumn)
    Debug.Assert Not Me Is ColumnTuple
    Set This.DstListColumn = vNewValue
End Property

Public Function IsValid() As Boolean
    If This.SrcListColumn Is Nothing Then Exit Function
    If This.DstListColumn Is Nothing Then Exit Function
    IsValid = True
End Function

Public Function Create(ByVal SrcListColumn As ListColumn, ByVal DstListColumn As ListColumn) As ColumnTuple
    Dim Result As ColumnTuple
    Set Result = New ColumnTuple
    With Result
        Set .SourceListColumn = SrcListColumn
        Set .DestinationListColumn = DstListColumn
    End With
    Set Create = Result
End Function

Public Function Equals(ByVal RHS As ColumnTuple) As Boolean
    If RHS Is Nothing Then Exit Function
    If Not This.SrcListColumn Is RHS.SourceListColumn Then Exit Function
    If Not This.DstListColumn Is RHS.DestinationListColumn Then Exit Function
    Equals = True
End Function

Public Function ToString() As String
    Debug.Assert Not SourceListColumn Is Nothing
    Debug.Assert Not DestinationListColumn Is Nothing
    ToString = SourceListColumn.Name & "," & DestinationListColumn.Name
End Function
