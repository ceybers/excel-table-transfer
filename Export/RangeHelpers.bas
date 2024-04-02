Attribute VB_Name = "RangeHelpers"
'@IgnoreModule ProcedureNotUsed
'@Folder "Helpers.Objects"
Option Explicit

'@Description "Adds a Range to an existing Range as a Union. If the existing Range object is blank, sets it to the appended Range."
'@Ignore FunctionReturnValueAlwaysDiscarded
Public Function AppendRange(ByVal RangeToAppend As Range, ByRef UnionRange As Range) As Boolean
Attribute AppendRange.VB_Description = "Adds a Range to an existing Range as a Union. If the existing Range object is blank, sets it to the appended Range."
    If RangeToAppend Is Nothing Then Exit Function
    
    If UnionRange Is Nothing Then
        Set UnionRange = RangeToAppend
        AppendRange = True
        Exit Function
    End If
    
    If Not RangeToAppend.Parent Is UnionRange.Parent Then Exit Function
    
    Set UnionRange = Application.Union(UnionRange, RangeToAppend)
    AppendRange = True
End Function

'@Description "Returns True if SpecialCells would have returned a Range. Returns False if no cells were selected."
Public Function HasSpecialCells(ByVal Range As Range, ByVal CellType As XlCellType, _
    Optional ByVal Value As XlSpecialCellsValue) As Boolean
Attribute HasSpecialCells.VB_Description = "Returns True if SpecialCells would have returned a Range. Returns False if no cells were selected."
    If Range Is Nothing Then Exit Function

    Dim Result As Range
    On Error Resume Next
    If Value = 0 Then
        Set Result = Range.SpecialCells(CellType)
    Else
        Set Result = Range.SpecialCells(CellType, Value)
    End If
    
    HasSpecialCells = (Not Result Is Nothing)
End Function
