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
    
    ' TODO CHK if `Worksheet Is Worksheet` is a valid equality test
    If Not RangeToAppend.parent Is UnionRange.parent Then Exit Function
    
    Set UnionRange = Application.Union(UnionRange, RangeToAppend)
    AppendRange = True
End Function
