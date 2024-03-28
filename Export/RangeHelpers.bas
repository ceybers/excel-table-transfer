Attribute VB_Name = "RangeHelpers"
'@Folder "Helpers.Objects"
Option Explicit

Public Function AppendRange(ByVal RangeToAppend As Range, ByRef UnionRange As Range) As Boolean
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

