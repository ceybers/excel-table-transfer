Attribute VB_Name = "ListObjectHelpers"
'@Folder("HelperFunctions")
Option Explicit

Public Function GetListColumnFromRange(ByVal rng As Range) As ListColumn
    If rng Is Nothing Then Exit Function
    If rng.ListObject Is Nothing Then Exit Function
    If rng.Columns.Count <> 1 Then Exit Function
    
    Dim lo As ListObject
    Dim lc As ListColumn
    
    Set lo = rng.ListObject
    For Each lc In lo.ListColumns
        If lc.Range.Column = rng.Cells(1, 1).Column Then
            Set GetListColumnFromRange = lc
            Exit Function
        End If
    Next lc
End Function
