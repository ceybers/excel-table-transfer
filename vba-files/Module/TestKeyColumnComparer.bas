Attribute VB_Name = "TestKeyColumnComparer"
'@Folder "MVVM.Models.KeyColumn"
Option Explicit
Option Private Module

Public Sub TestCompareKeyColumns()
    Dim compare As KeyColumnComparer
    Set compare = KeyColumnComparer.Create(GetLHS, GetRHS)
    
    compare.SrcKeyColumn.DebugPrint
    
    Debug.Print "TEST"
    Debug.Print "===="
    Debug.Print "IsSubsetLHS = " & compare.IsSubsetLHS
    Debug.Print "IsSubsetRHS = " & compare.IsSubsetRHS
    Debug.Print "IsMatch = " & compare.IsMatch
    Debug.Print "LHSOnly = " & compare.SrcSetOnly.Count
    Debug.Print "RHSOnly = " & compare.DstSetOnly.Count
    Debug.Print "Intersection = " & compare.IntersectionSet.Count
    Debug.Print vbNullString
    
    Dim mapResult As Variant
    mapResult = compare.Map
    SubPasteMap mapResult
    Debug.Print "mapped"
End Sub

Private Function GetLHS() As KeyColumn
    Set GetLHS = KeyColumn.FromRange(ThisWorkbook.Worksheets(2).Range("A2:A5,A14"), False)
End Function

Private Function GetRHS() As KeyColumn
    Set GetRHS = KeyColumn.FromRange(ThisWorkbook.Worksheets(2).Range("C2:C13"))
End Function

Private Sub PrintCollection(ByVal coll As Collection)
    Dim V As Variant
    For Each V In coll
        Debug.Print CStr(V)
    Next V
End Sub

Private Sub SubPasteMap(ByVal Map As Variant)
    Dim rng As Range
    Set rng = ThisWorkbook.Worksheets(2).ListObjects(2).ListColumns(2).DataBodyRange
    Dim arr As Variant
    arr = rng.Value2
    Dim i As Integer
    For i = LBound(arr, 1) To UBound(arr, 1)
        arr(i, 1) = Map(i + 1)
    Next i
    rng.Value2 = arr
End Sub

