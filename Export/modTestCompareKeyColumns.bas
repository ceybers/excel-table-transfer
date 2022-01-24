Attribute VB_Name = "modTestCompareKeyColumns"
'@Folder("KeyColumn")
Option Explicit

Public Sub TestCompareKeyColumns()
    Dim compare As clsCompareKeyColumns
    Set compare = clsCompareKeyColumns.Create(GetLHS, GetRHS)
    
    compare.LHS.PrintKeys
    
    Debug.Print "TEST"
    Debug.Print "===="
    Debug.Print "IsSubsetLHS = " & compare.IsSubsetLHS
    Debug.Print "IsSubsetRHS = " & compare.IsSubsetRHS
    Debug.Print "IsMatch = " & compare.IsMatch
    Debug.Print "LHSOnly = " & compare.LeftOnly.Count
    Debug.Print "RHSOnly = " & compare.RightOnly.Count
    Debug.Print "Intersection = " & compare.Intersection.Count
    Debug.Print ""
    
    Dim mapResult As Variant
    mapResult = compare.Map
    SubPasteMap mapResult
    Debug.Print "mapped"
End Sub

Private Function GetLHS() As clsKeyColumn
    Set GetLHS = clsKeyColumn.FromRange(ThisWorkbook.Worksheets(2).Range("A2:A5,A14"), False)
End Function

Private Function GetRHS() As clsKeyColumn
    Set GetRHS = clsKeyColumn.FromRange(ThisWorkbook.Worksheets(2).Range("C2:C13"))
End Function

Private Sub PrintCollection(ByRef coll As Collection)
    Dim v As Variant
    For Each v In coll
        Debug.Print CStr(v)
    Next v
End Sub

Private Sub SubPasteMap(ByRef Map As Variant)
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

