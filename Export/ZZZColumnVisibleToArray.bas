Attribute VB_Name = "ZZZColumnVisibleToArray"
'@IgnoreModule
'@Folder "ZZZRefactor"
Option Explicit

'@Obsolete "ZZZ"
Public Function VisibleRangeToArray(ByVal rng As Range) As Variant
    Debug.Assert False
    
    'RangeToArray = rng.Value
    
    Dim arr As Variant
    Dim vis As Variant
    arr = rng.Value
    vis = GetVisibilityMask(rng)
    ApplyBitmask arr, vis
    'arr = modArrayEx.ArrayDistinct(arr)
    'Debug.Assert False
    
    VisibleRangeToArray = arr
End Function

Private Function GetVisibilityMask(ByVal rng As Range) As Variant
    Dim bitmask As Variant
    Dim maskRng As Range
    Dim origin As Range
    
    bitmask = rng.Value
    'ReDim bitmask(LBound(bitmask, 1) To UBound(bitmask, 1), LBound(bitmask, 2) To UBound(bitmask, 2))
    ReDim bitmask(LBound(bitmask, 1) To UBound(bitmask, 1), 1 To 1)
    
    Set maskRng = rng.SpecialCells(xlCellTypeVisible)
    Set origin = rng.Cells(1, 1)
    
    Dim i As Integer, j As Integer, k As Integer
    Dim a As Range
    For i = 1 To maskRng.Areas.Count
        Set a = maskRng.Areas(i)
        For j = 1 To a.rows.Count
            For k = 1 To a.Columns.Count
                bitmask(a.row - origin.row + 0 + j, k) = 1
            Next k
        Next j
    Next i
    
    GetVisibilityMask = bitmask
End Function

Private Function ApplyBitmask(ByVal arr As Variant, ByVal bitmask As Variant) As Boolean
    Dim myEmpty As Variant
    Dim i As Integer, j As Integer
    
    Debug.Assert LBound(arr, 1) = LBound(bitmask, 1)
    Debug.Assert UBound(arr, 1) = UBound(bitmask, 1)
    Debug.Assert LBound(arr, 2) = LBound(bitmask, 2)
    Debug.Assert UBound(arr, 2) = UBound(bitmask, 2)
    
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            If bitmask(i, 1) <> 1 Then           ' bitmask is (n x 1), i.e. we only keep 1 column
                arr(i, j) = myEmpty              ' TODO Change to = Empty?
            End If
        Next j
    Next i
    
    ApplyBitmask = True
End Function

