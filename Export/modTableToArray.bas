Attribute VB_Name = "modTableToArray"
'@Folder "HelperFunctions"
Option Explicit

Public Function GetVisibleTableAsArray(ByVal lo As ListObject) As Variant
    Dim arr As Variant, vismask As Variant
    arr = GetDBR(lo)
    vismask = GetVisibilityMask(lo)
    ApplyBitmask arr, vismask
    GetVisibleTableAsArray = arr
End Function

' If value of bitmask is not 1, then set cell value to empty variant
Private Function ApplyBitmask(ByVal arr As Variant, ByVal bitmask As Variant) As Boolean
    Dim myEmpty As Variant
    Dim i As Integer, j As Integer
    
    Debug.Assert LBound(arr, 1) = LBound(bitmask, 1)
    Debug.Assert UBound(arr, 1) = UBound(bitmask, 1)
    Debug.Assert LBound(arr, 2) = LBound(bitmask, 2)
    Debug.Assert UBound(arr, 2) = UBound(bitmask, 2)
    
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            If bitmask(i, j) <> 1 Then
                arr(i, j) = myEmpty
            End If
        Next j
    Next i
    
    ApplyBitmask = True
End Function

Private Function GetVisibilityMask(ByVal lo As ListObject) As Variant
    Dim bitmask As Variant
    Dim maskRng As Range
    Dim origin As Range
    
    bitmask = GetDBR(lo)
    ReDim bitmask(LBound(bitmask, 1) To UBound(bitmask, 1), LBound(bitmask, 2) To UBound(bitmask, 2))
    Set maskRng = lo.DataBodyRange.SpecialCells(xlCellTypeVisible)
    Set origin = lo.DataBodyRange.Cells(1, 1)
    
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

Private Function GetDBR(ByVal lo As ListObject) As Variant
    Debug.Assert Not lo Is Nothing
    Dim Result As Variant
    Result = lo.DataBodyRange.Value
    GetDBR = Result
End Function

Private Function SetDBR(ByVal lo As ListObject, ByVal arr As Variant) As Boolean
    Dim dbr As Range
    Set dbr = lo.DataBodyRange
    Dim arrHeight As Integer: arrHeight = UBound(arr, 1)
    Dim arrWidth As Integer: arrWidth = UBound(arr, 2)
    Debug.Assert dbr.rows.Count = arrHeight
    Debug.Assert dbr.Columns.Count = arrWidth
    
    dbr.Value2 = arr
    SetDBR = True
End Function

Private Sub FillTableWithAddresses(ByVal lo As ListObject)
    Dim c As Range
    For Each c In lo.DataBodyRange.Cells
        c.Value = CStr(c.Address)
    Next c
End Sub

