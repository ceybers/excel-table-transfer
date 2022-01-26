Attribute VB_Name = "TestArrayToFilteredRange"
'@Folder("TableTransferTool")
Option Explicit

Private Sub Test()
    Dim lo As ListObject
    Dim lc As ListColumn
    Dim rng As Range
    Dim fltRng As Range
    
    Set lo = ThisWorkbook.Worksheets(1).ListObjects(2)
    Set lc = lo.ListColumns(4)
    Set rng = lc.DataBodyRange
    Set fltRng = rng.SpecialCells(xlCellTypeVisible)
    
    'Debug.Print "rng:"
    'PrintRange rng
    
    'Debug.Print "fltRng:"
    'PrintRange fltRng
    
    Dim mask As Variant
    mask = rng.Value2
    SetArrayElementsToEmpty mask
    mask(4, 1) = "dd"
    mask(6, 1) = "ff"
    mask(7, 1) = "gg"
    
    'Debug.Print "mask:"
    'PrintArray mask
    
    'DoWork rng, mask
End Sub

Private Sub PrintRange(ByVal rng As Range)
    PrintArray rng.Value2
End Sub

Private Sub PrintArray(ByVal arr As Variant)
    Dim i As Long
    If VarType(arr) = vbArray + vbVariant Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            Debug.Print i & ": " & CStr(arr(i, 1)) & " (" & VarType(arr(i, 1)) & ")"
        Next i
    Else
        Debug.Print "0" & ": " & CStr(arr) & " (" & VarType(arr) & ")"
    End If
End Sub

Private Sub SetArrayElementsToEmpty(ByRef arr As Variant)
    Dim i As Long
    If VarType(arr) = vbArray + vbVariant Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            arr(i, 1) = vbEmpty
        Next i
    Else
        arr = vbEmpty
    End If
End Sub

Public Sub ArrayToFilteredRange(ByVal rng As Range, ByVal arr As Variant)
    Dim fltRng As Range
    Dim area As Range
    Dim v As Variant
    
    Dim fst As Long
    Dim top As Long
    Dim bot As Long
    Dim hei As Long
    
    On Error Resume Next
    Set fltRng = rng.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If fltRng Is Nothing Then Exit Sub
    
    If rng.Columns.Count <> 1 Then
        Err.Raise vbObjectError + 2, , "DisjointRangeToArray only works with range with a column count of 1"
    End If
    
    Dim i As Long
    
    fst = rng.Rows(1).Row
    
    For Each area In fltRng.Areas
        'Debug.Print area.Address & " of " & rng.Address
        top = area.Rows(1).Row
        bot = area.Rows(area.Rows.Count).Row
        hei = bot - top + 1
        'Debug.Print top & " to " & bot & " (" & hei & ")"
        
        v = area.Value2
        
        If hei = 1 Then
            'Debug.Print "0# " & v & " <-- " & (1 + top - fst) & "# " & arr(1 + top - fst, 1)
            v = arr(1 + top - fst, 1)
        Else
            For i = 1 To hei
                'Debug.Print i & "# " & v(i, 1) & " <-- " & (i + top - fst) & "# " & arr(i + top - fst, 1)
                v(i, 1) = arr(i + top - fst, 1)
            Next
        End If
        
        'If VarType(v) = vbArray + vbVariant Then Debug.Print LBound(v, 1) & ", " & UBound(v, 1)
        area.Value2 = v
    Next area
End Sub
