Attribute VB_Name = "ArrayHelpers"
'@Folder "Common.Helpers.Array"
Option Explicit

Public Sub ArrayToFilteredRange(ByVal rng As Range, ByVal arr As Variant)
    Dim fltRng As Range
    Dim Area As Range
    Dim V As Variant
    
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
    
    fst = rng.rows(1).row
    
    For Each Area In fltRng.Areas
        'Debug.Print area.Address & " of " & rng.Address
        top = Area.rows(1).row
        bot = Area.rows(Area.rows.Count).row
        hei = bot - top + 1
        'Debug.Print top & " to " & bot & " (" & hei & ")"
        
        V = Area.Value2
        
        If hei = 1 Then
            'Debug.Print "0# " & v & " <-- " & (1 + top - fst) & "# " & arr(1 + top - fst, 1)
            V = arr(1 + top - fst, 1)
        Else
            For i = 1 To hei
                'Debug.Print i & "# " & v(i, 1) & " <-- " & (i + top - fst) & "# " & arr(i + top - fst, 1)
                V(i, 1) = arr(i + top - fst, 1)
            Next
        End If
        
        'If VarType(v) = vbArray + vbVariant Then Debug.Print LBound(v, 1) & ", " & UBound(v, 1)
        Area.Value2 = V
    Next Area
End Sub