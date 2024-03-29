Attribute VB_Name = "SortedIndexedKeyColumn"
'@Folder "Model.TransferInstruction2"
Option Explicit


Public Function GetSortedIndexedKeyColumn(ByVal KeyColumn As ListColumn) As Variant
    Dim ValueVariant As Variant
    ValueVariant = KeyColumn.DataBodyRange.Value2
    
    ' Expand the array from (1 to n, 1 to 1) to (1 to n, 1 to 3)
    ReDim Preserve ValueVariant(LBound(ValueVariant) To UBound(ValueVariant), 1 To 3)
    
    ' Store the row index (relative to .Value2 variant) in (i, 2)
    Dim i As Long
    For i = LBound(ValueVariant) To UBound(ValueVariant)
        ValueVariant(i, 2) = i
    Next i
    
    ' Sort on (i, 1), retaining original row index in  (i, 2)
    QuickSort2 ValueVariant
    
    GetSortedIndexedKeyColumn = ValueVariant
End Function

