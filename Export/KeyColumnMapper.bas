Attribute VB_Name = "KeyColumnMapper"
'@Folder "Model2.TransferInstruction2"
Option Explicit

Public Function MapKeyColumns(ByVal LHS As ListColumn, ByVal RHS As ListColumn) As Variant
    Dim Result As Variant
    
    Dim SourceKeyArray As Variant
    SourceKeyArray = GetSortedIndexedKeyColumn(LHS)
    
    Dim DestinationKeyArray As Variant
    DestinationKeyArray = GetSortedIndexedKeyColumn(RHS)
    
    Dim MatchCount As Long
    
    Dim i As Long
    For i = 1 To UBound(SourceKeyArray, 1)
        Dim SearchResult As Long
        SearchResult = BinarySearch2(DestinationKeyArray, SourceKeyArray(i, 1))
        
        If SearchResult > -1 Then
            SourceKeyArray(i, 3) = DestinationKeyArray(SearchResult, 2)
            MatchCount = MatchCount + 1
        End If
    Next i
    
    If MatchCount = 0 Then
        MapKeyColumns = Empty
        Exit Function
    End If
    
    ReDim Result(1 To MatchCount, 1 To 3)
    
    Dim Cursor As Long
    Cursor = 1
    
    For i = 1 To UBound(SourceKeyArray, 1)
        If Not (IsEmpty(SourceKeyArray(i, 3))) Then
            Result(Cursor, 1) = SourceKeyArray(i, 1)
            Result(Cursor, 2) = SourceKeyArray(i, 2)
            Result(Cursor, 3) = SourceKeyArray(i, 3)
            Cursor = Cursor + 1
        End If
    Next i
    
    MapKeyColumns = Result
End Function

Public Function GetSortedIndexedKeyColumn(ByVal KeyColumn As ListColumn) As Variant
    Dim ValueVariant As Variant
    ValueVariant = KeyColumn.DataBodyRange.Value2
    
    ' Expand the array from (1 to n, 1 to 1) to (1 to n, 1 to 3)
    ReDim Preserve ValueVariant(LBound(ValueVariant) To UBound(ValueVariant), 1 To 3)
    
    ' Store the row index (relative to .Value2 variant) in (i, 2)
    Dim i As Long
    For i = LBound(ValueVariant) To UBound(ValueVariant)
        ValueVariant(i, 1) = CStr(ValueVariant(i, 1)) ' Cast errors to strings
        ValueVariant(i, 2) = i
    Next i
    
    ' Sort on (i, 1), retaining original row index in  (i, 2)
    QuickSort2 ValueVariant
    
    GetSortedIndexedKeyColumn = ValueVariant
End Function
