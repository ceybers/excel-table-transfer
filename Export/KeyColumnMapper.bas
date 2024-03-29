Attribute VB_Name = "KeyColumnMapper"
'@Folder "Model.TransferInstruction2"
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

