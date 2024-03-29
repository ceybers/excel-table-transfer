Attribute VB_Name = "ArraySearch"
'@Folder "Helpers.Array"
Option Explicit

'@IgnoreModule ProcedureNotUsed

' Source: https://en.wikipedia.org/wiki/Binary_search_algorithm#Procedure

' Searches for SearchItem in the sorted array ArrayToSearch.
' Returns the index if SearchItem is found.
' Returns -1 if it is not in the list.
Public Function BinarySearch(ByRef ArrayToSearch As Variant, ByVal SearchItem As Variant) As Long
    Debug.Assert IsArray(ArrayToSearch)
    
    Dim LeftIndex As Long
    LeftIndex = LBound(ArrayToSearch)
    
    Dim RightIndex As Long
    RightIndex = UBound(ArrayToSearch)
    
    Do While LeftIndex <= RightIndex
        Dim MiddleIndex As Long
        MiddleIndex = (LeftIndex + RightIndex) / 2
        
        If ArrayToSearch(MiddleIndex) < SearchItem Then
            LeftIndex = MiddleIndex + 1
            
        ElseIf ArrayToSearch(MiddleIndex) > SearchItem Then
            RightIndex = MiddleIndex - 1
        Else
            BinarySearch = MiddleIndex
            Exit Function
        End If
    Loop
    
    BinarySearch = -1
End Function

' Searchs in a array of shape (1 to n, 1 to m), only testing (i, 1)
Public Function BinarySearch2(ByRef ArrayToSearch As Variant, ByVal SearchItem As Variant) As Long
    Debug.Assert IsArray(ArrayToSearch)
    
    Dim LeftIndex As Long
    LeftIndex = LBound(ArrayToSearch)
    
    Dim RightIndex As Long
    RightIndex = UBound(ArrayToSearch)
    
    Do While LeftIndex <= RightIndex
        Dim MiddleIndex As Long
        MiddleIndex = (LeftIndex + RightIndex) / 2
        
        If ArrayToSearch(MiddleIndex, 1) < SearchItem Then
            LeftIndex = MiddleIndex + 1
            
        ElseIf ArrayToSearch(MiddleIndex, 1) > SearchItem Then
            RightIndex = MiddleIndex - 1
        Else
            BinarySearch2 = MiddleIndex
            Exit Function
        End If
    Loop
    
    BinarySearch2 = -1
End Function

