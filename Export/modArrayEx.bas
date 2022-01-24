Attribute VB_Name = "modArrayEx"
'@Folder "ArrayExtensions"
Option Explicit

Public Type ArrayExAnalyseOne
    Distinct As Integer
    Unique As Integer
    length As Integer
    Errors As Integer
    Blanks As Integer
End Type

Public Type ArrayExAnalyseTwo
    LeftOnly As Variant
    RightOnly As Variant
    Intersection As Variant
    Match As Boolean
    LeftSubset As Boolean
    RightSubset As Boolean
    LeftOnlyCount As Integer
    IntersectionCount As Integer
    RightOnlyCount As Integer
End Type

Public Function ArrayAnalyseOne(arr As Variant) As ArrayExAnalyseOne
    With ArrayAnalyseOne
        .Distinct = ArrayLength(ArrayDistinct(arr))
        .Unique = ArrayLength(ArrayUnique(arr))
        .length = ArrayLength(arr)
        .Errors = ArrayErrorCount(arr)
        .Blanks = ArrayBlankCount(arr)
    End With
End Function

Public Function ArrayAnalyseTwo(LHS As Variant, RHS As Variant) As ArrayExAnalyseTwo
    With ArrayAnalyseTwo
        .LeftOnly = ArrayDistinct(ArrayAntiJoinLeft(LHS, RHS))
        .RightOnly = ArrayDistinct(ArrayAntiJoinLeft(RHS, LHS))
        .Intersection = ArrayDistinct(ArrayIntersect(LHS, RHS))
        .Match = ArrayMatch(LHS, RHS)
        .LeftSubset = ArraySubset(LHS, RHS)
        .RightSubset = ArraySubset(LHS, RHS)
        .LeftOnlyCount = ArrayLength(.LeftOnly)
        .IntersectionCount = ArrayLength(.Intersection)
        .RightOnlyCount = ArrayLength(.RightOnly)
    End With
End Function

' AntiJoinLeft(lhs, rhs) - Returns items only in lhs
' Distinct(lhs)          - Returns with duplicates removed
' Find(v, lhs)           - Returns first index of v in lhs
' Intersect(lhs, rhs)    - Returns items in both lhs and rhs
' Length(lhs)            - Returns length of lhs
' Match(lhs, rhs)        - True is all items in both lhs and rhs exist in each other
' Subset(lhs, rhs)       - True if all items in lhs exist in rhs
' Trim(lhs, n)           - Returns first n items from lhs
' Unique(lhs)            - Returns items that appear in lhs exactly once
' ErrorCount
' BlankCount


Public Function ArrayErrorCount(arr As Variant) As Integer
    Dim i As Integer
    Dim c As Integer
    For i = LBound(arr, 1) To UBound(arr, 1)
        If IsError(arr(i, 1)) Then
            c = c + 1
        End If
    Next i
    ArrayErrorCount = c
End Function

Public Function ArrayBlankCount(arr As Variant) As Integer
    Dim i As Integer
    Dim c As Integer
    For i = LBound(arr, 1) To UBound(arr, 1)
        If IsError(arr(i, 1)) Then
        ElseIf arr(i, 1) = vbNullString Then
            c = c + 1
        End If
    Next i
    ArrayBlankCount = c
End Function

Public Function ArrayFilterTextOnly(arr As Variant) As Variant
    Dim i As Integer
    Dim n As Integer
    Dim result As Variant
    result = arr
    For i = LBound(arr, 1) To UBound(arr, 1)
        If TypeName(arr(i, 1)) = "String" Then
            n = n + 1
            result(n, 1) = arr(i, 1)
        End If
    Next i
    ArrayFilterTextOnly = ArrayTrim(result, n)
End Function

' Returns a unique copy of the array.
' Only items that appear exactly once are included.
' Duplicates (including first instance), blanks and errors are excluded.
Public Function ArrayUnique(arr As Variant) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim c As Integer
    Dim this As Variant
    Dim result As Variant
    ReDim result(1 To UBound(arr, 1), 1 To 1)
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        this = arr(i, 1)
        If IsError(this) Then
        ElseIf this = vbNullString Then
        Else
            c = 0
            For j = LBound(arr, 1) To UBound(arr, 1)
                If IsError(arr(j, 1)) Then
                ElseIf this = arr(j, 1) Then
                    c = c + 1
                End If
            Next
            If c = 1 Then
                n = n + 1
                result(n, 1) = this
            End If
        End If
    Next i
    
    ArrayUnique = ArrayTrim(result, n)
End Function
' Returns a distinct copy of the array.
' In the case of duplicate values, only one instance is returned.
' Blanks and errors are excluded.
Public Function ArrayDistinct(arr As Variant) As Variant
    Dim i As Integer
    Dim n As Integer
    Dim this As Variant
    Dim result As Variant
    
    If IsEmpty(arr) Then
        ReDim result(0 To 0, 1 To 1)
        ArrayDistinct = result
        Exit Function
    End If
    
    If VarType(arr) < vbArray Then
        Debug.Assert False
    End If
    
    If UBound(arr, 1) = 0 Then
        ArrayDistinct = arr
        Exit Function
    End If
    
    ReDim result(1 To UBound(arr, 1), 1 To 1)
    
    n = 0
    For i = 1 To UBound(arr, 1)
        this = arr(i, 1)
        
        If IsError(this) Then
        ElseIf this = vbNullString Then
        ElseIf ArrayFind(this, result) <> -1 Then
        Else
            n = n + 1
            result(n, 1) = this
        End If
    Next i
    
    ArrayDistinct = ArrayTrim(result, n)
End Function

' Returns a list of items in both lhs and rhs.
' Excludes blanks and errors.
' Only checks 1st instance of each duplicate.
Public Function ArrayIntersect(LHS As Variant, RHS As Variant) As Variant
    Dim i As Integer
    Dim n As Integer
    Dim this As Variant
    Dim result As Variant
    ' Intersect size is always less or equal to min(lhs, rhs)
    ReDim result(1 To UBound(LHS, 1), 1 To 1)
    
    n = 0
    For i = 1 To UBound(LHS, 1)
        this = LHS(i, 1)
        
        If IsError(this) Then
        ElseIf this = vbNullString Then
        ElseIf ArrayFind(this, result) <> -1 Then
        ElseIf ArrayFind(this, RHS) < 0 Then
        Else
            n = n + 1
            result(n, 1) = this
        End If
    Next i
    
    ArrayIntersect = ArrayTrim(result, n)
End Function

' Returns the length of the first dimension of an array
' i.e. the number of rows in an array that was created
' from a single column (nx1) range.
Public Function ArrayLength(arr As Variant) As Integer
    ArrayLength = UBound(arr, 1) - LBound(arr, 1) + 1
    If UBound(arr, 1) = 0 Then ArrayLength = 0
End Function

' Returns a copy of the array, retaining only the first n items.
' If the length is longer than the provided array,
' the provided array is returned (length is ignored)
' If the length is zero or negative, the provided array is returned.
Public Function ArrayTrim(arr As Variant, length As Integer) As Variant
    Dim i As Integer
    Dim result() As Variant
    
    If length <= 0 Then
        ReDim result(1 To 1, LBound(arr, 2) To UBound(arr, 2))
        Exit Function
    End If
    
    If length > ArrayLength(arr) Then
        ArrayTrim = arr
        Exit Function
    End If
    
    ReDim result(1 To length, 1 To 1) As Variant
    
    For i = 1 To length
        result(i, 1) = arr(i, 1)
    Next i
    ArrayTrim = result
End Function

' Returns an array(n, 1) of all the items that are in the lhs array but
' not in the rhs array. Excludes blanks and errors.
Public Function ArrayAntiJoinLeft(LHS As Variant, RHS As Variant) As Variant
    Dim i As Integer
    Dim n As Integer
    Dim this As Variant
    Dim result As Variant
    
    'result = lhs
    ReDim result(1 To UBound(LHS, 1), 1 To 1)
    n = 0
    
    For i = LBound(LHS, 1) To UBound(LHS, 1)
        this = LHS(i, 1)
        If IsError(this) Then
        ElseIf this = vbNullString Then
        ElseIf ArrayFind(this, RHS) = -1 Then
            n = n + 1
            result(n, 1) = this
        End If
    Next i
    
    If n = 0 Then
        'result = ArrayTrim(result, 1)
        'result(1, 1) = ""
        ReDim result(0 To 0, 1 To 1)
    Else
        result = ArrayTrim(result, n)
    End If
    
    ArrayAntiJoinLeft = result
End Function

' Check if every item in lhs exists in rhs.
' Does not check if rhs items all exist in lhs.
' Ignores blanks and errors
Public Function ArraySubset(LHS As Variant, RHS As Variant) As Boolean
    Dim i As Integer
    Dim this As Variant
    
    If UBound(LHS, 1) <> UBound(RHS, 1) Then
        ArraySubset = False
        Exit Function
    End If
    
    For i = LBound(LHS, 1) To UBound(LHS, 1)
        this = LHS(i, 1)
        If IsError(this) Then
        ElseIf this = vbNullString Then
        ElseIf ArrayFind(this, RHS) = -1 Then
            ArraySubset = False
            Exit Function
        End If
    Next i
    ArraySubset = True
End Function

' Checks if all items exist in both lhs and rhs
Public Function ArrayMatch(LHS As Variant, RHS As Variant) As Boolean
    If ArraySubset(LHS, RHS) = False Then Exit Function
    If ArraySubset(RHS, LHS) = False Then Exit Function
    ArrayMatch = True
End Function

' Checks if the provided variant exists in the array.
' -1 means no match
' -2 means a blank (string "") was provided
' -3 means an error was provided
Public Function ArrayFind(Match As Variant, arr As Variant) As Integer
    Dim i As Integer
    Dim chk As Variant
    ArrayFind = -1
    If IsError(Match) Then
        ArrayFind = -3
        Exit Function
    End If
    If Match = vbNullString Then
        ArrayFind = -2
        Exit Function
    End If
    For i = LBound(arr, 1) To UBound(arr, 1)
        chk = arr(i, 1)
        If IsError(chk) Then
        ElseIf chk = vbNullString Then
        ElseIf Match = chk Then
            ArrayFind = i
            Exit Function
        End If
    Next i
End Function

' EOF
