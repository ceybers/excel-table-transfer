Attribute VB_Name = "modArrayExTEST"
'@Folder "ArrayExtensions"
Option Explicit

Private Sub TESTArrayEx()
    Dim LHS As Variant
    Dim RHS As Variant
    
    LHS = ThisWorkbook.Worksheets("Sheet1").ListObjects("Table1").ListColumns("KeyA").DataBodyRange.value
    RHS = ThisWorkbook.Worksheets("Sheet1").ListObjects("Table1").ListColumns("KeyA").DataBodyRange.value
    'rhs = ThisWorkbook.Worksheets("Sheet1").listobjects("Table2").ListColumns("KeyB").DataBodyRange.Value
    
    Debug.Print "***"
    Debug.Print "START TESTING"
    'Debug.Print "Match Arrays = " & ArraySubset(lhs, rhs)
    
    Dim antiTest As Variant: antiTest = ArrayAntiJoinLeft(LHS, RHS)
    Dim distTest As Variant: distTest = ArrayDistinct(LHS)
    Dim findTest As Integer: findTest = ArrayFind("left1", LHS)
    Dim interTest As Variant: interTest = ArrayIntersect(LHS, RHS)
    Dim lenTest As Integer: lenTest = ArrayLength(LHS)
    Dim matchTest As Boolean: matchTest = ArrayMatch(LHS, RHS)
    Dim subsetTest As Boolean: subsetTest = ArraySubset(LHS, RHS)
    Dim trimTest As Variant: trimTest = ArrayTrim(LHS, 2)
    Dim uniqTest As Variant: uniqTest = ArrayUnique(LHS)
    Dim fltTxtTest As Variant: fltTxtTest = ArrayFilterTextOnly(LHS)
    
    ArrayPrint antiTest, "ANTI JOIN LEFT:"
    ArrayPrint distTest, "DISTINCT:"
    Debug.Print "Find Test = " & findTest
    ArrayPrint interTest, "INTERSECT:"
    Debug.Print "Length Test = " & lenTest
    Debug.Print "MatchTest = " & matchTest
    Debug.Print "SubsetTest = " & subsetTest
    ArrayPrint trimTest, "TRIM"
    ArrayPrint uniqTest, "UNIQUE"
    ArrayPrint fltTxtTest, "TEXT ONLY"
    
    Dim one As ArrayExAnalyseOne
    Dim two As ArrayExAnalyseTwo
    
    one = ArrayAnalyseOne(LHS)
    two = ArrayAnalyseTwo(LHS, RHS)
    
    'ArrayPrint (two.LeftOnly)
    
    Debug.Print "STOP TESTING"
End Sub

Private Function ArrayPrint(arr As Variant, Optional header As String)
    If Not (IsMissing(header)) Then
        Debug.Print header
    End If
    Dim i As Integer
    If IsEmpty(arr) Then
        Debug.Print " >>> Variant/Empty"
        Debug.Print vbNullString
        Exit Function
    End If
    Debug.Print "Printing array(1 to " & UBound(arr, 1) & ", 1 to 1)"
    For i = 1 To UBound(arr, 1)
        If IsError(arr(i, 1)) Then
            Debug.Print " " & CStr(i) & ") #ERR"
        Else
            Debug.Print " " & CStr(i) & ") " & arr(i, 1)
        End If
    Next i
    Debug.Print vbNullString
End Function
