Attribute VB_Name = "ArraySort"
'@Folder "Helpers.Array"
Option Explicit

'@IgnoreModule ProcedureNotUsed

' Source: https://wellsr.com/vba/2018/excel/vba-quicksort-macro-to-sort-arrays-fast

' Sorts an Array in-place using the Quick Sort algorithm. Array must be of shape (m to n)
Public Sub QuickSort(ByRef ArrayToSort As Variant)
    DoQuickSort ArrayToSort, LBound(ArrayToSort), UBound(ArrayToSort)
End Sub

Private Sub DoQuickSort(ByRef ArrayToSort As Variant, _
                        ByVal ArrayLowerBound As Long, _
                        ByVal ArrayUpperBound As Long)
        
    Dim LowIndex As Long
    LowIndex = ArrayLowerBound
    
    Dim HighIndex As Long
    HighIndex = ArrayUpperBound
    
    Dim PivotValue As Variant
    PivotValue = ArrayToSort((ArrayLowerBound + ArrayUpperBound) \ 2)
     
    Do While (LowIndex <= HighIndex)
        Do While (ArrayToSort(LowIndex) < PivotValue And LowIndex < ArrayUpperBound)
            LowIndex = LowIndex + 1
        Loop
      
        Do While (PivotValue < ArrayToSort(HighIndex) And HighIndex > ArrayLowerBound)
            HighIndex = HighIndex - 1
        Loop
     
        If (LowIndex <= HighIndex) Then
            Dim SwapValue As Variant
            SwapValue = ArrayToSort(LowIndex)
          
            ArrayToSort(LowIndex) = ArrayToSort(HighIndex)
            ArrayToSort(HighIndex) = SwapValue
          
            LowIndex = LowIndex + 1
            HighIndex = HighIndex - 1
        End If
    Loop
     
    If (ArrayLowerBound < HighIndex) Then DoQuickSort ArrayToSort, ArrayLowerBound, HighIndex
    If (LowIndex < ArrayUpperBound) Then DoQuickSort ArrayToSort, LowIndex, ArrayUpperBound
End Sub

' Sorts an array of shape (1 to n, 1 to 2), sorting on (i, 1)
Public Sub QuickSort2(ByRef ArrayToSort As Variant)
    DoQuickSort2 ArrayToSort, LBound(ArrayToSort), UBound(ArrayToSort)
End Sub

Private Sub DoQuickSort2(ByRef ArrayToSort As Variant, _
                        ByVal ArrayLowerBound As Long, _
                        ByVal ArrayUpperBound As Long)
        
    Dim LowIndex As Long
    LowIndex = ArrayLowerBound
    
    Dim HighIndex As Long
    HighIndex = ArrayUpperBound
    
    Dim PivotValue As Variant
    PivotValue = ArrayToSort((ArrayLowerBound + ArrayUpperBound) \ 2, 1)
     
    Do While (LowIndex <= HighIndex)
        Do While (ArrayToSort(LowIndex, 1) < PivotValue And LowIndex < ArrayUpperBound)
            LowIndex = LowIndex + 1
        Loop
      
        Do While (PivotValue < ArrayToSort(HighIndex, 1) And HighIndex > ArrayLowerBound)
            HighIndex = HighIndex - 1
        Loop
     
        If (LowIndex <= HighIndex) Then
            Dim SwapValue1 As Variant
            SwapValue1 = ArrayToSort(LowIndex, 1)
            Dim SwapValue2 As Variant
            SwapValue2 = ArrayToSort(LowIndex, 2)
          
            ArrayToSort(LowIndex, 1) = ArrayToSort(HighIndex, 1)
            ArrayToSort(LowIndex, 2) = ArrayToSort(HighIndex, 2)
            
            ArrayToSort(HighIndex, 1) = SwapValue1
            ArrayToSort(HighIndex, 2) = SwapValue2
          
            LowIndex = LowIndex + 1
            HighIndex = HighIndex - 1
        End If
    Loop
     
    If (ArrayLowerBound < HighIndex) Then DoQuickSort2 ArrayToSort, ArrayLowerBound, HighIndex
    If (LowIndex < ArrayUpperBound) Then DoQuickSort2 ArrayToSort, LowIndex, ArrayUpperBound
End Sub


