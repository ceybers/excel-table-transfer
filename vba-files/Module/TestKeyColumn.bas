Attribute VB_Name = "TestKeyColumn"
'@Folder "MVVM.Models.KeyColumn"
Option Explicit
Option Private Module

Public Sub Test()
    Dim ws As Worksheet
    Dim rng As Range
    Dim Key As KeyColumn
    
    Set ws = ThisWorkbook.Worksheets(2)
    Set rng = ws.Range("A2:A5,A14") ',C2:C13")
    Set Key = KeyColumn.FromRange(rng, True, True)
    
    Debug.Print "TEST"
    Debug.Print "===="
    Debug.Print "Distinct = " & Key.Count
    Debug.Print "Unique = " & Key.UniqueKeys.Count
    Debug.Print "IsDistinct = " & Key.IsDistinct
    Debug.Print "Errors = " & Key.ErrorCount
    Debug.Print "Blanks = " & Key.BlankCount
    Debug.Print "Find 'def' = " & Key.Find("def")
    Debug.Print "Find '1234567890' = " & Key.Find("1234567890")
    Debug.Print "Find 'Right Only2' = " & Key.Find("Right Only2")
    Debug.Print vbNullString
End Sub

Public Sub TestPerformance()
    Dim ws As Worksheet
    Dim rng As Range
    Dim Key As KeyColumn
    Dim arr As Variant
    Dim findThis As String
    
    Set ws = ThisWorkbook.Worksheets(3)
    Set rng = ws.ListObjects(1).ListColumns(1).DataBodyRange
    Set Key = KeyColumn.FromRange(rng, True)
    findThis = rng.Cells(500, 1).Value2
    arr = rng.Value2
    
    Dim i As Long
    Dim start As Double
    start = Timer
    
    For i = 1 To 100
        Key.Find findThis
        'FindInArray arr, findThis
    Next i
    
    Debug.Print Timer - start
    
    ' 10000 6.69140625   0.62890625   0.59375
    '  1000 0.63671875   0.0625       0.0625
    '   100 0.05859375   0.00390625   0.0078125
End Sub

Private Function FindInArray(ByVal arr As Variant, ByVal Value As String) As Long
    Dim V As Variant
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        If arr(i, 1) = Value Then
            FindInArray = i
            Exit Function
        End If
    Next i
End Function
