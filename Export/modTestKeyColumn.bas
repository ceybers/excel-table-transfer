Attribute VB_Name = "modTestKeyColumn"
'@Folder "KeyColumn"
Option Explicit

Public Sub TestKeyColumn()
    Dim ws As Worksheet
    Dim rng As Range
    Dim key As clsKeyColumn
    
    Set ws = ThisWorkbook.Worksheets(2)
    Set rng = ws.Range("A2:A5,A14") ',C2:C13")
    Set key = clsKeyColumn.FromRange(rng, True)
    
    Debug.Print "TEST"
    Debug.Print "===="
    Debug.Print "Distinct = " & key.Count
    Debug.Print "Unique = " & key.UniqueKeys.Count
    Debug.Print "IsDistinct = " & key.IsDistinct
    Debug.Print "Errors = " & key.ErrorCount
    Debug.Print "Blanks = " & key.BlankCount
    Debug.Print "Find 'def' = " & key.Find("def")
    Debug.Print "Find '1234567890' = " & key.Find("1234567890")
    Debug.Print "Find 'Right Only2' = " & key.Find("Right Only2")
    Debug.Print ""
End Sub

Public Sub TestKeyColumnPerf()
    Dim ws As Worksheet
    Dim rng As Range
    Dim key As clsKeyColumn
    Dim arr As Variant
    Dim findThis As String
    
    Set ws = ThisWorkbook.Worksheets(3)
    Set rng = ws.ListObjects(1).ListColumns(1).DataBodyRange
    Set key = clsKeyColumn.FromRange(rng, True)
    findThis = rng.Cells(500, 1).Value2
    arr = rng.Value2
    
    Dim i As Long
    Dim start As Double
    start = Timer
    
    For i = 1 To 100
        key.Find findThis
        'FindInArray arr, findThis
    Next i
    
    Debug.Print Timer - start
    
    ' 10000 6.69140625   0.62890625   0.59375
    '  1000 0.63671875   0.0625       0.0625
    '   100 0.05859375   0.00390625   0.0078125
End Sub

Private Function FindInArray(ByRef arr As Variant, ByVal value As String) As Long
    Dim v As Variant
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        If arr(i, 1) = value Then
            FindInArray = i
            Exit Function
        End If
    Next i
End Function
