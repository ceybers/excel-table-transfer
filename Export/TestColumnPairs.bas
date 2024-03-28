Attribute VB_Name = "TestColumnPairs"
'@Folder "Tests.Model"
Option Explicit
Option Private Module

Public Sub TestColumnPairs()
    Dim lhs As ListObject
    Dim rhs As ListObject
    
    Set lhs = ThisWorkbook.Worksheets(1).ListObjects(1)
    Set rhs = ThisWorkbook.Worksheets(1).ListObjects(2)
    
    Dim colPairs As ColumnPairs
    Set colPairs = New ColumnPairs
    
    Dim colPair As ColumnPair
    
    
    Set colPair = ColumnPair.Create(lhs.ListColumns(2), rhs.ListColumns(2))
    colPairs.Add colPair
    
    Set colPair = ColumnPair.Create(lhs.ListColumns(3), rhs.ListColumns(4))
    colPairs.Add colPair
    
    Set colPair = ColumnPair.Create(lhs.ListColumns(4), rhs.ListColumns(3))
    colPairs.Add colPair
    
    'PrintColumnPairs colPairs
    Dim Result As Variant
    
    Set Result = colPairs.GetPair(rhs:=rhs.ListColumns(1))
    If Result Is Nothing Then
        Debug.Print "Not found"
    Else
        Debug.Print Result.ToString
    End If
    
    Set colPair = ColumnPair.Create(lhs.ListColumns(1), rhs.ListColumns(2))
    colPairs.Add colPair
    'PrintColumnPairs colPairs
    
    Debug.Print "OrReplace"
    colPairs.AddOrReplace colPair
    'PrintColumnPairs colPairs
End Sub

