Attribute VB_Name = "TestColumnPairs"
'@Folder "Tests.Model"
Option Explicit
Option Private Module

Public Sub TestColumnPairs()
    Dim LHS As ListObject
    Dim RHS As ListObject
    
    Set LHS = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    Set RHS = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(2)
    
    Dim colPairs As ColumnPairs
    Set colPairs = New ColumnPairs
    
    Dim colPair As ColumnPair
    
    
    Set colPair = ColumnPair.Create(LHS.ListColumns.Item(2), RHS.ListColumns.Item(2))
    colPairs.Add colPair
    
    Set colPair = ColumnPair.Create(LHS.ListColumns.Item(3), RHS.ListColumns.Item(4))
    colPairs.Add colPair
    
    Set colPair = ColumnPair.Create(LHS.ListColumns.Item(4), RHS.ListColumns.Item(3))
    colPairs.Add colPair
    
    'PrintColumnPairs colPairs
    Dim Result As Variant
    
    Set Result = colPairs.GetPair(RHS:=RHS.ListColumns.Item(1))
    If Result Is Nothing Then
        Debug.Print "Not found"
    Else
        Debug.Print Result.ToString
    End If
    
    Set colPair = ColumnPair.Create(LHS.ListColumns.Item(1), RHS.ListColumns.Item(2))
    colPairs.Add colPair
    'PrintColumnPairs colPairs
    
    Debug.Print "OrReplace"
    colPairs.AddOrReplace colPair
    'PrintColumnPairs colPairs
End Sub

