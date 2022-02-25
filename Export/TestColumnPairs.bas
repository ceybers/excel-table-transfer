Attribute VB_Name = "TestColumnPairs"
'@Folder "ColumnPairs"
Option Explicit
Option Private Module

Public Sub Test()
    Dim lhs As ListObject
    Dim RHS As ListObject
    
    Set lhs = ThisWorkbook.Worksheets(1).ListObjects(1)
    Set RHS = ThisWorkbook.Worksheets(1).ListObjects(2)
    
    Dim colPairs As ColumnPairs
    Set colPairs = New ColumnPairs
    
    Dim colPair As ColumnPair
    
    
    Set colPair = ColumnPair.Create(lhs.ListColumns(2), RHS.ListColumns(2))
    colPairs.Add colPair
    
    Set colPair = ColumnPair.Create(lhs.ListColumns(3), RHS.ListColumns(4))
    colPairs.Add colPair
    
    Set colPair = ColumnPair.Create(lhs.ListColumns(4), RHS.ListColumns(3))
    colPairs.Add colPair
    
    'PrintColumnPairs colPairs
    Dim result As Variant
    
    Set result = colPairs.GetPair(RHS:=RHS.ListColumns(1))
    If result Is Nothing Then
        Debug.Print "Not found"
    Else
        Debug.Print result.ToString
    End If
    
    Set colPair = ColumnPair.Create(lhs.ListColumns(1), RHS.ListColumns(2))
    colPairs.Add colPair
    'PrintColumnPairs colPairs
    
    Debug.Print "OrReplace"
    colPairs.AddOrReplace colPair
    'PrintColumnPairs colPairs
End Sub

