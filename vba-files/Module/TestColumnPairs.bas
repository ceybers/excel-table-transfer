Attribute VB_Name = "TestColumnPairs"
'@Folder "MVVM.Models.ColumnPairs"
Option Explicit
Option Private Module

Public Sub Test()
    Dim lhs As ListObject
    Dim RHS As ListObject
    
    Set lhs = ThisWorkbook.Worksheets(1).ListObjects(1)
    Set RHS = ThisWorkbook.Worksheets(1).ListObjects(2)
    
    Dim colPairs As ColumnPairs
    Set colPairs = New ColumnPairs
    
    Dim ColPair As ColumnPair
    
    
    Set ColPair = ColumnPair.Create(lhs.ListColumns(2), RHS.ListColumns(2))
    colPairs.Add ColPair
    
    Set ColPair = ColumnPair.Create(lhs.ListColumns(3), RHS.ListColumns(4))
    colPairs.Add ColPair
    
    Set ColPair = ColumnPair.Create(lhs.ListColumns(4), RHS.ListColumns(3))
    colPairs.Add ColPair
    
    'PrintColumnPairs colPairs
    Dim Result As Variant
    
    Set Result = colPairs.GetPair(Dst:=RHS.ListColumns(1))
    If Result Is Nothing Then
        Debug.Print "Not found"
    Else
        Debug.Print Result.ToString
    End If
    
    Set ColPair = ColumnPair.Create(lhs.ListColumns(1), RHS.ListColumns(2))
    colPairs.Add ColPair
    'PrintColumnPairs colPairs
    
    Debug.Print "OrReplace"
    colPairs.AddOrReplace ColPair
    'PrintColumnPairs colPairs
End Sub

