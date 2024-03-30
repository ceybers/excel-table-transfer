Attribute VB_Name = "TestDebugTransferInstruction"
'@Folder("MVVM2.AppContext")
Option Explicit

Public Function GetDebugTransfer() As TransferInstruction2
    Dim Result As TransferInstruction2
    Set Result = New TransferInstruction2
    
    With Result
        Set .Source.Table = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
        Set .Destination.Table = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(2)
        
        .Source.KeyColumnName = "KeyA"
        .Destination.KeyColumnName = "KeyB"
    End With
    
    Set GetDebugTransfer = Result
End Function

