Attribute VB_Name = "TestDebugTransferInstruction"
'@Folder "Tests.MVVM"
Option Explicit

Public Function GetDebugTransfer() As TransferInstruction
    Dim Result As TransferInstruction
    Set Result = New TransferInstruction
    
    With Result
        Set .Source.Table = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
        Set .Destination.Table = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(2)
        
        .Source.KeyColumnName = "KeyA"
        .Destination.KeyColumnName = "KeyB"
    End With
    
    Set GetDebugTransfer = Result
End Function

