Attribute VB_Name = "TestTransferInstruction"
'@Folder("TableTransfer")
Option Explicit

Private Const TRANSFER_SERIALIZED_OBJECT_ROW_COUNT As Integer = 8

Public Sub TestAutoTransfer()
    Dim transfer As TransferInstruction
    Set transfer = GetTestTransferInstruction
    
    PrintTransferInstruction transfer
    
    'SaveTransferInstruction Transfer
    'Exit Sub
    
    'Dim loadedTI As Collection
    'Set loadedTI = DeserializeTransferInstructions(ThisWorkbook.Worksheets(4))
    'Debug.Print loadedTI.Count
    
    'Debug.Print "Deserialized:"
    'Debug.Print "#1"
    'PrintTransferInstruction loadedTI(1)
    'Debug.Print "#2"
    'PrintTransferInstruction loadedTI(2)
    
    'loadedTI(2).Transfer
    
    'TransferHistorySerializer.SaveTransferInstructionsFromWorksheet loadedTI, ThisWorkbook.Worksheets(4)
    
    'Exit Sub
    
    With transfer.Destination
        .ListColumns(2).DataBodyRange.Clear
        .ListColumns(3).DataBodyRange.Clear
        .ListColumns(4).DataBodyRange.Clear
    End With
    
    transfer.transfer
End Sub

Private Sub PrintTransferInstruction(ByVal transfer As TransferInstruction)
    Dim i As Long
    
    Debug.Print "TRANSFER"
    Debug.Print " SRC," & transfer.Source.Range.Address(external:=True)
    Debug.Print " SRCKEY," & transfer.SourceKey.Name
    Debug.Print " DST," & transfer.Destination.Range.Address(external:=True)
    Debug.Print " DSTKEY," & transfer.DestinationKey.Name
    Debug.Print " FLAGS," & transfer.Flags
    Debug.Print " VALUES," & transfer.ValuePairs.Count
    For i = 1 To transfer.ValuePairs.Count
         Debug.Print "  " & transfer.ValuePairs(i).ToString
    Next i
    Debug.Print "END"
End Sub

Public Function GetTestTransferInstruction() As TransferInstruction
    Dim LHS As ListObject
    Dim RHS As ListObject
    
    Set GetTestTransferInstruction = New TransferInstruction
    
    Set LHS = ThisWorkbook.Worksheets(1).ListObjects(1)
    Set RHS = ThisWorkbook.Worksheets(1).ListObjects(2)
    
    With GetTestTransferInstruction
        Set .Source = LHS
        Set .Destination = RHS
        
        Set .SourceKey = LHS.ListColumns(1)
        Set .DestinationKey = RHS.ListColumns(1)
            
        .Flags = AddFlag(.Flags, ClearDestinationFirst)
        .Flags = AddFlag(.Flags, DestinationFilteredOnly)
        .Flags = AddFlag(.Flags, HighlightMapped)
    
        .ValuePairs.Add ColumnPair.Create(LHS.ListColumns(2), RHS.ListColumns(4))
        .ValuePairs.Add ColumnPair.Create(LHS.ListColumns(3), RHS.ListColumns(2))
        .ValuePairs.Add ColumnPair.Create(LHS.ListColumns(4), RHS.ListColumns(3))
    End With
End Function







