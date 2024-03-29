Attribute VB_Name = "TestTransferInstruction"
'@Folder "Tests.AppContext"
Option Explicit

'Private Const TRANSFER_SERIALIZED_OBJECT_ROW_COUNT As Long = 8

'@EntryPoint
Public Sub TestAutoTransfer()
    Dim Transfer As TransferInstruction
    Set Transfer = GetTestTransferInstruction
    
    PrintTransferInstruction Transfer
    
    'SaveTransferInstruction Transfer
    'Exit Sub
    
    'Dim loadedTI As Collection
    'Set loadedTI = DeserializeTransferInstructions(ThisWorkbook.Worksheets.Item(4))
    'Debug.Print loadedTI.Count
    
    'Debug.Print "Deserialized:"
    'Debug.Print "#1"
    'PrintTransferInstruction loadedTI(1)
    'Debug.Print "#2"
    'PrintTransferInstruction loadedTI(2)
    
    'loadedTI(2).Transfer
    
    'TransferHistorySerializer.SaveTransferInstructionsFromWorksheet loadedTI, ThisWorkbook.Worksheets.Item(4)
    
    'Exit Sub
    
    With Transfer.Destination
        .ListColumns.Item(2).DataBodyRange.Clear
        .ListColumns.Item(3).DataBodyRange.Clear
        .ListColumns.Item(4).DataBodyRange.Clear
    End With
    
    '@Ignore FunctionReturnValueDiscarded
    Transfer.Transfer
End Sub

Private Sub PrintTransferInstruction(ByVal Transfer As TransferInstruction)
    Dim i As Long
    
    Debug.Print "TRANSFER"
    Debug.Print " SRC," & Transfer.Source.Range.Address(external:=True)
    Debug.Print " SRCKEY," & Transfer.SourceKey.Name
    Debug.Print " DST," & Transfer.Destination.Range.Address(external:=True)
    Debug.Print " DSTKEY," & Transfer.DestinationKey.Name
    Debug.Print " FLAGS," & Transfer.Flags
    Debug.Print " VALUES," & Transfer.ValuePairs.Count
    For i = 1 To Transfer.ValuePairs.Count
        Debug.Print "  " & Transfer.ValuePairs.Item(i).ToString
    Next i
    Debug.Print "END"
End Sub

Public Function GetTestTransferInstruction() As TransferInstruction
    Dim LHS As ListObject
    Dim RHS As ListObject
    
    Set GetTestTransferInstruction = New TransferInstruction
    
    Set LHS = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(1)
    Set RHS = ThisWorkbook.Worksheets.Item(1).ListObjects.Item(2)
    
    With GetTestTransferInstruction
        Set .Source = LHS
        Set .Destination = RHS
        
        Set .SourceKey = LHS.ListColumns.Item(1)
        Set .DestinationKey = RHS.ListColumns.Item(1)
            
        .Flags = AddFlag(.Flags, ClearDestinationFirst)
        .Flags = AddFlag(.Flags, DestinationFilteredOnly)
        .Flags = AddFlag(.Flags, HighlightMapped)
    
        .ValuePairs.Add ColumnPair.Create(LHS.ListColumns.Item(2), RHS.ListColumns.Item(4))
        .ValuePairs.Add ColumnPair.Create(LHS.ListColumns.Item(3), RHS.ListColumns.Item(2))
        .ValuePairs.Add ColumnPair.Create(LHS.ListColumns.Item(4), RHS.ListColumns.Item(3))
    End With
End Function

