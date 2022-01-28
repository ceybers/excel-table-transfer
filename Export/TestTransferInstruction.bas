Attribute VB_Name = "TestTransferInstruction"
'@Folder("TableTransfer")
Option Explicit

Private Const TRANSFER_SERIALIZED_OBJECT_ROW_COUNT As Integer = 8

Public Sub TestTransfer()
    Dim Transfer As TransferInstruction
    Set Transfer = GetTestTransferInstruction
    
    'SaveTransferInstruction Transfer
    'Exit Sub
    
    Dim loadedTI As Collection
    'Set loadedTI = DeserializeTransferInstructions(ThisWorkbook.Worksheets(4))
    Exit Sub
    Debug.Print loadedTI.Count
    
    Debug.Print "Deserialized:"
    Debug.Print "#1"
    'PrintTransferInstruction loadedTI(1)
    Debug.Print "#2"
    'PrintTransferInstruction loadedTI(2)
    
    'loadedTI(2).Transfer
    
    TransferHistorySerializer.SaveTransferInstructionsFromWorksheet loadedTI, ThisWorkbook.Worksheets(4)
    
    Exit Sub
    
    With Transfer.Destination
        .ListColumns(2).DataBodyRange.Clear
        .ListColumns(3).DataBodyRange.Clear
        .ListColumns(4).DataBodyRange.Clear
    End With
    
    Transfer.Transfer
End Sub

Private Function PrintTransferInstruction(ByVal Transfer As TransferInstruction)
    Dim i As Long
    
    Debug.Print "TRANSFER"
    Debug.Print " SRC," & Transfer.Source.Range.Address(external:=True)
    Debug.Print " SRCKEY," & Transfer.SourceKey.Name
    Debug.Print " DST," & Transfer.Destination.Range.Address(external:=True)
    Debug.Print " DSTKEY," & Transfer.DestinationKey.Name
    Debug.Print " FLAGS," & Transfer.Flags
    Debug.Print " VALUES," & Transfer.ValuePairs.Count
    For i = 1 To Transfer.ValuePairs.Count
         Debug.Print "  " & Transfer.ValuePairs(i).ToString
    Next i
    Debug.Print "END"
End Function

Private Function GetTestTransferInstruction() As TransferInstruction
    Dim lhs As ListObject
    Dim rhs As ListObject
    
    Set GetTestTransferInstruction = New TransferInstruction
    
    Set lhs = ThisWorkbook.Worksheets(1).ListObjects(1)
    Set rhs = ThisWorkbook.Worksheets(1).ListObjects(2)
    
    With GetTestTransferInstruction
        Set .Source = lhs
        Set .Destination = rhs
        
        Set .SourceKey = lhs.ListColumns(1)
        Set .DestinationKey = rhs.ListColumns(1)
            
        .Flags = AddFlag(.Flags, ClearDestinationFirst)
        .Flags = AddFlag(.Flags, DestinationFilteredOnly)
    
        .ValuePairs.Add ColumnPair.Create(lhs.ListColumns(2), rhs.ListColumns(2))
        .ValuePairs.Add ColumnPair.Create(lhs.ListColumns(3), rhs.ListColumns(3))
        .ValuePairs.Add ColumnPair.Create(lhs.ListColumns(4), rhs.ListColumns(4))
    End With
End Function







