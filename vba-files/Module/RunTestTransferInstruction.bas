Attribute VB_Name = "RunTestTransferInstruction"
'@IgnoreModule ImplicitActiveSheetReference, ImplicitActiveWorkbookReference
'@Folder("TransferEngine")
Option Explicit

'@EntryPoint "DoTestTransferInstruction"
Public Sub DoTestTransferInstruction()
    Worksheets(1).Activate
    Range("A2").Activate
    
    Dim ThisTransfer As TransferInstruction
    Set ThisTransfer = New TransferInstruction
    With ThisTransfer
        Set .SourceKey = Worksheets(1).ListObjects(1).ListColumns(1)
        Set .Source = .SourceKey.Parent
        Set .DestinationKey = Worksheets(2).ListObjects(1).ListColumns(1)
        Set .Destination = .DestinationKey.Parent
        Set .ValuePairs = New Collection
        .ValuePairs.Add ColumnTuple.Create( _
            Worksheets(1).ListObjects(1).ListColumns(2), _
            Worksheets(2).ListObjects(1).ListColumns(2))
        .ValuePairs.Add ColumnTuple.Create( _
            Worksheets(1).ListObjects(1).ListColumns(3), _
            Worksheets(2).ListObjects(1).ListColumns(3))
        .RHStoLHSRowMap = KeyColumnMapper.Create(.SourceKey, .DestinationKey).GenerateMap
    End With
    
    Debug.Print ThisTransfer.IsValid
    
    Debug.Print ThisTransfer.ToString
    
    Worksheets(2).Activate
    Range("A2").Activate
    
    ThisTransfer.Transfer
End Sub


