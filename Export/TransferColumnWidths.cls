VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferColumnWidths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model2.PostProcessors")
Option Explicit
Implements ITransferPostProcesser

Private Sub ITransferPostProcesser_Execute(ByVal TransferInstruction As TransferInstruction2)
    Dim i As Long
    For i = 0 To TransferInstruction.Source.ValueColumnCount - 1
        Dim SourceColumn As ListColumn
        Set SourceColumn = TransferInstruction.Source.ValueColumns(i)
        
        Dim DestinationColumn As ListColumn
        Set DestinationColumn = TransferInstruction.Destination.ValueColumns(i)
        
        DestinationColumn.DataBodyRange.ColumnWidth = SourceColumn.DataBodyRange.ColumnWidth
    Next i
End Sub

