VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferNumberFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Model.PostProcessors.Strategies"
Option Explicit
Implements ITransferPostProcesser

Private Sub ITransferPostProcesser_Execute(ByVal TransferInstruction As TransferInstruction)
    Dim i As Long
    For i = 0 To TransferInstruction.Source.ValueColumnCount - 1
        With TransferInstruction
            UpdateColumn .Source.ValueColumns(i), .Destination.ValueColumns(i)
        End With
    Next i
End Sub

Private Sub UpdateColumn(ByVal SourceColumn As ListColumn, ByVal DestinationColumn As ListColumn)
    DestinationColumn.DataBodyRange.NumberFormat = SourceColumn.DataBodyRange.Cells.Item(1, 1).NumberFormat
End Sub
