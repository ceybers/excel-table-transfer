VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferColumnWidths"
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
            UpdateColumn .Source.ValueColumnCount(i), .Destination.ValueColumnCount(i)
        End With
    Next i
End Sub

Private Sub UpdateColumn(ByVal SourceColumn As ListColumn, ByVal DestinationColumn As ListColumn)
    DestinationColumn.DataBodyRange.ColumnWidth = SourceColumn.DataBodyRange.ColumnWidth
End Sub
