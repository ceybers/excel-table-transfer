VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommitFullColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Model2.Commiter")
Option Explicit
Implements ICommitStrategy

Private Sub ICommitStrategy_Commit(ByVal Destination As TransferTable, ByVal Deltas As Variant)
    Dim i As Long
    For i = 1 To Destination.ValueColumnCount
        CommitColumn Destination.ValueColumns(i - 1), Deltas
    Next i
End Sub

Private Sub CommitColumn(ByVal ListColumn As ListColumn, ByVal Deltas As Variant)
    Dim ValueVariant As Variant
    ValueVariant = ListColumn.DataBodyRange.Value2
    
    Dim i As Long
    For i = 0 To UBound(Deltas)
        If Deltas(i)(DeltaIndex.tdCol) = ListColumn.Index Then
            ValueVariant(Deltas(i)(DeltaIndex.tdRow), 1) = Deltas(i)(DeltaIndex.tdValueAfter)
        End If
    Next i
    
    ListColumn.DataBodyRange.Value2 = ValueVariant
End Sub

