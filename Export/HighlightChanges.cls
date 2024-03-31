VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HighlightChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.Model.PostProcessors.Strategies"
Option Explicit
Implements ITransferPostProcesser

Private Type TState
    ChangeType As Long
    Color As Long
End Type
Private This As TState

Public Property Get ChangeType() As Long
    ChangeType = This.ChangeType
End Property

Public Property Let ChangeType(ByVal vNewValue As Long)
    This.ChangeType = vNewValue
End Property

Public Property Get Color() As Long
    Color = This.Color
End Property

Public Property Let Color(ByVal vNewValue As Long)
    This.Color = vNewValue
End Property

Private Sub ITransferPostProcesser_Execute(ByVal TransferInstruction As TransferInstruction)
    Dim Target As Range
    
    Dim DestinationRange As Range
    Set DestinationRange = TransferInstruction.Destination.Table.DataBodyRange
    
    Dim i As Long
    For i = 0 To UBound(TransferInstruction.Deltas)
        If TransferInstruction.Deltas(i)(DeltaIndex.tdChangeType) = This.ChangeType Then
            Dim Row As Long
            Row = TransferInstruction.Deltas(i)(DeltaIndex.tdRow)
            
            Dim Column As Long
            Column = TransferInstruction.Deltas(i)(DeltaIndex.tdCol)
            
            AppendRange RangeToAppend:=DestinationRange.Cells.Item(Row, Column), UnionRange:=Target
        End If
    Next i
    
    If Target Is Nothing Then Exit Sub
    
    ApplyHighlighting Target, This.Color
End Sub

Public Function Create(ByVal ChangeType As Long, ByVal Color As Long) As HighlightChanges
    Dim Result As HighlightChanges
    Set Result = New HighlightChanges
    
    With Result
        .ChangeType = ChangeType
        .Color = Color
    End With
    
    Set Create = Result
End Function
