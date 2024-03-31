Attribute VB_Name = "ColumnPairToTransferInstruction"
'@Folder "MVVM.Model.TransferInstruction"
Option Explicit

Public Sub UpdateTransferInstruction(ByVal TransferInstruction As TransferInstruction, _
    ByVal ColPairCollection As Collection)
    With TransferInstruction.Source
        .Load .Table, .KeyColumnName, ColPairCollectionToArray(ColPairCollection, True)
    End With
    
    With TransferInstruction.Destination
        .Load .Table, .KeyColumnName, ColPairCollectionToArray(ColPairCollection, False)
    End With
End Sub

Private Function ColPairCollectionToArray(ByVal ColPairCollection As Collection, _
    ByVal LHS As Boolean) As Variant
    Dim Result() As String
    ReDim Result(0 To ColPairCollection.Count - 1)
    
    Dim i As Long
    For i = 0 To UBound(Result)
        If LHS Then
            Result(i) = ColPairCollection.Item(i + 1).LHS.Name
        Else
            Result(i) = ColPairCollection.Item(i + 1).RHS.Name
        End If
    Next i
    
    ColPairCollectionToArray = Result
End Function


