VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Model.TransferInstruction"
Option Explicit

Private Const MAX_PRINT_DELTAS As Long = 10

Private Type TState
    Source As TransferTable
    Destination As TransferTable
    Comparer As IValueComparer

    KeyMap As Variant ' Variant(1 to n, 1 to 3)
    ' (i, 1) is key, (i, 2) is LHS index (rel to .Value2), (i, 3) is RHS index

    ValueDeltas As Variant
    ValueDeltaCursor As Long
    
    TransferDeltas As TransferDeltas
End Type
Private This As TState

Public Property Get Source() As TransferTable
    Set Source = This.Source
End Property

Public Property Get Destination() As TransferTable
    Set Destination = This.Destination
End Property

Public Property Get TransferDeltas() As TransferDeltas
    Set TransferDeltas = This.TransferDeltas
End Property

Public Property Get Comparer() As IValueComparer
    Set Comparer = This.Comparer
End Property

Public Property Set Comparer(ByVal vNewValue As IValueComparer)
    Set This.Comparer = vNewValue
End Property

Public Property Get Deltas() As Variant
    Deltas = This.ValueDeltas
End Property

Private Sub Class_Initialize()
    Set This.Source = New TransferTable
    Set This.Destination = New TransferTable
    Set This.TransferDeltas = New TransferDeltas
End Sub

Public Sub SetComparer(ByVal Comparer As IValueComparer)
    Set This.Comparer = Comparer
End Sub

Public Sub Evaluate()
    Set This.Comparer = New NonBlankComparer ' TODO This needs to be here, since we aren't assigning it somewhere else yet
    
    If Not IsValid Then
        Err.Raise vbObjectError + 15, ERR_SOURCE, "Could not load Transfer Instruction" ' TODO ERR MSG Constants
        Exit Sub
    End If
    
    GenerateKeyMap
    If IsEmpty(This.KeyMap) Then Exit Sub
    
    TransferColumns
    UpdateTransferDeltas
End Sub

Private Function IsValid() As Boolean
    If Source Is Nothing Then Exit Function
    If Destination Is Nothing Then Exit Function
    
    IsValid = True
End Function

Private Sub GenerateKeyMap()
    Dim SourceKeyListColumn As ListColumn
    If Not Source.TryGetKeyListColumn(SourceKeyListColumn) Then Exit Sub
    
    Dim DestinationKeyListColumn As ListColumn
    If Not Destination.TryGetKeyListColumn(DestinationKeyListColumn) Then Exit Sub
    
    This.KeyMap = MapKeyColumns(SourceKeyListColumn, DestinationKeyListColumn)
End Sub

Private Sub TransferColumns()
    ReDim This.ValueDeltas(0 To (UBound(This.KeyMap) * Source.ValueColumnCount))
    
    This.ValueDeltaCursor = 0
    
    Dim i As Long
    For i = 1 To Source.ValueColumnCount
        TransferColumn i - 1
    Next i
    
    If This.ValueDeltaCursor > 0 Then
        ReDim Preserve This.ValueDeltas(0 To This.ValueDeltaCursor - 1)
    Else
        ReDim This.ValueDeltas(0 To 0)
    End If
End Sub

Private Sub TransferColumn(ByVal FieldIndex As Long)
    Dim SourceColumn As ListColumn
    Set SourceColumn = Source.ValueColumns(FieldIndex)
    
    Dim DestinationColumn As ListColumn
    Set DestinationColumn = Destination.ValueColumns(FieldIndex)

    Dim SourceVariantArray As Variant
    SourceVariantArray = SourceColumn.DataBodyRange.Value2
    
    Dim DestinationVariantArray As Variant
    DestinationVariantArray = DestinationColumn.DataBodyRange.Value2
    
    Dim i As Long
    For i = 1 To UBound(This.KeyMap)
        Dim SourceValue As Variant
        SourceValue = SourceVariantArray(This.KeyMap(i, 2), 1)
        
        Dim DestinationValue As Variant
        DestinationValue = DestinationVariantArray(This.KeyMap(i, 3), 1)
        
        If This.Comparer.Compare(SourceValue, DestinationValue) Then
            Dim CellDelta As Variant
            CellDelta = GetTransferCellDelta( _
                This.KeyMap(i, 3), _
                Destination.ValueColumns(FieldIndex).Index, _
                This.KeyMap(i, 1), _
                SourceColumn.Name, DestinationColumn.Name, _
                SourceValue, DestinationValue)
                
            This.ValueDeltas(This.ValueDeltaCursor) = CellDelta
            This.ValueDeltaCursor = This.ValueDeltaCursor + 1
        End If
    Next i
End Sub

Public Sub Commit(ByVal Commiter As ICommitStrategy)
    If IsEmpty(This.ValueDeltas) Then Exit Sub
    
    Commiter.Commit Destination, This.ValueDeltas
End Sub

Private Sub UpdateTransferDeltas()
    Set This.TransferDeltas = New TransferDeltas
    
    This.TransferDeltas.Load This.ValueDeltas
    
    Dim i As Long
    For i = LBound(This.KeyMap) To UBound(This.KeyMap)
        This.TransferDeltas.Keys.Add This.KeyMap(i, 1)
    Next i
    
    For i = 0 To (Destination.ValueColumnCount - 1)
        This.TransferDeltas.Fields.Add Destination.ValueColumns(i)
    Next i
End Sub

Public Sub PostProcess(ByVal PostProcessor As ITransferPostProcesser)
    PostProcessor.Execute Me
End Sub
