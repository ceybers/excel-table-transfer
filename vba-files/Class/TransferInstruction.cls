VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "TransferEngine"
Option Explicit

Private Type TState
    Source As ListObject
    Destination As ListObject
    SourceKey As ListColumn
    DestinationKey As ListColumn
    
    ValuePairs As Collection
    RHStoLHSRowMap As Variant
End Type
Private This As TState

Public Property Get Source() As ListObject
    Set Source = This.Source
End Property

Public Property Set Source(ByVal vNewValue As ListObject)
    Set This.Source = vNewValue
End Property

Public Property Get Destination() As ListObject
    Set Destination = This.Destination
End Property

Public Property Set Destination(ByVal vNewValue As ListObject)
    Set This.Destination = vNewValue
End Property

Public Property Get SourceKey() As ListColumn
    Set SourceKey = This.SourceKey
End Property

Public Property Set SourceKey(ByVal vNewValue As ListColumn)
    Set This.SourceKey = vNewValue
End Property

Public Property Get DestinationKey() As ListColumn
    Set DestinationKey = This.DestinationKey
End Property

Public Property Set DestinationKey(ByVal vNewValue As ListColumn)
    Set This.DestinationKey = vNewValue
End Property

Public Property Get ValuePairs() As Collection
    Set ValuePairs = This.ValuePairs
End Property

Public Property Set ValuePairs(ByVal vNewValue As Collection)
    Set This.ValuePairs = vNewValue
End Property

Public Property Get RHStoLHSRowMap() As Variant
    RHStoLHSRowMap = This.RHStoLHSRowMap
End Property

Public Property Let RHStoLHSRowMap(ByVal vNewValue As Variant)
    This.RHStoLHSRowMap = vNewValue
End Property

' Ctor
Private Sub Class_Initialize()
    Set ValuePairs = New Collection
End Sub

' ---
Public Function IsValid() As Boolean
    If SourceKey Is Nothing Then Exit Function
    If DestinationKey Is Nothing Then Exit Function
    ' Commented these out, so that we can still transfer Keys across without value columns
    'If ValuePairs Is Nothing Then Exit Function
    'If ValuePairs.Count = 0 Then Exit Function
    IsValid = True
End Function

Public Function Transfer() As Boolean
    If Not IsValid Then Exit Function
        
    SetMapResult
    
    Dim ColumnTuple As ColumnTuple
    For Each ColumnTuple In ValuePairs
        TransferColumnTuple ColumnTuple
    Next ColumnTuple

    Transfer = True
End Function

Private Sub SetMapResult()
    'Dim Comparer As KeyColumnComparer
    'Set Comparer = KeyColumnComparer.Create(KeyColumn.FromColumn(SourceKey), KeyColumn.FromColumn(DestinationKey))
    
    'RHStoLHSRowMap = Comparer.Map(HasFlag(This.Flags, SourceFilteredOnly), HasFlag(This.Flags, DestinationFilteredOnly))
End Sub

Private Sub TransferColumnTuple(ByVal ColumnTuple As ColumnTuple)
   DoTransfer RHStoLHSRowMap, ColumnTuple.SourceListColumn, ColumnTuple.DestinationListColumn
End Sub

'@Description "Transfers data from a Source value column to a Destination value column using the provided KeyColumn mapping."
Private Sub DoTransfer(ByVal Map As Variant, ByVal Source As ListColumn, ByVal Destination As ListColumn)
Attribute DoTransfer.VB_Description = "Transfers data from a Source value column to a Destination value column using the provided KeyColumn mapping."
    Dim i As Integer
    Dim arrLHS As Variant
    Dim arrRHS As Variant
    Dim arrLHSOffset As Long
    Dim arrRHSOffset As Long
    Dim oldValue As Variant
    Dim newValue As Variant
    Dim doCopy As Boolean
    
    If Source.DataBodyRange Is Nothing Then
        ' This should never happen
        Exit Sub
    End If
    
    If Destination.DataBodyRange Is Nothing Then
        ' This will happen when transfering into an empty table without the flag
        ' append new keys
        Exit Sub
    End If
    
    arrLHS = Source.DataBodyRange.Value2
    arrRHS = Destination.DataBodyRange.Value2
    arrLHSOffset = 1 'Source.DataBodyRange.Row
    arrRHSOffset = 1 'Destination.DataBodyRange.Row
    
    Dim cellRange As Range
    Dim copiedRange As Range
    For i = LBound(Map) To UBound(Map)
        If Map(i) > -1 Then
            oldValue = arrRHS(i - arrRHSOffset + 1, 1)
            newValue = arrLHS(Map(i) - arrLHSOffset + 1, 1)
            doCopy = True
            
            If IsError(oldValue) Then doCopy = False
            If IsError(newValue) Then doCopy = False
            
            If doCopy Then
                If arrRHS(i - arrRHSOffset + 1, 1) <> newValue Then
                    arrRHS(i - arrRHSOffset + 1, 1) = newValue
                    'RangeHelpers.AppendRange Destination.DataBodyRange.Cells(i - arrRHSOffset + 1, 1), This.ModifiedCells
                Else
                    'Debug.Print "Cell was mapped but didn't change"
                End If
            End If
        End If
    Next i
    
    Destination.DataBodyRange.Value2 = arrRHS
    Destination.Parent.ShowAutoFilter = True
End Sub

Public Function ToString() As String
    Dim Result As String
    Dim i As Long
    
    Result = Result & "TRANSFER" & vbCrLf
    Result = Result & " SRC," & This.Source.Range.Address(external:=True) & vbCrLf
    Result = Result & " SRCKEY," & This.SourceKey.Name & vbCrLf
    Result = Result & " DST," & This.Destination.Range.Address(external:=True) & vbCrLf
    Result = Result & " DSTKEY," & This.DestinationKey.Name & vbCrLf
    Result = Result & " VALUES," & This.ValuePairs.Count & vbCrLf
    For i = 1 To This.ValuePairs.Count
         Result = Result & "  " & This.ValuePairs(i).ToString & vbCrLf
    Next i
    Result = Result & "END"
    
    ToString = Result
End Function

Public Function DebugName() As String
    If IsValid Then
        DebugName = This.Source.Name & "." & This.SourceKey.Name & " -> " & This.Destination.Name & "." & This.DestinationKey.Name & " (" & This.ValuePairs.Count & ")"
    Else
        DebugName = "unknown transfer instruction (could not de-serialize)"
    End If
End Function
