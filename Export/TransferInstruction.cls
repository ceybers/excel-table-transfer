VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "TableTransfer"
Option Explicit

Private Type TState
    Source As ListObject
    Destination As ListObject
    SourceKey As ListColumn
    DestinationKey As ListColumn
    ValuePairs As Collection
    mapResult As Variant
    Flags As Integer
End Type

Private This As TState

' Properties
Public Property Get Flags() As Integer
    Flags = This.Flags
End Property

Public Property Let Flags(ByVal Flags As Integer)
    This.Flags = Flags
End Property

Public Property Get Source() As ListObject
    Set Source = This.Source
End Property

Public Property Set Source(ByVal rhs As ListObject)
    Set This.Source = rhs
End Property

Public Property Get Destination() As ListObject
    Set Destination = This.Destination
End Property

Public Property Set Destination(ByVal rhs As ListObject)
    Set This.Destination = rhs
End Property

Public Property Get SourceKey() As ListColumn
    Set SourceKey = This.SourceKey
End Property

Public Property Set SourceKey(ByVal rhs As ListColumn)
    Set This.SourceKey = rhs
End Property

Public Property Get DestinationKey() As ListColumn
    Set DestinationKey = This.DestinationKey
End Property

Public Property Set DestinationKey(ByVal rhs As ListColumn)
    Set This.DestinationKey = rhs
End Property

Public Property Get ValuePairs() As Collection
    Set ValuePairs = This.ValuePairs
End Property

Public Property Set ValuePairs(ByVal rhs As Collection)
    Set This.ValuePairs = rhs
End Property

Public Property Get mapResult() As Variant
    If IsObject(This.mapResult) Then
        Set mapResult = This.mapResult
    Else
        mapResult = This.mapResult
    End If
End Property

Public Property Let mapResult(ByVal rhs As Variant)
    This.mapResult = rhs
End Property

Public Property Set mapResult(ByVal rhs As Variant)
    Set This.mapResult = rhs
End Property

' Ctor
Private Sub Class_Initialize()
    Set ValuePairs = New Collection
End Sub

' Methods
Public Sub Add(ByVal lhs As ListColumn, ByVal rhs As ListColumn)
    ValuePairs.Add ColumnPair.Create(lhs, rhs)
End Sub

Public Function IsValid() As Boolean
    If SourceKey Is Nothing Then Exit Function
    If DestinationKey Is Nothing Then Exit Function
    If ValuePairs Is Nothing Then Exit Function
    If ValuePairs.Count = 0 Then Exit Function
    IsValid = True
End Function

Public Sub Transfer()
    If Not IsValid Then Exit Sub
        
    If HasFlag(This.Flags, RemoveUnmapped) Then
        RemoveUnmappedKeys
    End If
    
    If HasFlag(This.Flags, appendunmapped) Then
        AppendUnmappedKeys
    End If
    
    SetMapResult
    
    Dim colPair As ColumnPair
    For Each colPair In ValuePairs
        TransferOne colPair
    Next colPair
End Sub

Private Sub SetMapResult()
    Dim compare As KeyColumnComparer
    Set compare = KeyColumnComparer.Create(KeyColumn.FromColumn(SourceKey), KeyColumn.FromColumn(DestinationKey))
    
    mapResult = compare.Map(HasFlag(This.Flags, SourceFilteredOnly), HasFlag(This.Flags, DestinationFilteredOnly))
End Sub

Private Sub TransferOne(ByVal colPair As ColumnPair)
   DoTransfer mapResult, colPair.lhs, colPair.rhs
End Sub

Private Sub DoTransfer(ByVal Map As Variant, ByVal Source As ListColumn, ByVal Destination As ListColumn)
    Dim i As Integer
    Dim arrLHS As Variant
    Dim arrRHS As Variant
    Dim arrLHSOffset As Long
    Dim arrRHSOffset As Long
    Dim oldValue As Variant
    Dim newValue As Variant
    Dim doCopy As Boolean

    If HasFlag(This.Flags, DestinationFilteredOnly) = False Then
        If Destination.DataBodyRange.AutoFilter Then
            If Not Destination.parent.AutoFilter Is Nothing Then
                Destination.parent.AutoFilter.ShowAllData
            End If
        End If
    End If
    
    arrLHS = Source.DataBodyRange.Value2
    arrRHS = Destination.DataBodyRange.Value2
    arrLHSOffset = Source.DataBodyRange.row
    arrRHSOffset = Destination.DataBodyRange.row
    
    For i = LBound(Map) To UBound(Map)
        If Map(i) > -1 Then
            oldValue = arrRHS(i - arrRHSOffset + 1, 1)
            newValue = arrLHS(Map(i) - arrLHSOffset + 1, 1)
            doCopy = True
            
            If HasFlag(This.Flags, ReplaceEmptyOnly) And (oldValue <> Empty) Then
                doCopy = False
            End If
            
            If (HasFlag(This.Flags, TransferBlanks) = False) And (newValue = Empty) Then
                doCopy = False
            End If
            
            If doCopy Then
                arrRHS(i - arrRHSOffset + 1, 1) = newValue
            End If
        Else ' Row is not mapped by keys
            If HasFlag(This.Flags, ClearDestinationFirst) Then
                arrRHS(i - arrRHSOffset + 1, 1) = Empty
            Else
                'arrRHS(i - arrRHSOffset + 1, 1) = "Unmapped"
            End If
        End If
    Next i
    
    If HasFlag(This.Flags, DestinationFilteredOnly) Then
        ArrayToFilteredRange Destination.DataBodyRange, arrRHS
    Else
        Destination.DataBodyRange.Value2 = arrRHS
        Destination.parent.ShowAutoFilter = True
    End If
End Sub

Private Sub RemoveUnmappedKeys()
    ' TODO Refactor
    Dim comp As KeyColumnComparer
    Set comp = KeyColumnComparer.Create(KeyColumn.FromColumn(SourceKey), KeyColumn.FromColumn(DestinationKey))
    
    Dim cachedMappedResults As Variant
    cachedMappedResults = comp.Map
   
    Dim i As Long
    Dim rng As Range
    Set rng = comp.rhs.Range
    
    For i = rng.rows.Count To 1 Step -1
        If cachedMappedResults(rng.rows(i).row) = -1 Then
             rng.rows(i).EntireRow.Delete
        End If
    Next i
End Sub

Private Sub AppendUnmappedKeys()
    ' TODO Refactor
    Dim comp As KeyColumnComparer
    Set comp = KeyColumnComparer.Create(KeyColumn.FromColumn(SourceKey), KeyColumn.FromColumn(DestinationKey))
    
    Dim lo As ListObject
    Dim lc As ListColumn
    Dim lr As ListRow
    Dim i As Long
     
    Set lc = GetListColumnFromRange(comp.rhs.Range)
    Set lo = lc.parent

    For i = 1 To comp.LeftOnly.Count
        Set lr = lo.ListRows.Add(alwaysinsert:=True)
        lr.Range.Cells(1, lc.Index).Value2 = comp.LeftOnly.Item(i)
    Next i
End Sub

Public Function ToString() As String
    Dim result As String
    Dim i As Long
    
    result = result & "TRANSFER" & vbCrLf
    result = result & " SRC," & This.Source.Range.Address(external:=True) & vbCrLf
    result = result & " SRCKEY," & This.SourceKey.Name & vbCrLf
    result = result & " DST," & This.Destination.Range.Address(external:=True) & vbCrLf
    result = result & " DSTKEY," & This.DestinationKey.Name & vbCrLf
    result = result & " FLAGS," & This.Flags & vbCrLf
    result = result & " VALUES," & This.ValuePairs.Count & vbCrLf
    For i = 1 To This.ValuePairs.Count
         result = result & "  " & This.ValuePairs(i).ToString & vbCrLf
    Next i
    result = result & "END"
    
    ToString = result
End Function

Public Function Name() As String
    Name = This.Source.Name & "." & This.SourceKey.Name & " -> " & This.Destination.Name & "." & This.DestinationKey.Name & " (" & This.ValuePairs.Count & ")"
End Function
