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
    CopiedCells As Range
End Type

Private this As TState

' Properties
Public Property Get Flags() As Integer
    Flags = this.Flags
End Property

Public Property Let Flags(ByVal Flags As Integer)
    this.Flags = Flags
End Property

Public Property Get Source() As ListObject
    Set Source = this.Source
End Property

Public Property Set Source(ByVal rhs As ListObject)
    Set this.Source = rhs
End Property

Public Property Get Destination() As ListObject
    Set Destination = this.Destination
End Property

Public Property Set Destination(ByVal rhs As ListObject)
    Set this.Destination = rhs
End Property

Public Property Get SourceKey() As ListColumn
    Set SourceKey = this.SourceKey
End Property

Public Property Set SourceKey(ByVal rhs As ListColumn)
    Set this.SourceKey = rhs
End Property

Public Property Get DestinationKey() As ListColumn
    Set DestinationKey = this.DestinationKey
End Property

Public Property Set DestinationKey(ByVal rhs As ListColumn)
    Set this.DestinationKey = rhs
End Property

Public Property Get ValuePairs() As Collection
    Set ValuePairs = this.ValuePairs
End Property

Public Property Set ValuePairs(ByVal rhs As Collection)
    Set this.ValuePairs = rhs
End Property

Public Property Get mapResult() As Variant
    If IsObject(this.mapResult) Then
        Set mapResult = this.mapResult
    Else
        mapResult = this.mapResult
    End If
End Property

Public Property Let mapResult(ByVal rhs As Variant)
    this.mapResult = rhs
End Property

Public Property Set mapResult(ByVal rhs As Variant)
    Set this.mapResult = rhs
End Property

' Ctor
Private Sub Class_Initialize()
    Set ValuePairs = New Collection
End Sub

' Methods
Public Sub Add(ByVal LHS As ListColumn, ByVal rhs As ListColumn)
    ValuePairs.Add ColumnPair.Create(LHS, rhs)
End Sub

Public Function IsValid() As Boolean
    If SourceKey Is Nothing Then Exit Function
    If DestinationKey Is Nothing Then Exit Function
    ' Commented these out, so that we can still transfer Keys across without value columns
    'If ValuePairs Is Nothing Then Exit Function
    'If ValuePairs.Count = 0 Then Exit Function
    IsValid = True
End Function

Public Sub transfer()
    If Not IsValid Then Exit Sub
        
    If HasFlag(this.Flags, RemoveUnmapped) Then
        RemoveUnmappedKeys
    End If
    
    If HasFlag(this.Flags, AppendUnmapped) Then
        AppendUnmappedKeys
    End If
    
    SetMapResult
    
    Dim colPair As ColumnPair
    For Each colPair In ValuePairs
        TransferOne colPair
    Next colPair
    
    UpdateHighlighting
End Sub

Private Sub SetMapResult()
    Dim compare As KeyColumnComparer
    Set compare = KeyColumnComparer.Create(KeyColumn.FromColumn(SourceKey), KeyColumn.FromColumn(DestinationKey))
    
    mapResult = compare.Map(HasFlag(this.Flags, SourceFilteredOnly), HasFlag(this.Flags, DestinationFilteredOnly))
End Sub

Private Sub TransferOne(ByVal colPair As ColumnPair)
   DoTransfer mapResult, colPair.LHS, colPair.rhs
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
    
    If HasFlag(this.Flags, DestinationFilteredOnly) = False Then
        If Destination.DataBodyRange.AutoFilter Then
            If Not Destination.parent.AutoFilter Is Nothing Then
                Destination.parent.AutoFilter.ShowAllData
            End If
        End If
    End If
    
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
    arrLHSOffset = Source.DataBodyRange.row
    arrRHSOffset = Destination.DataBodyRange.row
    
    Dim cellRange As Range
    Dim copiedRange As Range
    For i = LBound(Map) To UBound(Map)
        If HasFlag(this.Flags, ClearDestinationFirst) Then
            arrRHS(i - arrRHSOffset + 1, 1) = Empty
        End If
        
        If Map(i) > -1 Then
            oldValue = arrRHS(i - arrRHSOffset + 1, 1)
            newValue = arrLHS(Map(i) - arrLHSOffset + 1, 1)
            doCopy = True
            
            If HasFlag(this.Flags, ReplaceEmptyOnly) And (oldValue <> Empty) Then
                doCopy = False
            End If
            
            If (HasFlag(this.Flags, TransferBlanks) = False) And (newValue = Empty) Then
                doCopy = False
            End If
            
            If doCopy Then
                If arrRHS(i - arrRHSOffset + 1, 1) <> newValue Then
                    arrRHS(i - arrRHSOffset + 1, 1) = newValue
                    RangeHelpers.AppendRange Destination.DataBodyRange.Cells(i - arrRHSOffset + 1, 1), this.CopiedCells
                Else
                    'Debug.Print "Cell was mapped but didn't change"
                End If
                
                
            End If
        End If
    Next i
    
    If HasFlag(this.Flags, DestinationFilteredOnly) Then
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
     
    Set lc = DestinationKey 'GetListColumnFromRange(comp.rhs.Range)
    Set lo = lc.parent

    Dim c As Long
    If comp.rhs Is Nothing Then
        c = comp.LHS.Keys.Count
    Else
        c = comp.LeftOnly.Count
    End If
    For i = 1 To c
        Set lr = lo.ListRows.Add(alwaysinsert:=True)
        If comp.rhs Is Nothing Then
            lr.Range.Cells(1, lc.Index).Value2 = comp.LHS.Keys(i)
        Else
            lr.Range.Cells(1, lc.Index).Value2 = comp.LeftOnly.Item(i)
        End If
    Next i
End Sub

Public Function ToString() As String
    Dim result As String
    Dim i As Long
    
    result = result & "TRANSFER" & vbCrLf
    result = result & " SRC," & this.Source.Range.Address(external:=True) & vbCrLf
    result = result & " SRCKEY," & this.SourceKey.Name & vbCrLf
    result = result & " DST," & this.Destination.Range.Address(external:=True) & vbCrLf
    result = result & " DSTKEY," & this.DestinationKey.Name & vbCrLf
    result = result & " FLAGS," & this.Flags & vbCrLf
    result = result & " VALUES," & this.ValuePairs.Count & vbCrLf
    For i = 1 To this.ValuePairs.Count
         result = result & "  " & this.ValuePairs(i).ToString & vbCrLf
    Next i
    result = result & "END"
    
    ToString = result
End Function

Public Function Name() As String
    Name = this.Source.Name & "." & this.SourceKey.Name & " -> " & this.Destination.Name & "." & this.DestinationKey.Name & " (" & this.ValuePairs.Count & ")"
End Function

Public Sub SetDefaultFlags()
    Flags = 0
    Flags = AddFlag(Flags, SourceFilteredOnly)
    Flags = AddFlag(Flags, DestinationFilteredOnly)
    Flags = AddFlag(Flags, HighlightMapped)
End Sub

Private Sub UpdateHighlighting()
    If Not this.CopiedCells Is Nothing Then
        Highlighting.ApplyHighlighting this.CopiedCells, HasFlag(this.Flags, HighlightMapped)
    End If
End Sub
