VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule
'@Folder "Model.TransferInstruction"
Option Explicit

Private Type TState
    Source As ListObject
    Destination As ListObject
    SourceKey As ListColumn
    DestinationKey As ListColumn
    ValuePairs As Collection
    MapResult As Variant
    Flags As Integer
    CopiedCells As Range
    LastRun As Date
    UnRef As TransferInstructionUnref
End Type

Private This As TState

' Properties
Public Property Get UnRef() As TransferInstructionUnref
    Set UnRef = This.UnRef
End Property

Public Property Set UnRef(ByVal RHS As TransferInstructionUnref)
    Set This.UnRef = RHS
End Property

Public Property Get LastRun() As Date
    LastRun = This.LastRun
End Property

Public Property Get Flags() As Integer
    Flags = This.Flags
End Property

Public Property Let Flags(ByVal Flags As Integer)
    This.Flags = Flags
End Property

Public Property Get Source() As ListObject
    Set Source = This.Source
End Property

Public Property Set Source(ByVal RHS As ListObject)
    Set This.Source = RHS
End Property

Public Property Get Destination() As ListObject
    Set Destination = This.Destination
End Property

Public Property Set Destination(ByVal RHS As ListObject)
    Set This.Destination = RHS
End Property

Public Property Get SourceKey() As ListColumn
    Set SourceKey = This.SourceKey
End Property

Public Property Set SourceKey(ByVal RHS As ListColumn)
    Set This.SourceKey = RHS
End Property

Public Property Get DestinationKey() As ListColumn
    Set DestinationKey = This.DestinationKey
End Property

Public Property Set DestinationKey(ByVal RHS As ListColumn)
    Set This.DestinationKey = RHS
End Property

Public Property Get ValuePairs() As Collection
    Set ValuePairs = This.ValuePairs
End Property

Public Property Set ValuePairs(ByVal RHS As Collection)
    Set This.ValuePairs = RHS
End Property

Public Property Get MapResult() As Variant
    If IsObject(This.MapResult) Then
        Set MapResult = This.MapResult
    Else
        MapResult = This.MapResult
    End If
End Property

Public Property Let MapResult(ByVal RHS As Variant)
    This.MapResult = RHS
End Property

Public Property Set MapResult(ByVal RHS As Variant)
    Set This.MapResult = RHS
End Property

' Ctor
Private Sub Class_Initialize()
    Set ValuePairs = New Collection
End Sub

' Methods
Public Sub Add(ByVal LHS As ListColumn, ByVal RHS As ListColumn)
    ValuePairs.Add ColumnPair.Create(LHS, RHS)
End Sub

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
        
    If HasFlag(This.Flags, RemoveUnmapped) Then
        RemoveUnmappedKeys
    End If
    
    If HasFlag(This.Flags, AppendUnmapped) Then
        AppendUnmappedKeys
    End If
    
    SetMapResult
    
    Dim ColPair As ColumnPair
    For Each ColPair In ValuePairs
        TransferOne ColPair
    Next ColPair
    
    UpdateHighlighting

    Transfer = True
End Function

Private Sub SetMapResult()
    Dim Compare As KeyColumnComparer
    Set Compare = KeyColumnComparer.Create(KeyColumn.FromColumn(SourceKey), KeyColumn.FromColumn(DestinationKey))
    
    MapResult = Compare.Map(HasFlag(This.Flags, SourceFilteredOnly), HasFlag(This.Flags, DestinationFilteredOnly))
End Sub

Private Sub TransferOne(ByVal ColPair As ColumnPair)
    DoTransfer MapResult, ColPair.LHS, ColPair.RHS
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
    arrLHSOffset = Source.DataBodyRange.Row
    arrRHSOffset = Destination.DataBodyRange.Row
    
    Dim cellRange As Range
    Dim copiedRange As Range
    For i = LBound(Map) To UBound(Map)
        If HasFlag(This.Flags, ClearDestinationFirst) Then
            arrRHS(i - arrRHSOffset + 1, 1) = Empty
        End If
        
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
                If arrRHS(i - arrRHSOffset + 1, 1) <> newValue Then
                    arrRHS(i - arrRHSOffset + 1, 1) = newValue
                    RangeHelpers.AppendRange Destination.DataBodyRange.Cells(i - arrRHSOffset + 1, 1), This.CopiedCells
                Else
                    'Debug.Print "Cell was mapped but didn't change"
                End If
                
                
            End If
        End If
    Next i
    
    If HasFlag(This.Flags, DestinationFilteredOnly) Then
        ArrayToFilteredRange arrRHS, Destination.DataBodyRange
    Else
        Destination.DataBodyRange.Value2 = arrRHS
        Destination.parent.ShowAutoFilter = True
    End If
    
    This.LastRun = Now()
End Sub

Private Sub RemoveUnmappedKeys()
    ' TODO Refactor
    Dim comp As KeyColumnComparer
    Set comp = KeyColumnComparer.Create(KeyColumn.FromColumn(SourceKey), KeyColumn.FromColumn(DestinationKey))
    
    Dim cachedMappedResults As Variant
    cachedMappedResults = comp.Map
   
    Dim i As Long
    Dim rng As Range
    Set rng = comp.RHS.Range
    
    For i = rng.rows.Count To 1 Step -1
        If cachedMappedResults(rng.rows(i).Row) = -1 Then
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
     
    Set lc = DestinationKey                      'GetListColumnFromRange(comp.rhs.Range)
    Set lo = lc.parent

    Dim c As Long
    If comp.RHS Is Nothing Then
        c = comp.LHS.Keys.Count
    Else
        c = comp.LeftOnly.Count
    End If
    For i = 1 To c
        Set lr = lo.ListRows.Add(alwaysinsert:=True)
        If comp.RHS Is Nothing Then
            lr.Range.Cells(1, lc.Index).Value2 = comp.LHS.Keys(i)
        Else
            lr.Range.Cells(1, lc.Index).Value2 = comp.LeftOnly.Item(i)
        End If
    Next i
End Sub

Public Function ToString() As String
    Dim Result As String
    Dim i As Long
    
    Result = Result & "TRANSFER" & vbCrLf
    Result = Result & " SRC," & This.Source.Range.Address(External:=True) & vbCrLf
    Result = Result & " SRCKEY," & This.SourceKey.Name & vbCrLf
    Result = Result & " DST," & This.Destination.Range.Address(External:=True) & vbCrLf
    Result = Result & " DSTKEY," & This.DestinationKey.Name & vbCrLf
    Result = Result & " FLAGS," & This.Flags & vbCrLf
    Result = Result & " VALUES," & This.ValuePairs.Count & vbCrLf
    For i = 1 To This.ValuePairs.Count
        Result = Result & "  " & This.ValuePairs(i).ToString & vbCrLf
    Next i
    Result = Result & "END"
    
    ToString = Result
End Function

Public Function Name() As String
    If IsValid Then
        Name = This.Source.Name & "." & This.SourceKey.Name & " -> " & This.Destination.Name & "." & This.DestinationKey.Name & " (" & This.ValuePairs.Count & ")"
    Else
        Name = "unknown transfer instruction (could not de-serialize)"
    End If
End Function

Public Sub SetDefaultFlags()
    Flags = 0
    Flags = AddFlag(Flags, SourceFilteredOnly)
    Flags = AddFlag(Flags, DestinationFilteredOnly)
    Flags = AddFlag(Flags, HighlightMapped)
End Sub

Private Sub UpdateHighlighting()
    If Not This.CopiedCells Is Nothing Then
        Highlighting.ApplyHighlighting This.CopiedCells, HasFlag(This.Flags, HighlightMapped)
    End If
End Sub

Public Sub LoadFlags()
    If This.UnRef Is Nothing Then Exit Sub
    This.Flags = This.UnRef.Flags
End Sub

Public Function TryLoadTables() As Boolean
    If This.UnRef Is Nothing Then Exit Function

    Dim wbLHS As Workbook
    Dim wbRHS As Workbook
    
    If ListObjectHelpers.TryGetWorkbook(This.UnRef.SourceFilename, wbLHS) = False Then
        Exit Function
    End If
    
    If ListObjectHelpers.TryGetWorkbook(This.UnRef.DestinationFilename, wbRHS) = False Then
        Exit Function
    End If
    
    Dim wsLHS As Worksheet
    Dim wsRHS As Worksheet

    On Error Resume Next
    Set wsLHS = wbLHS.Worksheets.Item(This.UnRef.SourceSheet)
    Set wsRHS = wbRHS.Worksheets.Item(This.UnRef.DestinationSheet)
    On Error GoTo 0
    
    If wsLHS Is Nothing Then Exit Function
    If wsRHS Is Nothing Then Exit Function
    
    Dim loLHS As ListObject
    Dim loRHS As ListObject

    On Error Resume Next
    Set loLHS = wsLHS.ListObjects.Item(This.UnRef.Source)
    Set loRHS = wsRHS.ListObjects.Item(This.UnRef.Destination)
    On Error GoTo 0
    
    If loLHS Is Nothing Then Exit Function
    If loRHS Is Nothing Then Exit Function
    
    Set This.Source = loLHS
    Set This.Destination = loRHS
    
    TryLoadTables = True
End Function

Public Function TryLoadKeyColumns() As Boolean
    If This.UnRef Is Nothing Then Exit Function
    TryLoadKeyColumns = True
    
    Dim lc As ListColumn
    
    On Error Resume Next
    Set lc = This.Source.ListColumns.Item(This.UnRef.SourceKey)
    On Error GoTo 0
    If Not lc Is Nothing Then
        Set This.SourceKey = lc
    Else
        TryLoadKeyColumns = False
    End If
    
    Set lc = Nothing
    On Error Resume Next
    Set lc = This.Destination.ListColumns.Item(This.UnRef.DestinationKey)
    On Error GoTo 0
    If Not lc Is Nothing Then
        Set This.DestinationKey = lc
    Else
        TryLoadKeyColumns = False
    End If
End Function

Public Function TryLoadValuePairs() As Boolean
    If This.UnRef Is Nothing Then Exit Function
    TryLoadValuePairs = True
    
    Dim vp As Variant
    vp = This.UnRef.ValuePairs
    
    CollectionHelpers.CollectionClear This.ValuePairs
    
    Dim i As Long
    For i = 1 To UBound(vp, 1)
        Dim LHS As ListColumn
        Dim RHS As ListColumn
        On Error Resume Next
        Set LHS = This.Source.ListColumns.Item(vp(i, 1))
        Set RHS = This.Destination.ListColumns.Item(vp(i, 2))
        On Error GoTo 0
        If Not LHS Is Nothing And Not RHS Is Nothing Then
            This.ValuePairs.Add ColumnPair.Create(LHS, RHS)
        Else
            TryLoadValuePairs = False
        End If
    Next i
End Function

Public Sub SaveToRange(ByVal rng As Range)
    Dim c As Long
    c = 12 + This.ValuePairs.Count

    Dim arr() As String
    ReDim arr(1 To c, 1 To 4)
    
    arr(1, 1) = "TRANSFER"
    arr(2, 2) = "NAME"
    arr(3, 2) = "TIME"
    arr(4, 2) = "PATH"
    arr(5, 2) = "FN"
    arr(6, 2) = "SHEET"
    arr(7, 2) = "RNG"
    arr(8, 2) = "TBL"
    arr(9, 2) = "KEYS"
    arr(10, 2) = "FLAGS"
    arr(11, 2) = "PAIRS"
    arr(c, 1) = "END"
    
    arr(2, 3) = Name
    arr(3, 3) = This.LastRun
    
    arr(4, 3) = This.Source.parent.parent.path
    arr(5, 3) = This.Source.parent.parent.Name
    arr(6, 3) = This.Source.parent.Name
    arr(7, 3) = This.Source.Range.Address
    arr(8, 3) = This.Source.Name
    arr(9, 3) = This.SourceKey.Name
    
    arr(4, 4) = This.Destination.parent.parent.path
    arr(5, 4) = This.Destination.parent.parent.Name
    arr(6, 4) = This.Destination.parent.Name
    arr(7, 4) = This.Destination.Range.Address
    arr(8, 4) = This.Destination.Name
    arr(9, 4) = This.DestinationKey.Name
    
    arr(10, 3) = This.Flags
    arr(11, 3) = This.ValuePairs.Count
    
    Dim n As Long
    n = 12
    Dim cp As ColumnPair
    For Each cp In This.ValuePairs
        arr(n, 3) = cp.LHS.Name
        arr(n, 4) = cp.RHS.Name
        n = n + 1
    Next cp
    
    Set rng = rng.Resize(c, 4)
    
    rng.Value2 = arr
    'rng.Select
End Sub

