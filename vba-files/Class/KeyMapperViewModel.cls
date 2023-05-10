VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyMapperViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("KeyMapper")
Option Explicit
Implements IViewModel

Private Const BEST_MATCH_DISTINCT_THRESHOLD = 0.75

Public Event PropertyChanged(ByVal propertyName As String)
Public Event CheckCompleted()
Public Event PreviewChanged()
Public Event MatchChanged()

Private Type TViewModel
    LHSTable As ListObject
    RHSTable As ListObject
    LHSKeyColumn As ListColumn
    RHSKeyColumn As ListColumn
    LHSColumns As Collection ' Remove?
    RHSColumns As Collection ' Remove?
    IsDirty As Boolean
    IsDirtyQuality As Boolean
    IsDirtyMatch As Boolean
    IsAutoGuessed As Boolean
    '---
    QualityLHS As KeyColumn
    QualityRHS As KeyColumn
    MatchResults As KeyColumnComparer
    '---
    GoBack As Boolean
    AppendNewKeys As Boolean
    RemoveOrphanKeys As Boolean
End Type

Private this As TViewModel

Public Property Get AppendNewKeys() As Boolean
    AppendNewKeys = this.AppendNewKeys
End Property
    
Public Property Let AppendNewKeys(ByVal rhs As Boolean)
    this.AppendNewKeys = rhs
End Property

Public Property Get RemoveOrphanKeys() As Boolean
    RemoveOrphanKeys = this.RemoveOrphanKeys
End Property
    
Public Property Let RemoveOrphanKeys(ByVal rhs As Boolean)
    this.RemoveOrphanKeys = rhs
End Property

Public Property Get GoBack() As Boolean
    GoBack = this.GoBack
End Property
    
Public Property Let GoBack(ByVal rhs As Boolean)
    this.GoBack = rhs
End Property

Public Property Let IsDirty(ByVal rhs As Boolean)
    this.IsDirty = rhs
End Property

Public Property Get LHSTable() As ListObject
    Set LHSTable = this.LHSTable
End Property

Public Property Set LHSTable(ByVal lo As ListObject)
    If lo Is Nothing Then
        Exit Property
    End If
    
    If this.LHSTable Is lo Then
        Exit Property
    End If
    
    Set this.LHSTable = lo
    this.IsDirty = True
    this.IsAutoGuessed = False
    
    UpdateColumns this.LHSTable, this.LHSColumns
    
    Set LHSKeyColumn = this.LHSColumns.Item(1)
    
    RaiseEvent PropertyChanged(KeyMapperEvents.LHS_COLUMNS)
    RaiseEvent PropertyChanged(KeyMapperEvents.LHS_TABLE)
End Property

Public Property Get RHSTable() As ListObject
    Set RHSTable = this.RHSTable
End Property

Public Property Set RHSTable(ByVal lo As ListObject)
    If lo Is Nothing Then
        Exit Property
    End If
    
    If this.RHSTable Is lo Then
        Exit Property
    End If
    
    Set this.RHSTable = lo
    this.IsDirty = True
    this.IsAutoGuessed = False
    
    UpdateColumns this.RHSTable, this.RHSColumns
    
    Set RHSKeyColumn = this.RHSColumns.Item(1)
        
    RaiseEvent PropertyChanged(KeyMapperEvents.RHS_COLUMNS)
    RaiseEvent PropertyChanged(KeyMapperEvents.RHS_TABLE)
End Property

Public Property Get LHSColumns() As Collection
    Set LHSColumns = this.LHSColumns
End Property

Public Property Set LHSColumns(ByVal rhs As Collection)
    Set this.LHSColumns = rhs
End Property

Public Property Get RHSColumns() As Collection
    Set RHSColumns = this.RHSColumns
End Property

Public Property Set RHSColumns(ByVal rhs As Collection)
    Set this.RHSColumns = rhs
End Property

Public Property Get LHSKeyColumn() As ListColumn
    Set LHSKeyColumn = this.LHSKeyColumn
End Property

Public Property Set LHSKeyColumn(ByVal lo As ListColumn)
    If Not lo Is Nothing Then
        If Not this.LHSKeyColumn Is Nothing Then
            If this.LHSKeyColumn.Name = lo.Name Then Exit Property
        End If
    End If
    this.IsDirtyQuality = True
    this.IsDirtyMatch = False
    this.IsDirty = True
    Set this.LHSKeyColumn = lo
    RaiseEvent PropertyChanged(KeyMapperEvents.LHS_KEY_COLUMN)
End Property

Public Property Get RHSKeyColumn() As ListColumn
    Set RHSKeyColumn = this.RHSKeyColumn
End Property

Public Property Set RHSKeyColumn(ByVal lo As ListColumn)
    If Not lo Is Nothing Then
        If Not this.RHSKeyColumn Is Nothing Then
            If this.RHSKeyColumn.Name = lo.Name Then Exit Property
        End If
    End If
    this.IsDirtyQuality = True
    this.IsDirtyMatch = False
    this.IsDirty = True
    Set this.RHSKeyColumn = lo
    RaiseEvent PropertyChanged(KeyMapperEvents.RHS_KEY_COLUMN)
End Property

' --- Subs
Public Sub LoadFromTransferInstruction(ByVal ti As TransferInstruction)
    If Not ti.Source Is Nothing Then
        Set LHSTable = ti.Source
    End If
    
    If Not ti.Source Is Nothing Then
        Set RHSTable = ti.Destination
    End If
    
    If Not ti.Source Is Nothing Then
        Set LHSKeyColumn = ti.SourceKey
    End If
    
    If Not ti.Source Is Nothing Then
        Set RHSKeyColumn = ti.DestinationKey
    End If
End Sub

Private Sub UpdateColumns(ByVal lo As ListObject, ByVal coll As Collection)
    ' TODO This fails when selecting new table in key select dialog
    Debug.Assert Not lo Is Nothing
    
    Dim i As Long
    Dim lc As ListColumn
    
    For i = coll.Count To 1 Step -1
        coll.Remove i
    Next i
    
    For Each lc In lo.ListColumns
        coll.Add Item:=lc, key:=lc.Name
    Next lc
End Sub

Private Sub Class_Initialize()
    Set this.LHSColumns = New Collection
    Set this.RHSColumns = New Collection
End Sub

Public Sub TrySelectLHS(ByVal key As String)
    If key <> vbNullString Then
        On Error Resume Next
        Set LHSKeyColumn = this.LHSColumns(key)
        On Error GoTo 0
        Debug.Assert Not LHSKeyColumn Is Nothing
    End If
End Sub

Public Sub TrySelectRHS(ByVal key As String)
    If key <> vbNullString Then
        ' TODO This failed in production
        On Error Resume Next
        Set RHSKeyColumn = this.RHSColumns(key)
        On Error GoTo 0
        Debug.Assert Not RHSKeyColumn Is Nothing
    End If
End Sub

Public Function TryAutoMatch(ByVal leftToRight As Boolean, ByVal Quiet As Boolean) As Boolean
    If leftToRight And LHSKeyColumn Is Nothing Then
        Debug.Assert False
    End If
    
    If leftToRight = False And LHSKeyColumn Is Nothing Then
        Debug.Assert False
    End If
    
    If Not LHSKeyColumn Is Nothing And Not RHSKeyColumn Is Nothing Then
        If LHSKeyColumn.Name = RHSKeyColumn.Name Then
            Exit Function
        End If
    End If
    
    Dim searchName As String
    Dim coll As Collection
    
    If leftToRight Then
        searchName = LHSKeyColumn.Name
        Set coll = this.RHSColumns
    Else
        searchName = RHSKeyColumn.Name
        Set coll = this.LHSColumns
    End If
    
    Dim lc As ListColumn
    Dim result As ListColumn
    For Each lc In coll
        If lc.Name = searchName Then
            Set result = lc
        End If
    Next lc
    
    If Not result Is Nothing Then
        ' TODO We should not be asking for user input via ViewModel
        If Quiet = False Then
            If (vbNo = MsgBox("Auto-match column name?", vbQuestion + vbYesNo)) Then
                Exit Function
            End If
        End If
        
        If leftToRight Then
            Set RHSKeyColumn = result
        Else
            Set LHSKeyColumn = result
        End If
        
        TryAutoMatch = True
    End If
End Function

Public Sub DoCheck()
    Debug.Assert this.IsDirty = True
    
    'MsgBox "DoWork"
    
    this.IsDirty = False
    
    RaiseEvent CheckCompleted
End Sub

Private Function IViewModel_IsValid() As Boolean
    If this.LHSKeyColumn Is Nothing Then Exit Function
    If this.RHSKeyColumn Is Nothing Then Exit Function
    IViewModel_IsValid = True
End Function

Public Function IsValid() As Boolean
    IsValid = IViewModel_IsValid
End Function

Public Function CanCheck() As Boolean
    If this.LHSTable Is Nothing Then Exit Function
    If this.RHSTable Is Nothing Then Exit Function
    If this.LHSKeyColumn Is Nothing Then Exit Function
    If this.RHSKeyColumn Is Nothing Then Exit Function
    If this.IsDirtyQuality = False Then Exit Function
    CanCheck = True
End Function

Public Function CanMatch() As Boolean
    CanMatch = this.IsDirtyMatch
End Function

Public Function CanContinue() As Boolean
    ' Should only continue if KeyColumnComparer is valid and .Inner.Count > 0
    CanContinue = (this.IsDirtyQuality = False) And (this.IsDirtyMatch = False)
End Function

Public Sub UpdatePreviews()
    this.IsDirtyQuality = False
    this.IsDirtyMatch = True
    RaiseEvent PreviewChanged
End Sub

Public Sub UpdateMatch()
    this.IsDirtyMatch = False
    RaiseEvent MatchChanged
End Sub

Public Function TryGuess(Optional Force As Boolean = False) As Boolean
    'Debug.Print "TryGuess()"
    If this.LHSTable Is Nothing Then Exit Function
    If this.RHSTable Is Nothing Then Exit Function
    If (Force = False) And (this.IsAutoGuessed) Then Exit Function
    
    Dim bestScore As Double
    Dim bestLHS As ListColumn
    Dim bestRHS As ListColumn
    bestScore = 0#
    
    Dim lhsCount As Long
    Dim rhsCount As Long
    lhsCount = this.LHSColumns.Count
    rhsCount = this.RHSColumns.Count
    If lhsCount > 10 Then lhsCount = 10
    If rhsCount > 10 Then rhsCount = 10
    
    Dim startTime As Double
    startTime = Timer()
    Dim timeLimit As Double
    timeLimit = 3#
    
    Dim i As Long
    Dim j As Long
    Dim thisScore As Double
    Dim comp As KeyColumnComparer
    Set comp = New KeyColumnComparer
    For i = 1 To lhsCount
        Set comp.lhs = KeyColumn.FromColumn(this.LHSColumns(i), False, 1000)
        'PrintTime "KeyColumn.FromColumn(this.LHSColumns(i))"
        If comp.lhs.DistinctPercentage > BEST_MATCH_DISTINCT_THRESHOLD Then
            For j = 1 To rhsCount
                If (Timer() - startTime) > timeLimit Then GoTo TimeLimitExceeded
                Set comp.rhs = KeyColumn.FromColumn(this.RHSColumns(j), False, 1000)
                'PrintTime "KeyColumn.FromColumn(this.RHSColumns(i))"
                If Not comp.rhs Is Nothing Then
                    If comp.rhs.DistinctPercentage > BEST_MATCH_DISTINCT_THRESHOLD Then
                        comp.Map
    
                        If comp.Score > bestScore Then
                            'Debug.Print "New best score: " & comp.Score & " (" & i & "," & j & ")"
                            bestScore = comp.Score
                            Set bestLHS = this.LHSColumns(i)
                            Set bestRHS = this.RHSColumns(j)
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    
TimeLimitExceeded:
    If bestScore = 0# Then Exit Function
    
    Set LHSKeyColumn = bestLHS
    Set RHSKeyColumn = bestRHS
    
    this.IsAutoGuessed = True
    TryGuess = True
End Function

Public Function CanSwap() As Boolean
    If LHSTable Is Nothing Then Exit Function
    If RHSTable Is Nothing Then Exit Function
    If LHSKeyColumn Is Nothing Then Exit Function
    If RHSKeyColumn Is Nothing Then Exit Function
    CanSwap = True
End Function

Public Sub TrySwap()
    If CanSwap = False Then
        Exit Sub
    End If
    
    Dim lhsKeyColName As String
    Dim rhsKeyColName As String
    
    lhsKeyColName = this.LHSKeyColumn.Name
    rhsKeyColName = this.RHSKeyColumn.Name
    
    Dim lo As ListObject
    Set lo = this.LHSTable
    Set LHSTable = this.RHSTable
    Set RHSTable = lo
    
    TrySelectLHS rhsKeyColName
    TrySelectRHS lhsKeyColName
End Sub
