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

Public Property Set LHSTable(ByVal rhs As ListObject)
    If this.LHSTable Is rhs Then Exit Property
    Set this.LHSTable = rhs
    this.IsDirty = True
    Debug.Assert Not rhs Is Nothing
    UpdateColumns this.LHSTable, this.LHSColumns
    RaiseEvent PropertyChanged("LHSTable")
    RaiseEvent PropertyChanged("LHSColumns")
End Property

Public Property Get RHSTable() As ListObject
    Set RHSTable = this.RHSTable
End Property

Public Property Set RHSTable(ByVal rhs As ListObject)
    If this.RHSTable Is rhs Then Exit Property
    Set this.RHSTable = rhs
    this.IsDirty = True
    UpdateColumns this.RHSTable, this.RHSColumns
    RaiseEvent PropertyChanged("RHSTable")
    RaiseEvent PropertyChanged("RHSColumns")
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

Public Property Set LHSKeyColumn(ByVal rhs As ListColumn)
    If Not this.LHSKeyColumn Is Nothing Then
        If this.LHSKeyColumn.Name = rhs.Name Then Exit Property
    End If
    this.IsDirtyQuality = True
    this.IsDirtyMatch = False
    this.IsDirty = True
    Set this.LHSKeyColumn = rhs
    RaiseEvent PropertyChanged("LHSKeyColumn")
End Property

Public Property Get RHSKeyColumn() As ListColumn
    Set RHSKeyColumn = this.RHSKeyColumn
End Property

Public Property Set RHSKeyColumn(ByVal rhs As ListColumn)
    If Not this.RHSKeyColumn Is Nothing Then
        If this.RHSKeyColumn.Name = rhs.Name Then Exit Property
    End If
    this.IsDirtyQuality = True
    this.IsDirtyMatch = False
    this.IsDirty = True
    Set this.RHSKeyColumn = rhs
    RaiseEvent PropertyChanged("RHSKeyColumn")
End Property

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

Public Sub TryAutoMatch(ByVal leftToRight As Boolean, ByVal Quiet As Boolean)
    If Not LHSKeyColumn Is Nothing And Not RHSKeyColumn Is Nothing Then
        If LHSKeyColumn.Name = RHSKeyColumn.Name Then
            Exit Sub
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
                Exit Sub
            End If
        End If
        
        If leftToRight Then
            Set RHSKeyColumn = result
        Else
            Set LHSKeyColumn = result
        End If
    End If
End Sub

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

Public Sub TryGuess(Optional Force As Boolean = False)
    If this.LHSTable Is Nothing Then Exit Sub
    If this.RHSTable Is Nothing Then Exit Sub
    If (Force = False) And (this.IsAutoGuessed) Then Exit Sub
    
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
    
    Dim i As Long
    Dim j As Long
    Dim thisScore As Double
    Dim comp As KeyColumnComparer
    Set comp = New KeyColumnComparer
    For i = 1 To lhsCount
        Set comp.lhs = KeyColumn.FromColumn(this.LHSColumns(i))
        For j = 1 To rhsCount
            Set comp.rhs = KeyColumn.FromColumn(this.RHSColumns(j))

            comp.Map

            If comp.Score > bestScore Then
                Debug.Print "New best score: " & comp.Score & " (" & i & "," & j & ")"
                bestScore = comp.Score
                Set bestLHS = this.LHSColumns(i)
                Set bestRHS = this.RHSColumns(j)
            End If
        Next j
    Next i
    
    Set LHSKeyColumn = bestLHS
    Set RHSKeyColumn = bestRHS
    
    this.IsAutoGuessed = True
End Sub
