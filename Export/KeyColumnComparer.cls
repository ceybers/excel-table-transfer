VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyColumnComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UnassignedVariableUsage
'@PredeclaredId
'@Folder("KeyColumn")
Option Explicit

Private Type TClsCompareKeyColumns
    lhs As KeyColumn
    rhs As KeyColumn
    LeftOnly As Collection
    Intersection As Collection
    RightOnly As Collection
    Dict As Object
End Type

Private this As TClsCompareKeyColumns

Public Property Get Score() As Double
    Dim total As Long
    total = this.LeftOnly.Count + this.Intersection.Count + this.RightOnly.Count
    
    Score = 0
    
    If total < 5 Then Exit Property
    
    If total > 0 Then
        Score = this.Intersection.Count / total
    End If
    
    'Debug.Print "Score: "; score
End Property

Public Property Get lhs() As KeyColumn
    Set lhs = this.lhs
End Property

Public Property Set lhs(ByVal rhs As KeyColumn)
    Set this.lhs = rhs
    TryDoWork
End Property

Public Property Get rhs() As KeyColumn
    Set rhs = this.rhs
End Property

Public Property Set rhs(ByVal rhs As KeyColumn)
    Set this.rhs = rhs
    TryDoWork
End Property

Public Property Get LeftOnly() As Collection
    Set LeftOnly = this.LeftOnly
End Property

Public Property Get Intersection() As Collection
    Set Intersection = this.Intersection
End Property

Public Property Get RightOnly() As Collection
    Set RightOnly = this.RightOnly
End Property

Public Property Get IsSubsetLHS() As Boolean
    IsSubsetLHS = this.LeftOnly.Count = 0
End Property

Public Property Get IsSubsetRHS() As Boolean
    IsSubsetRHS = this.RightOnly.Count = 0
End Property

Public Property Get IsMatch() As Boolean
    IsMatch = IsSubsetLHS And IsSubsetRHS
End Property

' Methods
Private Sub TryDoWork()
    If this.lhs Is Nothing Then Exit Sub
    If this.rhs Is Nothing Then Exit Sub
    DoWork
End Sub

Private Sub DoWork()
    Set this.Dict = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    For Each key In lhs.Keys
        Upsert key, 1
    Next
    For Each key In rhs.Keys
        Upsert key, 2
    Next
    
    Set this.Intersection = New Collection
    Set this.LeftOnly = New Collection
    Set this.RightOnly = New Collection
    
    Dim val As Integer
    
    For Each key In this.Dict.Keys
        Select Case this.Dict.Item(key)
            Case 1
                this.LeftOnly.Add key
            Case 2
                this.RightOnly.Add key
            Case 3
                this.Intersection.Add key
        End Select
    Next key
End Sub

Private Sub Upsert(ByVal key As String, ByVal value As Integer)
    If this.Dict.Exists(key) Then
        this.Dict.Item(key) = this.Dict.Item(key) + value
    Else
        this.Dict.Item(key) = value
    End If
End Sub

Public Function Create(ByVal lhs As KeyColumn, ByVal rhs As KeyColumn) As KeyColumnComparer
    Set Create = New KeyColumnComparer
    Set Create.lhs = lhs
    Set Create.rhs = rhs
End Function

Public Function Map(Optional ByVal FilterSource As Boolean = False, Optional ByVal FilterDestination As Boolean = False) As Variant
    'Debug.Assert Not this.LHS Is Nothing
    'Debug.Assert Not this.rhs Is Nothing
    If this.lhs Is Nothing Or this.rhs Is Nothing Then
        'Err.Raise vbObjectError + 4, Err.Source, "Cannot Map KeyColumns with an empty table"
        Exit Function
    End If
    
    Dim result As Variant
    Dim arr As Variant
    Dim i As Long
    
    Dim lhsFlt As KeyColumn
    Set lhsFlt = this.lhs
    If FilterSource Then
        Set lhsFlt = KeyColumn.FromRange(this.lhs.Range, this.lhs.IncludeNumbersAsText, True)
    End If
    
    Dim rhsFlt As KeyColumn
    Set rhsFlt = this.rhs
    If FilterDestination Then
        Set rhsFlt = KeyColumn.FromRange(this.rhs.Range, this.rhs.IncludeNumbersAsText, True)
    End If
    
    arr = rhsFlt.GetArray
    ReDim result(LBound(arr) To UBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        result(i) = lhsFlt.Find(arr(i))
    Next i
    
    Map = result
End Function
